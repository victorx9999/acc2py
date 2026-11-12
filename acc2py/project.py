"""Orquestra a transpilação de um banco Access inteiro em um pacote Python."""
from __future__ import annotations

import difflib
import json
import fnmatch
import os
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

from .diagnostics import Diagnostics
from .extract import ExtractorError, ModuleSource, open_extractor
from .sql.parser import parse as parse_sql
from .sql.to_pandas import SQLCompiler, py_identifier
from .vba.parser import parse_module
from .vba.to_python import Project, ModuleGenerator, py_name


@dataclass
class Result:
    package_dir: Path
    queries: dict = field(default_factory=dict)      # nome -> (função, status)
    modules: dict = field(default_factory=dict)      # nome -> (arquivo, status)
    added_deps: list = field(default_factory=list)
    sequence: list = field(default_factory=list)
    inputs: dict = field(default_factory=dict)
    modern: dict = field(default_factory=dict)
    diags: Diagnostics = field(default_factory=Diagnostics)


def _natural(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def _match_names(wanted: list[str], available: list[str], what: str) -> list[str]:
    lower = {a.lower(): a for a in available}
    out, missing = [], []
    for w in wanted:
        w2 = w.strip().strip("[]")
        if any(c in w2 for c in "*?"):  # curinga: "Vendas*", "rel_??_*"
            hits = [a for a in available if fnmatch.fnmatchcase(a.lower(), w2.lower())]
            if hits:
                out.extend(h for h in sorted(hits, key=_natural) if h not in out)
            else:
                missing.append(w2)
        elif w2.lower() in lower:
            if lower[w2.lower()] not in out:
                out.append(lower[w2.lower()])
        else:
            missing.append(w2)
    if missing:
        msgs = []
        for m in missing:
            close = difflib.get_close_matches(m, available, n=3, cutoff=0.5)
            msgs.append(f"  - {m!r}" + (f" (você quis dizer: {', '.join(close)}?)" if close else ""))
        raise SystemExit(f"{what} não encontrada(s) no banco:\n" + "\n".join(msgs)
                         + f"\nDisponíveis: {', '.join(sorted(available)) or '(nenhuma)'}")
    return out


def transpile_database(source: str, out_dir: str, *, queries: Optional[list[str]] = None,
                       modules: Optional[list[str]] = None, no_vba: bool = False,
                       backend: str = "auto", export_data: bool = False, locale: str = "pt-BR",
                       include_hidden: bool = False, package: Optional[str] = None,
                       sequence: Optional[list[str]] = None,
                       input_modes: Optional[dict] = None, all_monthly: bool = False,
                       vendor: bool = True, modern: bool = False) -> Result:
    src_path = Path(source)
    pkg_name = package or py_identifier(src_path.stem if src_path.suffix else src_path.name) or "banco"
    pkg_dir = Path(out_dir) / pkg_name
    res = Result(pkg_dir)
    diags = res.diags
    ex = open_extractor(source, backend)
    try:
        all_queries = ex.queries(include_hidden)
        try:
            table_names = ex.table_names()
        except Exception as e:  # noqa: BLE001
            table_names = []
            diags.warn("extração", f"não foi possível listar as tabelas: {e}")

        # ------------------------------------------------ consultas e dependências
        steps: list[str] = []
        if sequence:
            steps = _match_names(sequence, list(all_queries), "Consulta(s) da sequência")
            res.sequence = steps
            queries = list(queries or []) + steps
        if queries:
            selected = _match_names(queries, list(all_queries), "Consulta(s)")
        else:
            selected = list(all_queries)
        # tabelas criadas por consultas "criar tabela" (SELECT … INTO): existem só depois que rodam
        created = {t.lower() for t in (_make_table_target(all_queries[q]) for q in selected) if t}
        compiled: dict[str, str] = {}
        func_names: dict[str, str] = {}
        used_udfs: set[str] = set()
        table_deps: dict[str, set] = {}
        query_lower = {q.lower(): q for q in all_queries}
        pending = list(selected)
        seen = set()
        while pending:
            name = pending.pop(0)
            if name.lower() in seen:
                continue
            seen.add(name.lower())
            fn = py_identifier(name, "q_")
            base, k = fn, 2
            while fn in func_names.values():
                fn = f"{base}_{k}"
                k += 1
            func_names[name] = fn
            sql = all_queries[name]
            comp = SQLCompiler(diags, locale="pt" if locale.startswith("pt") else "en")
            try:
                if sql.startswith("-- ERRO"):
                    raise ValueError(sql[3:])
                compiled[name] = comp.compile_function(name, sql, fn)
                res.queries[name] = (fn, "ok")
            except Exception as e:  # noqa: BLE001
                diags.error(f"consulta '{name}'", f"não convertida: {e}")
                compiled[name] = _stub_query(name, sql, fn, e)
                res.queries[name] = (fn, "erro")
                continue
            table_deps[name] = {d for d in comp.deps if d.lower() not in query_lower}
            for d in comp.deps:
                if d.lower() in query_lower and d.lower() not in seen:
                    dep = query_lower[d.lower()]
                    pending.append(dep)
                    if queries:
                        res.added_deps.append((dep, name))
                elif d.lower() not in query_lower and table_names and d.lower() not in created \
                        and d.lower() not in {t.lower() for t in table_names}:
                    diags.warn(f"consulta '{name}'", f"fonte '{d}' não é tabela nem consulta deste banco")
            for item in diags.items:
                m = re.match(r"função '(.+)' não é nativa", item.message)
                if m and item.where == f"consulta '{name}'":
                    used_udfs.add(m.group(1).lower())

        # ------------------------------------------------ VBA
        mod_sources: list[ModuleSource] = []
        if not no_vba and (modules or not queries or used_udfs):
            try:
                mod_sources = ex.modules()
            except ExtractorError as e:
                diags.warn("extração", f"VBA não extraído: {e}")
            except Exception as e:  # noqa: BLE001
                diags.warn("extração", f"VBA não extraído: {e}")
        parsed = []
        for ms in mod_sources:
            try:
                parsed.append((ms, parse_module(ms.code, ms.name, ms.kind)))
            except Exception as e:  # noqa: BLE001
                diags.error(f"módulo {ms.name}", f"erro de sintaxe, módulo não convertido: {e}")
                res.modules[ms.name] = (None, "erro")
        if modules and [m.lower() for m in modules] != ["todos"]:
            keep = set(n.lower() for n in _match_names(modules, [m.name for _, m in parsed], "Módulo(s)"))
        elif queries and not modules:
            keep = {m.name.lower() for _, m in parsed
                    if any(p.public and p.name.lower() in used_udfs for p in m.procedures)}
        else:
            keep = {m.name.lower() for _, m in parsed}
        for udf in used_udfs:
            if not any(p.name.lower() == udf for _, m in parsed for p in m.procedures):
                diags.error("consultas", f"função '{udf}' usada em consulta não foi encontrada no VBA")
        project = Project([m for _, m in parsed])  # símbolos de todos, para resolver chamadas
        generated: dict[str, str] = {}
        public_funcs: dict[str, str] = {}
        for ms, mod in parsed:
            if mod.name.lower() not in keep:
                continue
            try:
                generated[mod.name] = ModuleGenerator(mod, project, diags).generate()
                res.modules[mod.name] = (f"modulos/{py_name(mod.name)}.py", "ok")
                if mod.kind == "standard":
                    for p in mod.procedures:
                        if p.public and p.kind == "function":
                            public_funcs[p.name] = f"{py_name(mod.name)}.{py_name(p.name)}"
            except Exception as e:  # noqa: BLE001
                diags.error(f"módulo {mod.name}", f"falha na geração: {e}")
                res.modules[mod.name] = (None, "erro")
        # módulos importados por outros (from . import X) precisam existir
        needed = set()
        for code in generated.values():
            needed |= set(re.findall(r"^from \. import (\w+)$", code, re.M))
        for ms, mod in parsed:
            if py_name(mod.name) in needed and mod.name not in generated:
                generated[mod.name] = ModuleGenerator(mod, project, diags).generate()
                res.modules[mod.name] = (f"modulos/{py_name(mod.name)}.py", "ok (dependência)")

        # ------------------------------------------------ escrita
        if pkg_dir.exists():
            shutil.rmtree(pkg_dir)
        (pkg_dir / "sql_original").mkdir(parents=True)
        kinds = {n: _query_kind(all_queries[n]) for n in compiled}
        requested = [n for n in (selected if queries else compiled) if kinds.get(n) == "seleção"]
        _write(pkg_dir / "consultas.py", _consultas_file(compiled, func_names, all_queries, source,
                                                         kinds, requested, steps))
        for name in compiled:
            _write(pkg_dir / "sql_original" / f"{_safe_file(name)}.sql", all_queries[name] + "\n")
        if generated:
            (pkg_dir / "modulos").mkdir()
            (pkg_dir / "vba_original").mkdir()
            _write(pkg_dir / "modulos" / "__init__.py", '"""Módulos VBA transpilados."""\n')
            for name, code in generated.items():
                _write(pkg_dir / "modulos" / f"{py_name(name)}.py", code)
            for ms, mod in parsed:
                if mod.name in generated:
                    ext = ".bas" if ms.kind == "standard" else ".cls"
                    _write(pkg_dir / "vba_original" / f"{_safe_file(ms.name)}{ext}", ms.code)
        autonumber = {}
        try:
            autonumber = ex.autonumber_columns()
        except Exception:  # noqa: BLE001
            pass
        # ------------------------------------------------ tabelas de entrada
        created_all = {t.lower() for t in (_make_table_target(all_queries[q]) for q in compiled) if t}
        real_tables = {t.lower(): t for t in table_names}
        try:
            linked = {k.lower(): v for k, v in ex.linked_tables().items()}
        except Exception:  # noqa: BLE001
            linked = {}
        modes = {k.lower(): v for k, v in (input_modes or {}).items()}
        entradas: dict[str, dict] = {}
        for qname in compiled:
            for d in sorted(table_deps.get(qname, ())):
                key = d.lower()
                if key in created_all or any(e.lower() == key for e in entradas):
                    continue
                if real_tables and key not in real_tables:
                    continue
                tname = real_tables.get(key, d)
                default = "mensal" if (all_monthly or key in linked) else "fixa"
                entradas[tname] = {"modo": modes.get(key, default), "vinculada": linked.get(key, ""),
                                   "colunas": {}, "usada_por": []}
                try:
                    entradas[tname]["colunas"] = ex.columns(tname)
                except Exception as e:  # noqa: BLE001
                    diags.warn(f"tabela {tname}", f"colunas/tipos não lidos (a conferência dos arquivos "
                                                  f"de entrada fica limitada): {e}")
        for qname in compiled:
            for d in table_deps.get(qname, ()):
                for t in entradas:
                    if t.lower() == d.lower() and qname not in entradas[t]["usada_por"]:
                        entradas[t]["usada_por"].append(qname)
        res.inputs = entradas
        _write(pkg_dir / "entradas.json", json.dumps(entradas, ensure_ascii=False, indent=2))
        data_dir = pkg_dir / "dados"
        data_dir.mkdir()
        to_export = list(table_names) if export_data else list(entradas)
        for t in to_export:
            try:
                df = ex.read_table(t)
                _save_table(df, data_dir, t)
                if t in entradas and not entradas[t]["colunas"]:
                    from .runtime.inputs import schema_from_frame
                    entradas[t]["colunas"] = schema_from_frame(df)
            except Exception as e:  # noqa: BLE001
                lvl = diags.error if entradas.get(t, {}).get("modo") == "fixa" else diags.warn
                lvl(f"tabela {t}", f"cópia não exportada para dados/: {e}")
        _write(pkg_dir / "entradas.json", json.dumps(entradas, ensure_ascii=False, indent=2))
        if vendor:
            _vendor_runtime(pkg_dir)
        _write(pkg_dir / ".gitignore", "# dados do banco e resultados não vão para o controle de versão\n"
                                       "dados/\nentradas/\nsaidas/\nexecucao.json\n__pycache__/\n")
        _write(pkg_dir / "__main__.py", _main_file(pkg_name))
        _write(Path(out_dir) / f"Abrir tela - {pkg_name}.bat",
               f'@echo off\r\ncd /d "%~dp0"\r\npython -m {pkg_name}\r\nif errorlevel 1 pause\r\n')
        _write(pkg_dir / "__init__.py", _init_file(pkg_name, source, generated, public_funcs,
                                                   autonumber, locale, export_data))
        if modern:
            from .moderno_pacote import gerar_moderno
            ok_sql = {n: all_queries[n] for n, (_, st_) in res.queries.items() if st_ == "ok"}
            alvo_seq = [n for n in steps if n in ok_sql]
            pedidas = [n for n in requested if n in ok_sql]
            try:
                res.modern = gerar_moderno(pkg_dir, pkg_name, ok_sql, alvo_seq, entradas, public_funcs,
                                           pedidas, locale)
                for n, sit in res.modern.items():
                    if sit != "moderna":
                        diags.info(f"moderno: {n}", f"ficou na versão fiel — {sit[6:]}")
            except Exception as e:  # noqa: BLE001
                diags.error("moderno", f"versão moderna não gerada: {e}")
        _write(Path(out_dir) / f"exemplo_{pkg_name}.py", _example_file(pkg_name, compiled, all_queries))
        _write(pkg_dir / "RELATORIO.md", _report(res, source, pkg_name, table_names))
    finally:
        try:
            ex.close()
        except Exception:  # noqa: BLE001
            pass
    return res


# ======================================================================= arquivos
def _write(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _safe_file(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def _save_table(df, folder: Path, name: str):
    try:
        import pyarrow  # noqa: F401
        df.to_parquet(folder / f"{_safe_file(name)}.parquet", index=False)
    except ImportError:
        df.to_pickle(folder / f"{_safe_file(name)}.pkl")


def _stub_query(name, sql, fn, err) -> str:
    safe = sql.replace("\\", "\\\\").replace('"""', "'''")
    body = "\n".join("    " + ln for ln in safe.splitlines())
    msg = f"Consulta '{name}' não foi convertida: {err}".replace('"', "'")
    return (f'def {fn}(db, p=None):\n    """ATENÇÃO: consulta não convertida automaticamente.\n\n'
            f'    SQL original:\n{body}\n    """\n'
            f'    # Alternativa sem converter: return db.sql(SQL_ORIGINAL[{name!r}], p)\n'
            f'    raise NotImplementedError("{msg}")')


def _query_kind(sql: str) -> str:
    from .sql import ast as SA
    try:
        st = parse_sql(sql)
    except Exception:  # noqa: BLE001
        return "desconhecido"
    if isinstance(st, SA.Select) and st.into:
        return "ação"
    return "seleção" if isinstance(st, (SA.Select, SA.Union, SA.Transform)) else "ação"


def _make_table_target(sql: str) -> Optional[str]:
    from .sql import ast as SA
    try:
        st = parse_sql(sql)
    except Exception:  # noqa: BLE001
        return None
    return st.into if isinstance(st, SA.Select) else None


def _consultas_file(compiled, func_names, all_queries, source, kinds, requested, steps) -> str:
    lines = [f'"""Consultas do Access ({Path(source).name}) transpiladas para pandas por acc2py.',
             "",
             "Cada função recebe o banco (acc2py.runtime.Database) e um dicionário de",
             "parâmetros, e devolve um DataFrame (seleção) ou o nº de linhas afetadas (ação).",
             '"""',
             "import numpy as np  # noqa: F401",
             "import pandas as pd  # noqa: F401",
             "",
             "from acc2py.runtime import core as rt",
             "from acc2py.runtime import sqlfuncs as F  # noqa: F401",
             "", ""]
    for name in compiled:
        lines.append(compiled[name])
        lines += ["", ""]
    lines.append("# Nome no Access -> função")
    lines.append("CONSULTAS = {")
    for name in compiled:
        lines.append(f"    {name!r}: {func_names[name]},")
    lines.append("}")
    lines.append("")
    lines.append("# 'seleção' devolve DataFrame; 'ação' altera tabelas (UPDATE/DELETE/INSERT/SELECT INTO)")
    lines.append("TIPOS = {")
    for name in compiled:
        lines.append(f"    {name!r}: {kinds[name]!r},")
    lines.append("}")
    lines.append("")
    lines.append("# Consultas de seleção pedidas na transpilação (padrão de exportar())")
    lines.append("PEDIDAS = [")
    for name in requested:
        lines.append(f"    {name!r},")
    lines.append("]")
    lines.append("")
    lines.append("# Ordem de execução (rotina em etapas). Edite aqui para mudar a sequência.")
    lines.append("SEQUENCIA = [")
    for name in steps:
        lines.append(f"    {name!r},  # {kinds.get(name, '?')}")
    lines.append("]")
    lines.append("")
    lines.append("SQL_ORIGINAL = {")
    for name in compiled:
        lines.append(f"    {name!r}: {all_queries[name]!r},")
    lines.append("}")
    return "\n".join(lines) + "\n"


def _vendor_runtime(pkg_dir: Path):
    """Copia o acc2py para dentro do pacote: ele roda só com Python + pandas."""
    src = Path(__file__).resolve().parent
    dst = pkg_dir / "_vendor" / "acc2py"
    shutil.copytree(src, dst, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))


def _main_file(pkg) -> str:
    return f'''"""python -m {pkg}  → abre a tela de execução no navegador."""
import sys

from . import pacote

if __name__ == "__main__":
    from acc2py.tela import servir
    sys.exit(servir(pacote=pacote()))
'''


def _init_file(pkg, source, generated, public_funcs, autonumber, locale, has_data) -> str:
    mod_imports = ""
    if generated:
        names = ", ".join(sorted(py_name(n) for n in generated))
        mod_imports = f"from .modulos import {names}  # noqa: F401,E402\n"
    funcs = "".join(f"    {k!r}: {v},\n" for k, v in sorted(public_funcs.items()))
    return f'''"""Pacote gerado por acc2py a partir de {Path(source).name} em {datetime.now():%d/%m/%Y %H:%M}.

Roda SEM Microsoft Access: as tabelas de entrada vêm de arquivos Excel/CSV/TXT.

    python -m {pkg}                               # abre a tela (escolher arquivos e executar)

    from {pkg} import executar, definir_entrada
    definir_entrada("Movimento", "C:/dados/movimento_jan.xlsx")
    etapas, arquivo = executar({{"Data inicial": "01/01/2026"}})
"""
import sys as _sys
from pathlib import Path as _Path

# usa a cópia do acc2py que veio junto (mesma versão com que o pacote foi gerado)
_VENDOR = str(_Path(__file__).parent / "_vendor")
if _VENDOR not in _sys.path and (_Path(_VENDOR) / "acc2py").exists():
    _sys.path.insert(0, _VENDOR)

from acc2py.runtime import Database, vbafuncs  # noqa: E402
from acc2py.runtime import vba as vb  # noqa: E402
from acc2py.runtime.pacote import Pacote, convert_param  # noqa: E402

from .consultas import CONSULTAS, PEDIDAS, SEQUENCIA, SQL_ORIGINAL, TIPOS  # noqa: F401,E402
{mod_imports}
ORIGEM = {str(Path(source).resolve())!r}
PASTA = _Path(__file__).parent

# Funções públicas dos módulos VBA (podem ser chamadas dentro das consultas)
FUNCOES = {{
{funcs}}}

# Colunas de AutoNumeração (preenchidas em INSERT/AddNew)
AUTONUMERACAO = {autonumber!r}

vbafuncs.set_locale({locale!r})


def pacote() -> Pacote:
    return Pacote(PASTA, CONSULTAS, FUNCOES, SEQUENCIA, SQL_ORIGINAL, TIPOS, AUTONUMERACAO, {pkg!r})


def abrir(origem=None, backend="auto") -> Database:
    """Carrega as tabelas e devolve o banco com consultas e funções registradas.

    Sem ``origem``: lê os arquivos de entrada configurados (tela ou definir_entrada) e,
    para tabelas fixas, a cópia em dados/. Com ``origem``: um .accdb/.mdb (precisa do
    Access, útil na migração) ou uma pasta com um arquivo por tabela.
    """
    if origem is None:
        db = pacote().abrir()
    elif _Path(origem).is_dir():
        db = Database.from_folder(origem, CONSULTAS, FUNCOES, autonumber=AUTONUMERACAO)
    else:
        db = Database.from_access(origem, CONSULTAS, FUNCOES, backend=backend,
                                  autonumber=AUTONUMERACAO)
    vb.set_database(db)
    vb._QUERY_SQL.update({{k.lower(): v for k, v in SQL_ORIGINAL.items()}})
    return db


def definir_entrada(tabela, arquivo, **opcoes):
    """Escolhe o arquivo (Excel/CSV/TXT) de uma tabela de entrada; devolve a conferência.

    opcoes: aba="Plan1", separador=";", encoding="cp1252", linha_cabecalho=1
    """
    return pacote().definir_entrada(tabela, arquivo, **opcoes)


def executar(parametros=None, destino=None, intermediarias=False):
    """Roda a sequência com as entradas configuradas. Devolve (log das etapas, arquivo de saída)."""
    pk = pacote()
    tipos = {{p["nome"].lower(): p["tipo"] for p in pk.parametros_declarados()}}
    params = {{k: (convert_param(v, tipos.get(k.lower(), "auto")) if isinstance(v, str) else v)
              for k, v in (parametros or {{}}).items()}}
    db = abrir()
    return pk.executar(params, destino, intermediarias, db=db)


def exportar(destino="resultados.xlsx", consultas=None, parametros=None, db=None):
    """Roda várias consultas de seleção independentes e grava tudo numa saída só."""
    from acc2py.runtime.core import export_results
    db = db or abrir()
    return export_results(db, destino, list(consultas or PEDIDAS), parametros)


def executar_sequencia(destino="saida_final.xlsx", parametros=None, etapas=None,
                       intermediarias=False, db=None):
    """Roda as consultas EM ORDEM (SEQUENCIA) sobre o banco ``db`` (ou abrir())."""
    from acc2py.runtime.core import run_sequence
    db = db or abrir()
    return run_sequence(db, list(etapas or SEQUENCIA), parametros, destino, intermediarias)
'''


def _example_file(pkg, compiled, all_queries) -> str:
    lines = [f'"""Exemplo de uso do pacote {pkg} (gerado por acc2py)."""',
             f"from {pkg} import abrir, exportar", "", "db = abrir()", ""]
    for name in list(compiled)[:5]:
        try:
            st = parse_sql(all_queries[name])
            params = [n for n, _ in st.parameters]
        except Exception:  # noqa: BLE001
            params = []
        p = ", {" + ", ".join(f"{n!r}: None" for n in params) + "}" if params else ""
        lines.append(f"print(db.run({name!r}{p}))")
    lines += ["", "# Todas as consultas escolhidas numa saída só (uma aba por consulta):",
              "print(exportar(\"resultados.xlsx\", db=db))"]
    return "\n".join(lines) + "\n"


def _report(res: Result, source, pkg, tables) -> str:
    d = res.diags
    ok_q = sum(1 for _, s in res.queries.values() if s == "ok")
    ok_m = sum(1 for _, s in res.modules.values() if s.startswith("ok"))
    out = [f"# Relatório de transpilação — {Path(source).name}", "",
           f"Gerado em {datetime.now():%d/%m/%Y %H:%M} · pacote `{pkg}`", "",
           f"- Consultas: **{ok_q}/{len(res.queries)}** convertidas",
           f"- Módulos VBA: **{ok_m}/{len(res.modules)}** convertidos",
           f"- Tabelas no banco: {len(tables)}",
           f"- Avisos: {d.count('erro')} erro(s), {d.count('aviso')} aviso(s), {d.count('info')} informação(ões)",
           ""]
    if res.added_deps:
        out += ["## Consultas incluídas como dependência", "",
                "Você não pediu estas, mas as consultas escolhidas dependem delas:", ""]
        out += [f"- **{dep}** (usada por *{by}*)" for dep, by in res.added_deps]
        out.append("")
    out += ["## Consultas", "", "| Consulta no Access | Função Python | Situação |", "|---|---|---|"]
    out += [f"| {n} | `{fn}` | {'✅' if s == 'ok' else '❌ não convertida'} |" for n, (fn, s) in res.queries.items()]
    if res.modules:
        out += ["", "## Módulos VBA", "", "| Módulo | Arquivo | Situação |", "|---|---|---|"]
        out += [f"| {n} | `{f or '-'}` | {'✅ ' + s if s.startswith('ok') else '❌ ' + s} |"
                for n, (f, s) in res.modules.items()]
    if res.inputs:
        out += ["", "## Tabelas de entrada", "",
                "Sem o Access, estas tabelas vêm de arquivos. **Mensal**: você informa o arquivo a cada "
                "execução (tela ou `definir_entrada`). **Fixa**: usa a cópia tirada do Access em `dados/` "
                "(tabelas de referência/de-para). Mude o modo em `entradas.json` se precisar.", "",
                "| Tabela | Modo | Vinculada a | Colunas | Usada por |", "|---|---|---|---|---|"]
        for t, e in res.inputs.items():
            out.append(f"| {t} | {e['modo']} | {e.get('vinculada') or '-'} | {len(e.get('colunas') or {})} "
                       f"| {', '.join(e.get('usada_por', [])[:4])} |")
    if res.modern:
        nm = sum(1 for v in res.modern.values() if v == "moderna")
        out += ["", "## Versão moderna (`moderno/`)", "",
                f"**{nm}/{len(res.modern)}** consultas em pandas idiomático; as demais rodam pela versão fiel. "
                "Confira com `python -m " + pkg + ".moderno --comparar`.", "",
                "| Consulta | Situação |", "|---|---|"]
        out += [f"| {n} | {'✅ moderna' if v == 'moderna' else '↩️ ' + v} |" for n, v in res.modern.items()]
    out += ["", "## Avisos", "", d.to_markdown(),
            "## Como validar", "",
            "1. Rode cada consulta no Access e no Python com os mesmos dados e parâmetros.",
            "2. Compare com `pandas.testing.assert_frame_equal(df_access, df_python, check_dtype=False)`.",
            "3. Procure `ATENÇÃO` e `NotImplementedError` no código gerado.", ""]
    return "\n".join(out)
