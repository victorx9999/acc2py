"""Escreve o subpacote ``<pacote>/moderno/`` a partir do GeradorModerno."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Optional

from .moderno import DIFERENCAS, GeradorModerno, ident
from .sql.parser import parse


def gerar_moderno(pkg_dir: Path, pkg_name: str, sql: dict, sequencia: list, entradas: dict,
                  udfs: dict, pedidas: list, locale: str = "pt-BR") -> dict:
    """Gera o pacote moderno. Devolve {consulta: "moderna" | "fiel: motivo"}."""
    esquemas = {t: (e.get("colunas") or {}) for t, e in entradas.items()}
    gen = GeradorModerno(sql, esquemas, udfs, locale)
    alvo = list(sequencia) if sequencia else list(pedidas)
    for nome in alvo:  # na ordem: tabelas criadas por uma etapa ficam conhecidas pelas seguintes
        gen.info(nome)
    usadas = _fechamento(gen, alvo)
    situacao = {n: ("moderna" if gen.infos[n].moderna else f"fiel: {gen.infos[n].motivo}") for n in usadas}
    out = pkg_dir / "moderno"
    out.mkdir(parents=True, exist_ok=True)
    (out / "etapas.py").write_text(_etapas(gen, usadas, pkg_name), encoding="utf-8")
    (out / "entradas.py").write_text(_entradas(entradas), encoding="utf-8")
    (out / "__init__.py").write_text(_init(gen, usadas, alvo, bool(sequencia), pkg_name), encoding="utf-8")
    (out / "__main__.py").write_text(_main(pkg_name), encoding="utf-8")
    return situacao


def _fechamento(gen: GeradorModerno, alvo: list) -> list:
    """Consultas usadas (as pedidas + consultas de seleção que elas leem), na ordem."""
    out = []

    def visit(n):
        if n in out:
            return
        info = gen.info(n)
        for f in info.fontes:
            if f.lower() in gen.lower and f.lower() not in gen.criadas:
                visit(gen.lower[f.lower()])
        if not info.moderna:  # a versão fiel resolve as próprias dependências
            pass
        out.append(n)
    for n in alvo:
        visit(n)
    return out


def _etapas(gen: GeradorModerno, usadas: list, pkg: str) -> str:
    mods = sorted(gen.modulos_usados)
    head = [
        f'"""Etapas da rotina {pkg} em pandas idiomático (gerado por acc2py em {datetime.now():%d/%m/%Y}).',
        "",
        "Cada função corresponde a uma consulta do Access e recebe as tabelas que lê",
        "(DataFrames) e os parâmetros que usa. Funções de seleção devolvem o resultado;",
        "funções de ação (criar tabela, atualização, exclusão, acréscimo) devolvem a",
        "tabela nova/alterada, sem modificar a original.",
        "",
        "Diferenças em relação ao Access (verifique com comparar_com_fiel):",
    ] + [f"* {d}" for d in DIFERENCAS] + ['"""']
    imports = []
    if "math" in gen.funcs_usadas:
        imports.append("import math")
    if gen.usa_np:
        imports.append("import numpy as np")
    imports.append("import pandas as pd")
    if mods:
        imports.append("")
        imports.append(f"from ..modulos import {', '.join(mods)}  # funções VBA usadas nas consultas")
    body = []
    for n in usadas:
        info = gen.infos[n]
        if info.moderna:
            body.append(info.codigo)
        else:
            body.append(f"# {n}: NÃO modernizada ({info.motivo}).\n"
                        f"# Esta etapa roda pela versão fiel (consultas.py) dentro de executar().")
    return "\n".join(head + [""] + imports + ["", ""] + ["\n\n\n".join(body)]) + "\n"


def _tipo_param(gen: GeradorModerno, nome: str):
    k = gen.param_kind(nome)
    return {"date": "Optional[pd.Timestamp]", "num": "Optional[float]", "str": "Optional[str]"}.get(k, "Any")


def _init(gen: GeradorModerno, usadas: list, alvo: list, em_sequencia: bool, pkg: str) -> str:
    params = []
    for n in usadas:
        for p in (gen.infos[n].params or []):
            if p.lower() not in [x.lower() for x in params]:
                params.append(p)
    # parâmetros declarados em consultas que ficaram na versão fiel também entram
    for n in usadas:
        if not gen.infos[n].moderna:
            try:
                for p, _ in parse(gen.sql[n]).parameters:
                    if p.lower() not in [x.lower() for x in params]:
                        params.append(p)
            except Exception:  # noqa: BLE001
                pass
    fiel = [n for n in usadas if not gen.infos[n].moderna]
    campos = "\n".join(f"    {ident(p)}: {_tipo_param(gen, p)} = None  # [{p}]" for p in params) or "    pass"
    nomes = ", ".join(f"{ident(p)!r}: {p!r}" for p in params)

    def chamada(nome: str) -> str:
        info = gen.infos[nome]
        if not info.moderna:
            return f"_consulta_fiel(t, {nome!r}, p)"
        args = []
        for f in info.fontes:
            low = f.lower()
            if low in gen.lower and low not in gen.criadas and gen.infos.get(gen.lower[low]) \
                    and gen.infos[gen.lower[low]].tipo == "seleção":
                args.append(chamada(gen.lower[low]))
            else:
                args.append(f"t[{f!r}]")
        args += [f"p.{ident(x)}" for x in info.params]
        return f"etapas.{info.func}({', '.join(args)})"

    passos = []
    for i, n in enumerate(alvo, 1):
        info = gen.infos[n]
        passos.append(f"    # {i}. {n} ({info.tipo}{'' if info.moderna else ' — versão fiel'})")
        if not info.moderna:
            passos.append(f"    r = _etapa_fiel(t, {n!r}, p)")
            if info.tipo == "seleção":
                passos.append(f"    resultados[{n!r}] = r")
        elif info.tipo == "seleção":
            passos.append(f"    resultados[{n!r}] = {chamada(n)}")
        else:
            passos.append(f"    t[{info.alvo!r}] = {chamada(n)}")
    corpo = "\n".join(passos)
    doc_seq = "na ordem da rotina" if em_sequencia else "(consultas independentes)"
    lista_fiel = "".join(f"\n    * {n}: {gen.infos[n].motivo}" for n in fiel) or "\n    (nenhuma: tudo foi modernizado)"
    return f'''"""Rotina {pkg} em pandas idiomático (gerado por acc2py).

Uso:
    from {pkg}.moderno import Parametros, executar, ler_entradas
    entradas = ler_entradas({{"Tabela": "C:/dados/arquivo.xlsx"}})
    resultado = executar(entradas, Parametros(...))

Consultas que ficaram na versão fiel (rodam pelo runtime do acc2py):{lista_fiel}
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Optional

import pandas as pd

from . import etapas
from .entradas import ESQUEMAS, ler_entradas, ler_tabela  # noqa: F401

ETAPAS = {alvo!r}
VERSAO_FIEL = {fiel!r}
_NOMES_ACCESS = {{{nomes}}}


@dataclass
class Parametros:
    """Parâmetros das consultas (entre colchetes, o nome no Access)."""
{campos}

    def no_access(self) -> dict:
        """Mesmos valores com os nomes do Access (para a versão fiel)."""
        return {{_NOMES_ACCESS[k]: v for k, v in asdict(self).items() if v is not None}}


def executar(entradas: dict, parametros: Optional[Parametros] = None) -> Optional[pd.DataFrame]:
    """Roda as etapas {doc_seq} e devolve o resultado da última consulta de seleção."""
    resultados, _ = executar_tudo(entradas, parametros)
    return list(resultados.values())[-1] if resultados else None


def executar_tudo(entradas: dict, parametros: Optional[Parametros] = None):
    """Devolve (resultados de cada consulta de seleção, tabelas ao final)."""
    p = parametros or Parametros()
    faltando = [nome for campo, nome in _NOMES_ACCESS.items() if getattr(p, campo) is None]
    if faltando:
        raise ValueError(f"Parâmetros não informados: {{', '.join(faltando)}}")
    t = dict(entradas)
    resultados = {{}}
{corpo}
    return resultados, t


# ---------------------------------------------------------------- versão fiel
def _banco(t):
    from .. import CONSULTAS, FUNCOES  # carrega o runtime que veio no pacote
    from acc2py.runtime import Database
    from acc2py.runtime import vba
    db = Database(dict(t), CONSULTAS, FUNCOES)
    vba.set_database(db)
    return db


def _etapa_fiel(t, nome, p):
    db = _banco(t)
    r = db.run(nome, p.no_access())
    t.clear()
    t.update({{db.tables._names[k]: v for k, v in db.tables.items()}})
    return r


def _consulta_fiel(t, nome, p):
    return _banco(t).run(nome, p.no_access())


def comparar_com_fiel(entradas: dict, parametros: Optional[Parametros] = None) -> pd.DataFrame:
    """Roda esta versão e a fiel (consultas.py) com os mesmos dados e compara cada resultado."""
    from .. import CONSULTAS, FUNCOES  # noqa: F401
    from acc2py.validacao import comparar
    p = parametros or Parametros()
    mod, tab_mod = executar_tudo({{k: v.copy() for k, v in entradas.items()}}, p)
    db = _banco({{k: v.copy() for k, v in entradas.items()}})
    linhas = []
    for nome in ETAPAS:
        r = db.run(nome, p.no_access())
        if isinstance(r, pd.DataFrame):
            c = comparar(r, mod[nome])
            linhas.append({{"Consulta": nome, "Linhas fiel": c["linhas_access"], "Linhas moderno": c["linhas_python"],
                           "Resultado": "IGUAL" if c["iguais"] else "DIFERENTE",
                           "Detalhe": "" if c["iguais"] else _detalhe(c)}})
    return pd.DataFrame(linhas)


def _detalhe(c) -> str:
    partes = []
    if c["linhas_so_access"] or c["linhas_so_python"]:
        partes.append(f"{{c['linhas_so_access']}} linha(s) só na fiel, {{c['linhas_so_python']}} só na moderna")
    if c["colunas_so_access"]:
        partes.append(f"colunas só na fiel: {{', '.join(map(str, c['colunas_so_access']))}}")
    if c["colunas_so_python"]:
        partes.append(f"colunas só na moderna: {{', '.join(map(str, c['colunas_so_python']))}}")
    return "; ".join(partes)
'''


def _entradas(entradas: dict) -> str:
    esquemas = {t: (e.get("colunas") or {}) for t, e in entradas.items()}
    fixas = [t for t, e in entradas.items() if e.get("modo") == "fixa"]
    esq_txt = json.dumps(esquemas, ensure_ascii=False, indent=4)
    return f'''"""Leitura das tabelas de entrada (Excel, CSV, TXT) só com pandas.

ESQUEMAS traz as colunas e os tipos que as tabelas tinham no Access. As tabelas
fixas ({", ".join(fixas) or "nenhuma"}) vêm da cópia em ../dados se não forem informadas.
"""
from __future__ import annotations

import unicodedata
from pathlib import Path

import pandas as pd

ESQUEMAS = {esq_txt}
FIXAS = {fixas!r}
PASTA_DADOS = Path(__file__).resolve().parent.parent / "dados"


def _chave(nome: str) -> str:
    s = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode()
    return "".join(ch for ch in s.lower() if ch.isalnum())


def _numero(serie: pd.Series) -> pd.Series:
    """Converte '1.234,56' (ou '1234.56') em número."""
    if pd.api.types.is_numeric_dtype(serie):
        return serie
    texto = serie.astype("string").str.strip()
    if texto.str.contains(r",\\d+$", regex=True, na=False).any():
        texto = texto.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(texto, errors="coerce")


def ler_tabela(caminho, tabela: str, aba=0, sep=None, encoding=None, linha_cabecalho: int = 1) -> pd.DataFrame:
    """Lê um arquivo e ajusta nomes e tipos de coluna ao esquema da tabela."""
    caminho = Path(caminho)
    esquema = ESQUEMAS.get(tabela, {{}})
    textos = [c for c, t in esquema.items() if t == "texto"]
    if caminho.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(caminho, sheet_name=aba, header=linha_cabecalho - 1)
    else:
        if encoding is None:
            try:
                caminho.read_bytes().decode("utf-8")
                encoding = "utf-8-sig"
            except UnicodeDecodeError:
                encoding = "cp1252"
        df = pd.read_csv(caminho, sep=sep, engine="python", encoding=encoding,
                         header=linha_cabecalho - 1, dtype=str, keep_default_na=False, na_values=[""])
    df = df.dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    por_chave = {{_chave(c): c for c in esquema}}
    df = df.rename(columns={{c: por_chave[_chave(c)] for c in df.columns if _chave(c) in por_chave}})
    faltando = [c for c in esquema if c not in df.columns]
    if faltando:
        raise ValueError(f"{{tabela}}: faltam as colunas {{faltando}} em {{caminho.name}}")
    for col, tipo in esquema.items():
        if tipo in ("numero", "inteiro"):
            df[col] = _numero(df[col])
        elif tipo == "data":
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
        elif tipo == "logico":
            df[col] = df[col].astype("string").str.strip().str.lower().map(
                {{"sim": True, "true": True, "verdadeiro": True, "1": True, "-1": True,
                  "não": False, "nao": False, "false": False, "falso": False, "0": False}})
        elif col in textos:
            df[col] = df[col].astype("string")
    return df.reset_index(drop=True)


def ler_entradas(arquivos: dict) -> dict:
    """{{"Tabela": caminho}} -> {{"Tabela": DataFrame}}. Tabelas fixas não informadas vêm de ../dados."""
    tabelas = {{}}
    for tabela in ESQUEMAS:
        if tabela in arquivos:
            tabelas[tabela] = ler_tabela(arquivos[tabela], tabela)
            continue
        copia = next((p for p in (PASTA_DADOS / f"{{tabela}}{{ext}}" for ext in (".parquet", ".pkl", ".csv"))
                      if p.exists()), None)
        if copia is None:
            raise ValueError(f"falta o arquivo da tabela de entrada {{tabela!r}}")
        if tabela not in FIXAS:
            print(f"Aviso: {{tabela}} é mensal mas não foi informada; usando a cópia da migração ({{copia.name}})")
        tabelas[tabela] = (pd.read_parquet(copia) if copia.suffix == ".parquet"
                           else pd.read_pickle(copia) if copia.suffix == ".pkl" else pd.read_csv(copia))
    return tabelas
'''


def _main(pkg: str) -> str:
    return f'''"""Linha de comando da versão moderna.

    python -m {pkg}.moderno --entrada "Pedidos=C:/dados/pedidos.csv" --param "Data inicial=01/01/2026" -o saida.xlsx
    python -m {pkg}.moderno --comparar          # compara com a versão fiel usando as cópias da migração
"""
import argparse
import sys
from dataclasses import fields

import pandas as pd

from . import Parametros, comparar_com_fiel, executar, ler_entradas


def _parametros(itens):
    tipos = {{f.name: str(f.type) for f in fields(Parametros)}}
    nomes = {{v.lower(): k for k, v in __import__(__package__, fromlist=["_NOMES_ACCESS"])._NOMES_ACCESS.items()}}
    valores = {{}}
    for item in itens:
        nome, valor = item.split("=", 1)
        campo = nomes.get(nome.strip().lower(), nome.strip())
        tipo = tipos.get(campo, "")
        if "Timestamp" in tipo:
            valores[campo] = pd.to_datetime(valor, dayfirst=True)
        elif "float" in tipo or valor.replace(",", "").replace(".", "").lstrip("-").isdigit():
            valores[campo] = float(valor.replace(".", "").replace(",", ".") if "," in valor else valor)
        else:
            valores[campo] = valor
    return Parametros(**valores)


def main():
    ap = argparse.ArgumentParser(prog="python -m {pkg}.moderno")
    ap.add_argument("--entrada", action="append", default=[], metavar="TABELA=ARQUIVO")
    ap.add_argument("--param", action="append", default=[], metavar="NOME=VALOR")
    ap.add_argument("-o", "--saida", default="saida_moderna.xlsx")
    ap.add_argument("--comparar", action="store_true", help="compara com a versão fiel em vez de gravar a saída")
    a = ap.parse_args()
    entradas = ler_entradas(dict(x.split("=", 1) for x in a.entrada))
    p = _parametros(a.param)
    if a.comparar:
        r = comparar_com_fiel(entradas, p)
        print(r.to_string(index=False))
        return 0 if (r["Resultado"] == "IGUAL").all() else 1
    resultado = executar(entradas, p)
    resultado.to_excel(a.saida, index=False)
    print(f"Saída gravada em {{a.saida}} ({{len(resultado)}} linhas)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
'''
