"""Operação de um pacote gerado, sem Access: entradas, parâmetros e execução.

Arquivos dentro do pacote gerado:
* ``entradas.json``  — fixo, gravado na migração: tabelas de entrada, modo e colunas/tipos.
* ``execucao.json``  — muda a cada uso: arquivo escolhido para cada entrada, últimos parâmetros.
* ``dados/``         — cópia das tabelas tiradas do Access na migração (usada pelas tabelas "fixas").
* ``entradas/``      — arquivos enviados pela tela.
* ``saidas/``        — resultados das execuções.
"""
from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

from .core import Database
from .inputs import Checagem, read_input


class EntradaError(Exception):
    def __init__(self, checagens: list[Checagem]):
        self.checagens = checagens
        linhas = [f"  - {c.tabela}: {c.resumo()}" for c in checagens if not c.ok]
        super().__init__("Problemas nas tabelas de entrada:\n" + "\n".join(linhas))


def _read_json(path: Path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (FileNotFoundError, ValueError):
        return default


def _write_json(path: Path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


class Pacote:
    """Acesso às configurações de um pacote gerado."""

    def __init__(self, pkg_dir, consultas: dict, funcoes: dict, sequencia: list, sql_original: dict,
                 tipos: dict, autonumeracao: Optional[dict] = None, nome: str = ""):
        self.dir = Path(pkg_dir)
        self.nome = nome or self.dir.name
        self.consultas = consultas
        self.funcoes = funcoes
        self.sequencia = list(sequencia)
        self.sql_original = sql_original
        self.tipos = tipos
        self.autonumeracao = autonumeracao or {}

    # -------------------------------------------------- configuração
    @property
    def entradas(self) -> dict:
        return _read_json(self.dir / "entradas.json", {})

    def execucao(self) -> dict:
        cfg = _read_json(self.dir / "execucao.json", {})
        cfg.setdefault("entradas", {})
        cfg.setdefault("parametros", {})
        return cfg

    def salvar_execucao(self, cfg: dict):
        _write_json(self.dir / "execucao.json", cfg)

    def definir_entrada(self, tabela: str, arquivo: str, **opcoes) -> Checagem:
        """Escolhe o arquivo de uma tabela de entrada e confere a leitura."""
        cfg = self.execucao()
        atual = {k: v for k, v in opcoes.items() if v not in (None, "")}
        atual["arquivo"] = str(arquivo)
        _, chk = self.ler_entrada(tabela, atual)
        try:
            atual["modificado"] = os.path.getmtime(arquivo)
        except OSError:
            pass
        cfg["entradas"][tabela] = atual
        self.salvar_execucao(cfg)
        return chk

    def ler_entrada(self, tabela: str, cfg_arquivo: Optional[dict] = None):
        esquema = self.entradas.get(tabela, {})
        cfg = dict(cfg_arquivo or self.execucao()["entradas"].get(tabela, {}))
        cfg["colunas"] = esquema.get("colunas") or {}
        return read_input(tabela, cfg, self.dir)

    def situacao_entradas(self) -> list[dict]:
        """Para a tela: cada entrada, de onde virá e se o arquivo mudou desde a última execução."""
        exe = self.execucao()
        out = []
        for tabela, esq in self.entradas.items():
            cfg = exe["entradas"].get(tabela, {})
            arq = cfg.get("arquivo", "")
            existe = bool(arq) and Path(arq).exists()
            snapshot = self._snapshot(tabela)
            if existe:
                origem = "arquivo"
            elif esq.get("modo") == "fixa" and snapshot:
                origem = "copia"
            else:
                origem = "faltando"
            repetido = False
            ultima = exe.get("ultima_execucao", {}).get("entradas", {}).get(tabela)
            if existe and ultima:
                try:
                    repetido = (ultima.get("arquivo") == arq
                                and abs(os.path.getmtime(arq) - ultima.get("modificado", -1)) < 1)
                except OSError:
                    pass
            out.append({"tabela": tabela, "modo": esq.get("modo", "mensal"),
                        "vinculada": esq.get("vinculada", ""), "colunas": esq.get("colunas", {}),
                        "arquivo": arq, "existe": existe, "origem": origem,
                        "tem_copia": bool(snapshot), "mesmo_arquivo_da_ultima": repetido,
                        "opcoes": {k: v for k, v in cfg.items() if k not in ("arquivo", "modificado")}})
        return out

    def _snapshot(self, tabela: str) -> Optional[Path]:
        for ext in (".parquet", ".pkl", ".csv"):
            p = self.dir / "dados" / (re.sub(r'[<>:"/\\|?*]', "_", tabela) + ext)
            if p.exists():
                return p
        return None

    # -------------------------------------------------- abrir / executar
    def abrir(self, usar_copias_mensais: bool = False) -> Database:
        """Monta o banco a partir dos arquivos de entrada (sem Access)."""
        exe = self.execucao()
        tables, checks = {}, []
        for tabela, esq in self.entradas.items():
            cfg = exe["entradas"].get(tabela)
            if cfg and cfg.get("arquivo"):
                df, chk = self.ler_entrada(tabela, cfg)
                checks.append(chk)
                if chk.ok:
                    tables[tabela] = df
                continue
            snap = self._snapshot(tabela)
            if snap and (esq.get("modo") == "fixa" or usar_copias_mensais):
                tables[tabela] = _load_snapshot(snap)
                checks.append(Checagem(tabela, str(snap), len(tables[tabela]), list(tables[tabela].columns)))
            else:
                checks.append(Checagem(tabela, erro="tabela mensal sem arquivo de entrada informado"
                                       if esq.get("modo") != "fixa" else "sem arquivo e sem cópia em dados/"))
        # tabelas que não são entrada mas existem em dados/ (ex.: --exportar-dados)
        if (self.dir / "dados").exists():
            for p in (self.dir / "dados").iterdir():
                if p.stem not in tables and p.stem not in self.entradas:
                    tables[p.stem] = _load_snapshot(p)
        bad = [c for c in checks if not c.ok]
        if bad:
            raise EntradaError(checks)
        db = Database(tables, self.consultas, self.funcoes, autonumber=self.autonumeracao)
        db.checagens = checks
        return db

    def executar(self, parametros: Optional[dict] = None, destino=None, intermediarias=False,
                 log=print, db: Optional[Database] = None):
        """Abre as entradas, roda a SEQUENCIA e grava a saída. Devolve (log_etapas, caminho)."""
        from .core import run_sequence
        from . import vba
        db = db or self.abrir()
        vba.set_database(db)
        vba._QUERY_SQL.update({k.lower(): v for k, v in self.sql_original.items()})
        if destino is None:
            (self.dir / "saidas").mkdir(exist_ok=True)
            destino = self.dir / "saidas" / f"saida_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        etapas = run_sequence(db, self.sequencia, parametros, destino, intermediarias, log=log)
        exe = self.execucao()
        exe["parametros"] = {k: _json_value(v) for k, v in (parametros or {}).items()}
        exe["ultima_execucao"] = {"quando": datetime.now().isoformat(timespec="seconds"),
                                  "destino": str(destino), "ok": bool((etapas["Situação"] == "ok").all()),
                                  "entradas": exe["entradas"]}
        self.salvar_execucao(exe)
        return etapas, Path(destino)

    def parametros_declarados(self) -> list[dict]:
        """PARAMETERS das consultas da sequência e das que elas usam (para a tela)."""
        from ..sql.parser import parse
        from ..sql.to_pandas import SQLCompiler
        seen, out, fila = set(), [], list(self.sequencia)
        lower = {k.lower(): k for k in self.sql_original}
        while fila:
            nome = fila.pop(0)
            if nome.lower() in seen or nome.lower() not in lower:
                continue
            seen.add(nome.lower())
            sql = self.sql_original[lower[nome.lower()]]
            try:
                st = parse(sql)
                comp = SQLCompiler()
                comp.compile_function(nome, sql)
                fila += [d for d in comp.deps if d.lower() in lower]
            except Exception:  # noqa: BLE001
                continue
            for p, t in st.parameters:
                if p.lower() not in {x["nome"].lower() for x in out}:
                    out.append({"nome": p, "tipo": _param_kind(t), "tipo_access": t or "Variant"})
        return out


def _param_kind(t: str) -> str:
    t = (t or "").lower()
    if t.startswith("datetime") or t == "date":
        return "data"
    if t in ("long", "short", "byte", "integer", "int"):
        return "inteiro"
    if t in ("currency", "ieeedouble", "ieeesingle", "double", "single", "decimal", "float"):
        return "numero"
    if t in ("bit", "yesno"):
        return "logico"
    return "texto"


def _json_value(v):
    if isinstance(v, (pd.Timestamp, datetime)):
        return {"__data__": v.strftime("%Y-%m-%d %H:%M:%S")}
    return v


def from_json_value(v):
    if isinstance(v, dict) and "__data__" in v:
        return pd.Timestamp(v["__data__"])
    return v


def _load_snapshot(p: Path) -> pd.DataFrame:
    if p.suffix == ".parquet":
        return pd.read_parquet(p)
    if p.suffix == ".pkl":
        return pd.read_pickle(p)
    return pd.read_csv(p)


def convert_param(valor, tipo: str):
    """Converte o texto vindo da tela/linha de comando para o tipo do parâmetro."""
    if valor is None or valor == "":
        return None
    if tipo in ("auto", "texto") and isinstance(valor, str):
        # parâmetro sem tipo declarado: o Access converte texto conforme o uso; aqui
        # número e data reconhecíveis viram número/data
        t = valor.strip()
        if tipo == "auto" and re.fullmatch(r"-?\d+([.,]\d+)?|-?\d{1,3}(\.\d{3})+(,\d+)?", t):
            return convert_param(t, "numero")
        if tipo == "auto" and re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}|\d{4}-\d{2}-\d{2}", t):
            return convert_param(t, "data")
        return valor
    if tipo == "data":
        if isinstance(valor, str) and re.match(r"^\d{4}-\d{2}-\d{2}", valor):
            return pd.Timestamp(valor)
        return pd.to_datetime(valor, dayfirst=True)
    if tipo in ("numero", "inteiro"):
        if isinstance(valor, str):
            t = valor.strip().replace(".", "").replace(",", ".") if "," in valor else valor.strip()
            n = float(t)
        else:
            n = float(valor)
        return int(n) if tipo == "inteiro" or n.is_integer() else n
    if tipo == "logico":
        return str(valor).strip().lower() in ("1", "true", "sim", "s", "verdadeiro", "-1")
    return valor
