"""Extração a partir de uma pasta com objetos exportados do Access.

Reconhece:
* ``*.sql``                      — SQL puro; o nome do arquivo é o nome da consulta.
* ``*.qry`` / ``*.txt``          — saída de ``Application.SaveAsText acQuery`` (linha ``dbMemo "SQL"``).
* ``*.bas`` / ``*.cls``          — módulos exportados do editor VBA.
* ``*.txt``/``*.frm``/``*.form`` — ``SaveAsText acForm/acReport/acModule``: o código vem
                                   depois de ``CodeBehindForm``.
* ``*.csv`` / ``*.xlsx``         — tabelas (uma por arquivo).

Subpastas são percorridas; o nome do objeto é o nome do arquivo sem extensão
(prefixos comuns de scripts de exportação como ``Query_``, ``Form_`` e ``Module_``
são tratados).
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from . import ModuleSource


def _read_text(p: Path) -> str:
    raw = p.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp1252")


def _unescape_saveastext(s: str) -> str:
    s = s.replace('\\"', '"')
    s = re.sub(r"\\(\d{3})", lambda m: chr(int(m.group(1), 8)), s)
    return s.replace("\\\\", "\\")


def sql_from_saveastext(text: str) -> str | None:
    m = re.search(r'dbMemo\s+"SQL"\s*=\s*"((?:[^"\\]|\\.)*)"', text)
    if not m:
        return None
    sql = _unescape_saveastext(m.group(1))
    return sql.replace("\r\n", "\n").strip()


class FolderExtractor:
    def __init__(self, folder: str):
        self.folder = Path(folder)
        self._queries: dict[str, str] = {}
        self._modules: list[ModuleSource] = []
        self._tables: dict[str, Path] = {}
        self._scan()

    def _scan(self):
        for p in sorted(self.folder.rglob("*")):
            if not p.is_file():
                continue
            ext = p.suffix.lower()
            stem = p.stem
            if ext == ".sql":
                self._queries[re.sub(r"(?i)^(query|qry|consulta)_", "", stem)] = _read_text(p).strip()
            elif ext in (".bas", ".cls"):
                text = _read_text(p)
                m = re.search(r'Attribute VB_Name = "([^"]+)"', text)
                name = m.group(1) if m else stem
                kind = "class" if ext == ".cls" else "standard"
                if name.lower().startswith(("form_", "report_")):
                    kind = "form" if name.lower().startswith("form_") else "report"
                self._modules.append(ModuleSource(name, kind, text))
            elif ext in (".txt", ".qry", ".frm", ".form", ".rpt", ".report", ".acm"):
                text = _read_text(p)
                sql = sql_from_saveastext(text)
                if sql is not None:
                    self._queries[re.sub(r"(?i)^(query|qry|consulta)_", "", stem)] = sql
                elif "CodeBehindForm" in text:
                    code = text.split("CodeBehindForm", 1)[1]
                    is_report = re.search(r"^\s*Begin Report", text, re.M) is not None
                    base = re.sub(r"(?i)^(form|report|frm|rpt)_", "", stem)
                    prefix = "Report_" if is_report else "Form_"
                    self._modules.append(ModuleSource(prefix + base, "report" if is_report else "form", code))
                elif re.search(r"^\s*(Option\s+(Compare|Explicit)|Attribute VB_Name|(Public |Private )?(Sub|Function) )",
                               text, re.M | re.I):
                    m = re.search(r'Attribute VB_Name = "([^"]+)"', text)
                    name = m.group(1) if m else re.sub(r"(?i)^(module|modulo|mod)_", "", stem)
                    self._modules.append(ModuleSource(name, "standard", text))
            elif ext in (".csv", ".xlsx", ".xls", ".parquet", ".pkl"):
                self._tables[stem] = p

    def queries(self, include_hidden: bool = False) -> dict[str, str]:
        return {k: v for k, v in self._queries.items() if include_hidden or not k.startswith("~")}

    def modules(self) -> list[ModuleSource]:
        return list(self._modules)

    def table_names(self) -> list[str]:
        return list(self._tables)

    def read_table(self, name: str) -> pd.DataFrame:
        p = self._tables[name]
        ext = p.suffix.lower()
        if ext == ".csv":
            from ..runtime.core import parse_dates
            return parse_dates(pd.read_csv(p))
        if ext in (".xlsx", ".xls"):
            return pd.read_excel(p)
        if ext == ".parquet":
            return pd.read_parquet(p)
        return pd.read_pickle(p)

    def autonumber_columns(self) -> dict[str, str]:
        return {}

    def close(self):
        pass
