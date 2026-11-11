"""Extração pelo próprio Microsoft Access via automação COM (Windows)."""
from __future__ import annotations

import os
from datetime import datetime

import pandas as pd

from . import ExtractorError, ModuleSource

DB_SYSTEM_OBJECT = 0x80000002
DB_HIDDEN_OBJECT = 0x1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


class ComExtractor:
    def __init__(self, path: str, password: str | None = None):
        try:
            import win32com.client
        except ImportError as e:
            raise ExtractorError("O backend 'com' precisa do pywin32: pip install pywin32") from e
        self.path = os.path.abspath(path)
        try:
            self.app = win32com.client.DispatchEx("Access.Application")
        except Exception as e:
            raise ExtractorError("Não foi possível abrir o Microsoft Access por COM. O Access "
                                 "(não só o runtime de banco) precisa estar instalado.") from e
        self.app.Visible = False
        try:
            # impede que macros AutoExec/formulário de inicialização rodem código do banco
            self.app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        except Exception:
            pass
        if password:
            self.app.OpenCurrentDatabase(self.path, False, password)
        else:
            self.app.OpenCurrentDatabase(self.path, False)
        self.db = self.app.CurrentDb()

    def close(self):
        try:
            self.app.CloseCurrentDatabase()
            self.app.Quit()
        except Exception:
            pass

    def __del__(self):
        self.close()

    # ---------------------------------------------------------------- consultas
    def queries(self, include_hidden: bool = False) -> dict[str, str]:
        out = {}
        for qd in self.db.QueryDefs:
            name = qd.Name
            if name.startswith("~") and not include_hidden:
                continue  # consultas embutidas de formulários/caixas de combinação
            out[name] = qd.SQL
        return out

    # ---------------------------------------------------------------- VBA
    def modules(self) -> list[ModuleSource]:
        out = []
        try:
            project = self.app.VBE.ActiveVBProject
        except Exception as e:
            raise ExtractorError("Sem acesso ao projeto VBA (protegido por senha?)") from e
        kinds = {1: "standard", 2: "class", 100: "form"}
        for comp in project.VBComponents:
            cm = comp.CodeModule
            n = cm.CountOfLines
            code = cm.Lines(1, n) if n else ""
            kind = kinds.get(comp.Type, "standard")
            if kind == "form" and comp.Name.lower().startswith("report_"):
                kind = "report"
            if code.strip():
                out.append(ModuleSource(comp.Name, kind, code))
        return out

    # ---------------------------------------------------------------- tabelas
    def table_names(self) -> list[str]:
        names = []
        for td in self.db.TableDefs:
            if td.Name.startswith(("MSys", "~", "USys")) or (td.Attributes & DB_SYSTEM_OBJECT):
                continue
            names.append(td.Name)
        return names

    def read_table(self, name: str) -> pd.DataFrame:
        rs = self.db.OpenRecordset(f"SELECT * FROM [{name}]", 4)  # dbOpenSnapshot
        cols = [rs.Fields(i).Name for i in range(rs.Fields.Count)]
        if rs.EOF:
            return pd.DataFrame(columns=cols)
        rs.MoveLast()
        n = rs.RecordCount
        rs.MoveFirst()
        data = rs.GetRows(n)  # tupla por coluna
        rs.Close()
        frame = {}
        for c, values in zip(cols, data):
            frame[c] = [_convert(v) for v in values]
        return pd.DataFrame(frame)

    def autonumber_columns(self) -> dict[str, str]:
        out = {}
        for name in self.table_names():
            td = self.db.TableDefs(name)
            for f in td.Fields:
                if f.Attributes & 0x10:  # dbAutoIncrField
                    out[name] = f.Name
        return out


def _convert(v):
    if hasattr(v, "year") and hasattr(v, "hour"):  # pywintypes.datetime (às vezes com fuso)
        return datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
    return v
