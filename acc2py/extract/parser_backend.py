"""Extração multiplataforma com a biblioteca ``access-parser``.

* Tabelas: lidas diretamente.
* Consultas: o Access não guarda o SQL das consultas salvas; ele guarda as
  partes em ``MSysQueries``. Aqui o SQL é remontado seguindo o mesmo formato
  documentado pelo projeto Jackcess (atributos 1..11). Junções aninhadas
  complexas e referências cruzadas podem sair diferentes do texto original
  (mas equivalentes); confira no RELATORIO.md.
* VBA (experimental): o código fica comprimido (MS-OVBA) dentro de
  ``MSysAccessStorage``. Se a extração falhar, use o backend ``com`` no Windows
  ou exporte os módulos para uma pasta.
"""
from __future__ import annotations

import re
import struct
from collections import defaultdict
from typing import Optional

import pandas as pd

from . import ExtractorError, ModuleSource

# ------------------------------------------------------------------ MSysQueries
ATTR_TYPE, ATTR_PARAM, ATTR_FLAG, ATTR_REMOTEDB, ATTR_TABLE, ATTR_COLUMN, ATTR_JOIN, \
    ATTR_WHERE, ATTR_GROUPBY, ATTR_HAVING, ATTR_ORDERBY = range(1, 12)
QTYPE = {1: "select", 2: "make_table", 3: "append", 4: "update", 5: "delete", 6: "crosstab",
         7: "ddl", 8: "passthrough", 9: "union"}
SELECT_STAR, DISTINCT, OWNERACCESS, DISTINCTROW, TOP, PERCENT = 0x01, 0x02, 0x04, 0x08, 0x10, 0x20
UNION_FLAG = 0x02
CROSSTAB_PIVOT, CROSSTAB_NORMAL = 0x01, 0x02
APPEND_VALUE = 0x8000
JOIN_TYPES = {1: "INNER JOIN", 2: "LEFT JOIN", 3: "RIGHT JOIN"}
PARAM_TYPES = {0: "Value", 1: "Bit", 2: "Byte", 3: "Short", 4: "Long", 5: "Currency",
               6: "IEEESingle", 7: "IEEEDouble", 8: "DateTime", 9: "Binary", 10: "Text",
               11: "LongBinary", 12: "LongText", 15: "Guid"}


def _q(name: Optional[str]) -> str:
    if name is None:
        return ""
    return ".".join(p if re.fullmatch(r"\w+", p) else f"[{p}]" for p in name.split("."))


class _Src:
    def __init__(self, key, sql):
        self.keys = {key.lower()}
        self.sql = sql
        self.ons: list[str] = []
        self.kind = None
        self.left = self.right = None

    def render(self, top=True):
        if self.kind is None:
            return self.sql
        on = " AND ".join(f"({o})" if len(self.ons) > 1 else o for o in self.ons)
        s = f"{self.left.render(False)} {self.kind} {self.right.render(False)} ON {on}"
        return s if top else f"({s})"


def build_sql(rows: list[dict]) -> str:
    """Remonta o SQL de uma consulta a partir das linhas de MSysQueries."""
    by = defaultdict(list)
    for r in rows:
        by[int(r.get("Attribute") or 0)].append(r)
    qtype_row = (by[ATTR_TYPE] or [{}])[0]
    qtype = QTYPE.get(int(qtype_row.get("Flag") or 1), "select")
    flag_row = (by[ATTR_FLAG] or [{}])[0]
    flags = int(flag_row.get("Flag") or 0)

    if qtype in ("passthrough", "ddl"):
        return (qtype_row.get("Expression") or "").strip()

    if qtype == "union":
        parts = sorted((r for r in by[ATTR_TABLE] if (r.get("Name1") or "").startswith("X7YZ_____")),
                       key=lambda r: r["Name1"])
        if len(parts) != 2:
            raise ExtractorError("UNION sem as duas partes em MSysQueries")
        kw = "UNION" if flags & UNION_FLAG else "UNION ALL"
        a = (parts[0].get("Expression") or "").strip().rstrip(";")
        b = (parts[1].get("Expression") or "").strip().rstrip(";")
        sql = f"{a}\n{kw} {b}"
        order = _orderings(by)
        return sql + (f"\nORDER BY {order}" if order else "") + ";"

    params = ""
    if by[ATTR_PARAM]:
        items = []
        for r in by[ATTR_PARAM]:
            t = PARAM_TYPES.get(int(r.get("Flag") or 0), "Value")
            extra = r.get("LvExtra")
            if t == "Text" and isinstance(extra, (int, float)) and extra:
                t += f"({int(extra)})"
            items.append(f"{_q(r.get('Name1'))} {t}")
        params = "PARAMETERS " + ", ".join(items) + ";\n"

    from_sql = _from_clause(by)
    where = (by[ATTR_WHERE] or [{}])[0].get("Expression")
    groups = [r.get("Expression") for r in by[ATTR_GROUPBY]]
    having = (by[ATTR_HAVING] or [{}])[0].get("Expression")
    order = _orderings(by)

    def tail(sql):
        if where:
            sql += f"\nWHERE {where}"
        if groups:
            sql += "\nGROUP BY " + ", ".join(groups)
        if having:
            sql += f"\nHAVING {having}"
        if order:
            sql += f"\nORDER BY {order}"
        return sql

    if qtype == "update":
        sets = ", ".join(f"{_q(r.get('Name2') or r.get('Name1'))} = {r.get('Expression')}"
                         for r in by[ATTR_COLUMN])
        sql = f"UPDATE {from_sql}\nSET {sets}"
        if where:
            sql += f"\nWHERE {where}"
        return params + sql + ";"

    if qtype == "delete":
        targets = ", ".join(r.get("Expression") for r in by[ATTR_COLUMN]) or "*"
        return params + tail(f"DELETE {targets}\nFROM {from_sql}") + ";"

    select_kw = "SELECT"
    if flags & DISTINCT:
        select_kw += " DISTINCT"
    elif flags & DISTINCTROW:
        select_kw += " DISTINCTROW"
    if flags & TOP:
        select_kw += f" TOP {flag_row.get('Name1')}" + (" PERCENT" if flags & PERCENT else "")

    if qtype == "crosstab":
        pivot = [r for r in by[ATTR_COLUMN] if int(r.get("Flag") or 0) & CROSSTAB_PIVOT]
        normal = [r for r in by[ATTR_COLUMN] if int(r.get("Flag") or 0) & CROSSTAB_NORMAL]
        value = [r for r in by[ATTR_COLUMN] if not int(r.get("Flag") or 0) & (CROSSTAB_PIVOT | CROSSTAB_NORMAL)]
        cols = ", ".join(_col(r) for r in normal)
        sql = f"TRANSFORM {_col(value[0]) if value else '?'}\n{select_kw} {cols}\nFROM {from_sql}"
        sql = tail(sql)
        if pivot:
            sql += f"\nPIVOT {pivot[0].get('Expression')}"
        return params + sql + ";"

    cols = [r for r in by[ATTR_COLUMN] if not (int(r.get("Flag") or 0) & APPEND_VALUE)]
    col_list = ", ".join(_col(r) for r in cols)
    if flags & SELECT_STAR:
        col_list = "*" + (", " + col_list if col_list else "")
    if qtype == "make_table":
        target = qtype_row.get("Name1")
        sql = f"{select_kw} {col_list} INTO {_q(target)}\nFROM {from_sql}"
        return params + tail(sql) + ";"
    if qtype == "append":
        target = qtype_row.get("Name1")
        targets = [r.get("Name2") for r in cols]
        tcols = f" ({', '.join(_q(t) for t in targets)})" if all(targets) else ""
        values = [r for r in by[ATTR_COLUMN] if int(r.get("Flag") or 0) & APPEND_VALUE]
        if values and not by[ATTR_TABLE]:
            return params + f"INSERT INTO {_q(target)}{tcols}\nVALUES ({', '.join(r.get('Expression') for r in values)});"
        sql = f"INSERT INTO {_q(target)}{tcols}\n{select_kw} {col_list}\nFROM {from_sql}"
        return params + tail(sql) + ";"
    sql = f"{select_kw} {col_list}" + (f"\nFROM {from_sql}" if from_sql else "")
    return params + tail(sql) + ";"


def _col(r) -> str:
    expr = r.get("Expression") or ""
    alias = r.get("Name1")
    return f"{expr} AS {_q(alias)}" if alias else expr


def _orderings(by) -> str:
    out = []
    for r in by[ATTR_ORDERBY]:
        out.append(r.get("Expression") + (" DESC" if (r.get("Name1") or "").upper() == "D" else ""))
    return ", ".join(out)


def _from_clause(by) -> str:
    sources: list[_Src] = []
    for r in by[ATTR_TABLE]:
        name, alias, expr = r.get("Name1"), r.get("Name2"), r.get("Expression")
        if expr and not name:  # subconsulta no FROM
            sql = f"({expr.strip().rstrip(';')}) AS {_q(alias)}"
        else:
            sql = (f"{_q(expr)}." if expr and expr.strip().upper() != "IN" else "") + _q(name) + \
                  (f" AS {_q(alias)}" if alias else "")
        src = _Src(alias or name or "", sql)
        if name:
            src.keys.add(name.lower())
        sources.append(src)

    def find(key):
        for s in sources:
            if key and key.lower() in s.keys:
                return s
        return None

    for r in by[ATTR_JOIN]:
        a, b = find(r.get("Name1")), find(r.get("Name2"))
        if a is None or b is None:
            raise ExtractorError(f"JOIN com tabela desconhecida: {r.get('Name1')} / {r.get('Name2')}")
        kind = JOIN_TYPES.get(int(r.get("Flag") or 1), "INNER JOIN")
        if a is b:
            a.ons.insert(0, r.get("Expression"))
            continue
        j = _Src("", "")
        j.keys = a.keys | b.keys
        j.kind, j.left, j.right, j.ons = kind, a, b, [r.get("Expression")]
        sources = [s for s in sources if s is not a and s is not b] + [j]
    return ", ".join(s.render() for s in sources)


# ------------------------------------------------------------------ MS-OVBA
def ovba_decompress(data: bytes) -> bytes:
    """Descompressão MS-OVBA (§2.4.1) usada nos streams do projeto VBA."""
    if not data or data[0] != 0x01:
        raise ValueError("assinatura de contêiner comprimido inválida")
    out = bytearray()
    pos = 1
    n = len(data)
    while pos < n:
        header = struct.unpack_from("<H", data, pos)[0]
        chunk_size = (header & 0x0FFF) + 3
        compressed = (header >> 15) & 1
        chunk_end = min(pos + chunk_size, n)
        pos += 2
        chunk_start = len(out)
        if not compressed:
            out += data[pos: pos + 4096]
            pos += 4096
            continue
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                else:
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    diff = len(out) - chunk_start
                    bit_count = max((diff - 1).bit_length(), 4)
                    length_mask = 0xFFFF >> bit_count
                    offset = (token >> (16 - bit_count)) + 1
                    length = (token & length_mask) + 3
                    for _ in range(length):
                        out.append(out[-offset])
    return bytes(out)


def parse_dir_stream(data: bytes) -> tuple[int, list[dict]]:
    """Lê o stream 'dir' (já descomprimido): página de código e módulos."""
    pos, codepage = 0, 1252
    modules: list[dict] = []
    cur: Optional[dict] = None
    while pos + 6 <= len(data):
        rid, size = struct.unpack_from("<HI", data, pos)
        pos += 6
        if rid == 0x0009:  # PROJECTVERSION: tamanho declarado 4, real 6
            size = 6
        body = data[pos: pos + size]
        pos += size
        if rid == 0x0003:
            codepage = struct.unpack_from("<H", body)[0]
        elif rid == 0x0019:
            cur = {"name": body.decode(f"cp{codepage}", "replace"), "type": "standard"}
            modules.append(cur)
        elif rid == 0x0047 and cur is not None:
            cur["name"] = body.decode("utf-16-le", "replace")
        elif rid == 0x001A and cur is not None:
            cur["stream"] = body.decode(f"cp{codepage}", "replace")
        elif rid == 0x0032 and cur is not None:
            cur["stream"] = body.decode("utf-16-le", "replace")
        elif rid == 0x0031 and cur is not None:
            cur["offset"] = struct.unpack_from("<I", body)[0]
        elif rid == 0x0022 and cur is not None:
            cur["type"] = "class"
        elif rid == 0x0010:  # fim do dir
            break
    return codepage, modules


# ------------------------------------------------------------------ backend
class ParserExtractor:
    def __init__(self, path: str):
        try:
            from access_parser import AccessParser
        except ImportError as e:
            raise ExtractorError("O backend 'parser' precisa de: pip install access-parser") from e
        self.path = path
        self.db = AccessParser(path)
        self._cache: dict[str, dict] = {}

    def _raw(self, name) -> dict:
        if name not in self._cache:
            self._cache[name] = self.db.parse_table(name)
        return self._cache[name]

    def _rows(self, name) -> list[dict]:
        t = self._raw(name)
        cols = list(t.keys())
        n = max((len(v) for v in t.values()), default=0)
        return [{c: (t[c][i] if i < len(t[c]) else None) for c in cols} for i in range(n)]

    def table_names(self) -> list[str]:
        return [n for n in self.db.catalog if not n.startswith(("MSys", "~", "USys", "f_"))]

    def read_table(self, name: str) -> pd.DataFrame:
        return pd.DataFrame(dict(self._raw(name)))

    def queries(self, include_hidden: bool = False) -> dict[str, str]:
        objs = {r["Id"]: r for r in self._rows("MSysObjects") if r.get("Type") == 5}
        rows = defaultdict(list)
        for r in self._rows("MSysQueries"):
            rows[r.get("ObjectId")].append(r)
        out = {}
        for oid, obj in objs.items():
            name = obj.get("Name") or f"consulta_{oid}"
            if name.startswith("~") and not include_hidden:
                continue
            qrows = sorted(rows.get(oid, []), key=lambda r: (r.get("Attribute") or 0, str(r.get("Order") or "")))
            try:
                out[name] = build_sql(qrows)
            except Exception as e:  # noqa: BLE001
                out[name] = f"-- ERRO ao remontar SQL de MSysQueries: {e}"
        return out

    def modules(self) -> list[ModuleSource]:
        try:
            storage = self._rows("MSysAccessStorage")
        except Exception as e:  # noqa: BLE001
            raise ExtractorError(f"não foi possível ler MSysAccessStorage: {e}") from e
        nodes = {r.get("Id"): r for r in storage}
        children = defaultdict(list)
        for r in storage:
            children[r.get("ParentId")].append(r)
        out = []
        for node in storage:
            if str(node.get("Name")).lower() != "dir":
                continue
            parent = node.get("ParentId")
            siblings = {str(s.get("Name")).lower(): s for s in children[parent]}
            try:
                codepage, mods = parse_dir_stream(ovba_decompress(_blob(node.get("Lv"))))
            except Exception:  # noqa: BLE001
                continue
            for m in mods:
                s = siblings.get(str(m.get("stream", "")).lower())
                if s is None:
                    continue
                data = _blob(s.get("Lv"))
                try:
                    code = ovba_decompress(data[m.get("offset", 0):]).decode(f"cp{codepage}", "replace")
                except Exception:  # noqa: BLE001
                    continue
                name = m["name"]
                kind = m["type"]
                if name.lower().startswith("form_"):
                    kind = "form"
                elif name.lower().startswith("report_"):
                    kind = "report"
                if code.strip():
                    out.append(ModuleSource(name, kind, code))
        if not out:
            raise ExtractorError("VBA não encontrado/legível em MSysAccessStorage. Use o backend 'com' "
                                 "no Windows ou exporte os módulos (.bas/.cls) para uma pasta.")
        return out

    def linked_tables(self) -> dict[str, str]:
        out = {}
        for r in self._rows("MSysObjects"):
            if r.get("Type") in (4, 6) and not str(r.get("Name", "")).startswith("MSys"):
                out[r["Name"]] = f"{r.get('Connect') or ''} {r.get('Database') or ''} | {r.get('ForeignName') or ''}".strip()
        return out

    def columns(self, name: str) -> dict[str, str]:
        from ..runtime.inputs import schema_from_frame
        return schema_from_frame(self.read_table(name))

    def autonumber_columns(self) -> dict[str, str]:
        return {}

    def close(self):
        pass


def _blob(v) -> bytes:
    if v is None:
        return b""
    if isinstance(v, (bytes, bytearray)):
        return bytes(v)
    if isinstance(v, str):
        try:
            return bytes.fromhex(v)
        except ValueError:
            return v.encode("latin-1")
    return bytes(v)
