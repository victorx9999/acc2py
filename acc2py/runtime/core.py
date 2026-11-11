"""Runtime das consultas transpiladas: semântica do Access sobre pandas.

Pontos de fidelidade tratados aqui:
* Null: comparações com Null dão Null (lógica de três valores, dtype "boolean").
* Texto: =, <>, GROUP BY, DISTINCT, JOIN e ORDER BY não diferenciam maiúsculas
  (como no Access com "Option Compare Database"). Desligue com
  ``CASE_INSENSITIVE = False``.
* JOIN: chaves Null nunca casam (o pandas casaria NaN com NaN).
* TOP n: inclui empates na última posição, como o Access.
* ORDER BY: Null vem primeiro em ordem crescente.
* Nomes de coluna resolvidos sem diferenciar maiúsculas; o que não é coluna
  vira parâmetro (é assim que o Access trata ``[Data inicial]``).
"""
from __future__ import annotations

import math
import re
from typing import Any, Callable, Optional

import numpy as np
import pandas as pd

from . import vbafuncs as V

CASE_INSENSITIVE = True


class MissingParameter(KeyError):
    def __str__(self):
        return f"Parâmetro não informado: {self.args[0]!r}. Passe-o em p={{...}}."


class AccessError(Exception):
    pass


# ================================================================ dicionários
def _norm(name: str) -> str:
    s = name.strip().lower().replace("!", ".")
    return re.sub(r"\s*\.\s*", ".", s.replace("[", "").replace("]", ""))


class CIDict(dict):
    """Dicionário com chaves de texto sem diferenciar maiúsculas/colchetes."""

    def __init__(self, data=None):
        super().__init__()
        self._names = {}
        for k, v in (data or {}).items():
            self[k] = v

    def __setitem__(self, k, v):
        n = _norm(k)
        self._names.setdefault(n, k)
        super().__setitem__(n, v)

    def __getitem__(self, k):
        return super().__getitem__(_norm(k))

    def __contains__(self, k):
        return super().__contains__(_norm(k))

    def get(self, k, default=None):
        return super().get(_norm(k), default)

    def pop(self, k, *a):
        self._names.pop(_norm(k), None)
        return super().pop(_norm(k), *a)

    def original_names(self):
        return [self._names[k] for k in self.keys()]


class Params(CIDict):
    """Parâmetros de consulta (``[Data inicial]``, ``Forms!Pedidos!txtCliente``…)."""

    def __init__(self, data=None):
        if isinstance(data, Params):
            data = {data._names[k]: v for k, v in dict.items(data)}
        super().__init__(data)

    def value(self, name):
        if name in self:
            return self[name]
        raise MissingParameter(name)


# ================================================================ banco
class Database:
    """Tabelas (DataFrames) + consultas transpiladas + funções VBA públicas."""

    def __init__(self, tables=None, queries=None, functions=None, autonumber=None):
        self.tables = CIDict(tables)
        self.queries = CIDict(queries)
        self.functions = CIDict(functions)
        self.autonumber = CIDict(autonumber)  # {"Tabela": "ColunaAutoNumeração"}
        self._sql_cache: dict[str, Callable] = {}

    # -------- acesso
    def get(self, name: str, p=None) -> pd.DataFrame:
        if name in self.tables:
            return self.tables[name]
        if name in self.queries:
            res = self.queries[name](self, p)
            if not isinstance(res, pd.DataFrame):
                raise AccessError(f"'{name}' é consulta de ação e não pode ser usada como fonte")
            return res
        raise AccessError(f"Tabela ou consulta não encontrada: {name!r}")

    def run(self, name: str, params: Optional[dict] = None):
        """Executa uma consulta salva pelo nome do Access."""
        return self.queries[name](self, params)

    def sql(self, sql: str, params: Optional[dict] = None):
        """Executa SQL do Access em tempo de execução (transpila e guarda em cache)."""
        fn = self._sql_cache.get(sql)
        if fn is None:
            fn = compile_sql(sql)
            self._sql_cache[sql] = fn
        return fn(self, params)

    # -------- construção
    @classmethod
    def from_folder(cls, folder, queries=None, functions=None, **kw):
        """Carrega tabelas de uma pasta com .parquet/.csv/.pkl (um arquivo por tabela)."""
        from pathlib import Path
        tables = {}
        for f in sorted(Path(folder).rglob("*")):
            if f.suffix == ".parquet":
                tables[f.stem] = pd.read_parquet(f)
            elif f.suffix == ".pkl":
                tables[f.stem] = pd.read_pickle(f)
            elif f.suffix == ".csv":
                tables[f.stem] = parse_dates(pd.read_csv(f))
        return cls(tables, queries, functions, **kw)

    @classmethod
    def from_access(cls, path, queries=None, functions=None, backend="auto", **kw):
        """Lê as tabelas direto do .accdb/.mdb (ver acc2py.extract)."""
        from ..extract import open_extractor
        ex = open_extractor(path, backend)
        tables = {name: ex.read_table(name) for name in ex.table_names()}
        return cls(tables, queries, functions, **kw)


_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$")


def parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    """Converte colunas de texto que são datas ISO (comum em CSV exportado) para datetime."""
    for c in df.columns:
        s = df[c]
        if s.dtype == object or pd.api.types.is_string_dtype(s):
            vals = s.dropna()
            if len(vals) and vals.map(lambda v: isinstance(v, str) and bool(_ISO_DATE.match(v))).all():
                df[c] = pd.to_datetime(s, errors="coerce")
    return df


def compile_sql(sql: str) -> Callable:
    from ..sql.to_pandas import SQLCompiler
    from . import sqlfuncs as F
    code = SQLCompiler().compile_function("(dinâmica)", sql, "_dyn")
    ns: dict[str, Any] = {"rt": __import__(__name__, fromlist=["x"]), "F": F, "pd": pd, "np": np}
    exec(code, ns)
    return ns["_dyn"]


# ================================================================ valores
def is_series(v) -> bool:
    return isinstance(v, pd.Series)


def _isna_scalar(v) -> bool:
    return V.is_null(v)


def broadcast(v, f) -> pd.Series:
    if is_series(v):
        return v
    if isinstance(v, pd.DataFrame):
        raise AccessError("subconsulta retornou uma tabela onde se esperava um valor")
    if v is None:
        return pd.Series([None] * len(f), index=f.index, dtype=object)
    return pd.Series([v] * len(f), index=f.index)


def _fold(s):
    """Normaliza texto para comparação sem diferenciar maiúsculas."""
    if not CASE_INSENSITIVE:
        return s
    if is_series(s):
        if s.dtype == object or pd.api.types.is_string_dtype(s):
            return s.map(lambda x: x.casefold() if isinstance(x, str) else x)
        return s
    return s.casefold() if isinstance(s, str) else s


# ================================================================ colunas
def _resolve(columns, name: str) -> Optional[str]:
    low = _norm(name)
    cmap = {_norm(c): c for c in columns}
    if low in cmap:
        return cmap[low]
    if "." not in low:
        hits = [c for k, c in cmap.items() if k.rsplit(".", 1)[-1] == low and not k.endswith("_rowid")]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            raise AccessError(f"Referência ambígua a {name!r}: {hits}. Qualifique com a tabela.")
    else:
        # "Tabela.Coluna" quando a fonte foi apelidada com o próprio nome da coluna base
        base = low.rsplit(".", 1)[-1]
        hits = [c for k, c in cmap.items() if k == base]
        if len(hits) == 1:
            return hits[0]
    return None


def ref(f, name: str, p: Params):
    """Valor de uma referência: coluna de ``f`` (DataFrame ou linha) ou parâmetro."""
    if f is not None:
        cols = f.index if isinstance(f, pd.Series) else f.columns
        c = _resolve(cols, name)
        if c is not None:
            v = f[c]
            return None if isinstance(f, pd.Series) and _isna_scalar(v) else v
    return p.value(name)


def col_name(df: pd.DataFrame, name: str) -> str:
    c = _resolve(df.columns, name)
    if c is None:
        raise AccessError(f"Coluna não encontrada: {name!r} (colunas: {list(df.columns)})")
    return c


# ================================================================ FROM
def source(db: Database, name: str, alias: str, p: Params, rowid: bool = False) -> pd.DataFrame:
    df = db.get(name, p)
    out = df.copy()
    if rowid:
        out["_rowid"] = out.index
    out.columns = [f"{alias}.{c}" for c in out.columns]
    return out.reset_index(drop=True)


def as_source(df: pd.DataFrame, alias: str) -> pd.DataFrame:
    out = df.copy()
    out.columns = [f"{alias}.{c}" for c in out.columns]
    return out.reset_index(drop=True)


def single_row() -> pd.DataFrame:
    return pd.DataFrame(index=[0])


def join(left: pd.DataFrame, right: pd.DataFrame, how: str, on=None, cond=None) -> pd.DataFrame:
    """JOIN com semântica SQL: chaves Null não casam; condição extra respeitada no LEFT/RIGHT."""
    if how == "right":
        res = join(right, left, "left", [(b, a) for a, b in (on or [])], cond)
        return res[list(left.columns) + list(right.columns)]
    lk = [col_name(left, a) for a, _ in (on or [])]
    rk = [col_name(right, b) for _, b in (on or [])]
    L = left.copy()
    L["__lid"] = np.arange(len(L))
    R = right
    if lk:
        lf = [f"__lk{i}" for i in range(len(lk))]
        rf = [f"__rk{i}" for i in range(len(rk))]
        for i, (a, b) in enumerate(zip(lk, rk)):
            L[lf[i]] = _fold(L[a])
            R = R.copy() if R is right else R
            R[rf[i]] = _fold(R[b])
        Lnn = L.dropna(subset=lf)
        Rnn = R.dropna(subset=rf)
        m = Lnn.merge(Rnn, left_on=lf, right_on=rf, how="inner", sort=False)
        m = m.drop(columns=lf + rf)
        L = L.drop(columns=lf)
    else:
        m = L.merge(R, how="cross")
    if cond is not None:
        m = m.reset_index(drop=True)
        m = m[truth(cond(m))]
    if how == "left":
        unmatched = L[~L["__lid"].isin(m["__lid"])]
        m = pd.concat([m, unmatched], ignore_index=True, sort=False)
    m = m.sort_values("__lid", kind="mergesort").drop(columns="__lid").reset_index(drop=True)
    return m[[c for c in list(left.columns) + list(right.columns) if c in m.columns]]


def where(f: pd.DataFrame, cond) -> pd.DataFrame:
    mask = truth(cond)
    if isinstance(mask, (bool, np.bool_)):
        return f if mask else f.iloc[0:0]
    return f[mask].reset_index(drop=True)


def truth(cond):
    """Converte condição em máscara: Null conta como Falso (como no WHERE)."""
    if is_series(cond):
        if cond.dtype == bool:
            return cond.to_numpy()
        return cond.astype("boolean").fillna(False).astype(bool).to_numpy()
    if _isna_scalar(cond):
        return False
    return bool(cond)


# ================================================================ GROUP BY
_AGG = {
    "sum": lambda s: s.sum(min_count=1),
    "avg": lambda s: s.mean(),
    "min": lambda s: s.min(),
    "max": lambda s: s.max(),
    "count": lambda s: s.count(),
    "count_distinct": lambda s: s.nunique(),
    "first": lambda s: s.iloc[0] if len(s) else None,
    "last": lambda s: s.iloc[-1] if len(s) else None,
    "stdev": lambda s: s.std(ddof=1),
    "stdevp": lambda s: s.std(ddof=0),
    "var": lambda s: s.var(ddof=1),
    "varp": lambda s: s.var(ddof=0),
}


def _agg_series(fn, s):
    if fn in ("min", "max") and s.dtype == object and CASE_INSENSITIVE:
        vals = s.dropna()
        if vals.empty:
            return None
        pick = vals.map(lambda x: x.casefold() if isinstance(x, str) else x)
        idx = pick.idxmin() if fn == "min" else pick.idxmax()
        return vals[idx]
    return _AGG[fn](s)


def group(f: pd.DataFrame, keys: dict, aggs: dict) -> pd.DataFrame:
    tmp = pd.DataFrame(index=f.index)
    for k, v in keys.items():
        tmp[k] = broadcast(v, f)
    for a, (fn, v) in aggs.items():
        tmp[a] = 1 if v is None else broadcast(v, f)
    if not keys:
        row = {}
        for a, (fn, v) in aggs.items():
            row[a] = len(f) if (fn == "count" and v is None) else _agg_series(fn, tmp[a])
        return pd.DataFrame([row])
    folded = []
    for k in keys:
        fk = "__f" + k
        tmp[fk] = _fold(tmp[k])
        folded.append(fk)
    g = tmp.groupby(folded, dropna=False, sort=True)
    out = {}
    for k in keys:
        out[k] = g[k].agg(lambda s: s.iloc[0])
    for a, (fn, v) in aggs.items():
        out[a] = g.size() if (fn == "count" and v is None) else g[a].agg(lambda s, fn=fn: _agg_series(fn, s))
    res = pd.DataFrame(out).reset_index(drop=True)
    return res if len(res.columns) else pd.DataFrame(columns=list(keys) + list(aggs))


# ================================================================ SELECT
class STAR:
    """Marca de ``*`` / ``tabela.*`` na lista do SELECT."""

    def __init__(self, table: Optional[str] = None):
        self.table = table


def _sort(df: pd.DataFrame, keys: list[tuple[Any, bool]]) -> pd.DataFrame:
    if not keys or df.empty:
        return df
    tmp = pd.DataFrame(index=df.index)
    by, asc = [], []
    for i, (v, desc) in enumerate(keys):
        s = broadcast(v, df)
        tmp[f"n{i}"] = s.notna().astype(int)  # Null é o menor valor
        tmp[f"k{i}"] = _fold(s)
        by += [f"n{i}", f"k{i}"]
        asc += [not desc, not desc]
    order = tmp.sort_values(by=by, ascending=asc, kind="mergesort", na_position="last").index
    return df.loc[order]


def project(f: pd.DataFrame, cols: list, order_by=None, distinct=False, top=None,
            top_percent=False) -> pd.DataFrame:
    data: dict[str, pd.Series] = {}
    for item in cols:
        if isinstance(item, STAR):
            src = [c for c in f.columns if not c.endswith("._rowid") and not c.startswith("_")]
            if item.table:
                src = [c for c in src if c.lower().startswith(item.table.lower() + ".")]
            bases = [c.split(".", 1)[-1] for c in src]
            for c, b in zip(src, bases):
                name = b if bases.count(b) == 1 else c
                data[name] = f[c]
        else:
            name, value = item
            data[name] = broadcast(value, f)
    out = pd.DataFrame(data, index=f.index)
    if order_by:
        sort_keys = [(broadcast(v, f), d) for v, d in order_by]
        sorter = pd.DataFrame({f"__o{i}": s for i, (s, _) in enumerate(sort_keys)}, index=f.index)
        out = pd.concat([out, sorter], axis=1)
        out = _sort(out, [(out[f"__o{i}"], d) for i, (_, d) in enumerate(sort_keys)])
    if distinct:
        visible = [c for c in out.columns if not c.startswith("__o")]
        fold = pd.DataFrame({c: _fold(out[c]) for c in visible}, index=out.index)
        out = out[~fold.duplicated(keep="first")]
    if top is not None:
        n = math.ceil(len(out) * top / 100) if top_percent else int(top)
        if order_by and 0 < n < len(out):
            okeys = [c for c in out.columns if c.startswith("__o")]
            last = out.iloc[n - 1][okeys]
            tail = out.iloc[n:]
            ties = (tail[okeys].fillna("\x00null") == last.fillna("\x00null")).all(axis=1)
            extra = ties.cummin()  # empates contíguos logo após a posição n
            out = pd.concat([out.iloc[:n], tail[extra]])
        else:
            out = out.iloc[:n]
    out = out[[c for c in out.columns if not c.startswith("__o")]]
    return out.reset_index(drop=True)


def order(df: pd.DataFrame, keys) -> pd.DataFrame:
    return _sort(df, keys).reset_index(drop=True)


def union(frames: list[pd.DataFrame], all_flags: list[bool]) -> pd.DataFrame:
    names = list(frames[0].columns)
    res = frames[0]
    for fr, keep_all in zip(frames[1:], all_flags):
        if len(fr.columns) != len(names):
            raise AccessError("UNION com número diferente de colunas")
        fr = fr.copy()
        fr.columns = names
        res = pd.concat([res, fr], ignore_index=True)
        if not keep_all:
            fold = pd.DataFrame({c: _fold(res[c]) for c in names})
            res = res[~fold.duplicated()]
    return res.reset_index(drop=True)


def crosstab(f, rows: dict, pivot, value, extras: dict, pivot_in=None) -> pd.DataFrame:
    fn, v = value
    keys = dict(rows)
    keys["__pv"] = pivot
    cell = group(f, keys, {"__v": (fn, v)})
    row_names = list(rows)
    cell["__pv"] = cell["__pv"].map(lambda x: "<>" if _isna_scalar(x) else V.CStr(x))
    if row_names:
        wide = cell.pivot_table(index=row_names, columns="__pv", values="__v",
                                aggfunc="first", dropna=False, sort=True)
    else:
        wide = cell.set_index("__pv")[["__v"]].T
    wide.columns = [str(c) for c in wide.columns]
    if pivot_in:
        heads = [V.CStr(x) for x in pivot_in]
        wide = wide.reindex(columns=heads)
    wide = wide.reset_index() if row_names else wide.reset_index(drop=True)
    if extras:
        ex = group(f, rows, extras)
        wide = wide.merge(ex, on=row_names, how="left") if row_names else pd.concat([wide, ex], axis=1)
        ordered = row_names + list(extras) + [c for c in wide.columns if c not in row_names and c not in extras]
        wide = wide[ordered]
    if row_names:  # o Access ordena as linhas pelos cabeçalhos (sem diferenciar maiúsculas)
        wide = order(wide, [(wide[r], False) for r in row_names])
    return wide[[c for c in wide.columns if not str(c).startswith("_g")]]


# ================================================================ ação
def _rowids(f, alias):
    return f[col_name(f, f"{alias}._rowid")]


def apply_update(db: Database, table: str, f: pd.DataFrame, alias: str, sets: dict) -> int:
    t = db.tables[table]
    ids = _rowids(f, alias)
    for col, val in sets.items():
        c = col_name(t, col)
        vals = broadcast(val, f)
        if t[c].dtype != object and vals.dtype == object:
            t[c] = t[c].astype(object)
        t.loc[ids.to_numpy(), c] = vals.to_numpy()
    return int(ids.nunique())


def apply_delete(db: Database, table: str, f: pd.DataFrame, alias: str) -> int:
    t = db.tables[table]
    ids = set(_rowids(f, alias))
    db.tables[table] = t.drop(index=list(ids)).reset_index(drop=True)
    return len(ids)


def apply_insert(db: Database, table: str, rows: pd.DataFrame, columns=None) -> int:
    t = db.tables[table] if table in db.tables else pd.DataFrame()
    rows = rows.copy()
    if columns:
        rows.columns = [col_name(t, c) if len(t.columns) else c for c in columns]
    else:
        rows.columns = [(_resolve(t.columns, c) or c) for c in rows.columns]
    auto = db.autonumber.get(table)
    if auto and auto not in rows.columns:
        start = int(t[auto].max()) + 1 if len(t) and t[auto].notna().any() else 1
        rows[auto] = range(start, start + len(rows))
    db.tables[table] = pd.concat([t, rows], ignore_index=True)
    return len(rows)


def apply_insert_values(db: Database, table: str, columns, values: list) -> int:
    t = db.tables[table]
    cols = columns or list(t.columns)[: len(values)]
    return apply_insert(db, table, pd.DataFrame([values], columns=cols))


def make_table(db: Database, name: str, out: pd.DataFrame) -> int:
    db.tables[name] = out.copy()
    return len(out)


# ================================================================ operadores
def _bool(x):
    if is_series(x):
        if x.dtype == "boolean":
            return x
        if x.dtype == bool:
            return x.astype("boolean")
        return x.map(lambda v: pd.NA if _isna_scalar(v) else bool(v)).astype("boolean")
    return pd.NA if _isna_scalar(x) else bool(x)


def _cmp(a, b, op):
    if not is_series(a) and not is_series(b):
        if _isna_scalar(a) or _isna_scalar(b):
            return None
        return op(_fold(a), _fold(b))
    if not is_series(a):
        a, b = broadcast(a, b), b
    elif not is_series(b):
        b = broadcast(b, a)
    na = a.isna() | b.isna()
    fa, fb = _fold(a), _fold(b)
    try:
        res = op(fa, fb)
    except TypeError:
        res = pd.Series([op(x, y) if not (_isna_scalar(x) or _isna_scalar(y)) else False
                         for x, y in zip(fa, fb)], index=a.index)
    res = res.astype("boolean")
    res[na.to_numpy()] = pd.NA
    return res


def eq(a, b): return _cmp(a, b, lambda x, y: x == y)
def ne(a, b): return _cmp(a, b, lambda x, y: x != y)
def lt(a, b): return _cmp(a, b, lambda x, y: x < y)
def gt(a, b): return _cmp(a, b, lambda x, y: x > y)
def le(a, b): return _cmp(a, b, lambda x, y: x <= y)
def ge(a, b): return _cmp(a, b, lambda x, y: x >= y)


def _scalar_bool(x):
    return None if x is pd.NA else x


def and_(a, b):
    r = _bool(a) & _bool(b)
    return r if is_series(r) else _scalar_bool(r)


def or_(a, b):
    r = _bool(a) | _bool(b)
    return r if is_series(r) else _scalar_bool(r)


def xor_(a, b):
    r = _bool(a) ^ _bool(b)
    return r if is_series(r) else _scalar_bool(r)


def not_(a):
    b = _bool(a)
    if is_series(b):
        return ~b
    return None if b is pd.NA else (not b)


def eqv(a, b): return not_(xor_(a, b))
def imp(a, b): return or_(not_(a), b)


def isnull(x):
    return x.isna() if is_series(x) else _isna_scalar(x)


def notnull(x):
    return x.notna() if is_series(x) else not _isna_scalar(x)


def between(x, lo, hi):
    return and_(ge(x, lo), le(x, hi))


def isin(x, items):
    if is_series(items):
        items = list(items)
    if any(is_series(i) for i in items):
        res = None
        for i in items:
            e = eq(x, i)
            res = e if res is None else or_(res, e)
        return res
    vals = {_fold(i) for i in items if not _isna_scalar(i)}
    if is_series(x):
        res = _fold(x).isin(vals).astype("boolean")
        res[x.isna().to_numpy()] = pd.NA
        return res
    return None if _isna_scalar(x) else _fold(x) in vals


_LIKE_CACHE: dict[str, re.Pattern] = {}


def like_regex(pattern: str) -> re.Pattern:
    """Curingas do Access: * ? # [abc] [!abc] [a-z]."""
    rx = _LIKE_CACHE.get(pattern)
    if rx:
        return rx
    out, i = [], 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        elif ch == "#":
            out.append(r"\d")
        elif ch == "[":
            j = pattern.find("]", i + 1)
            if j == -1:
                out.append(r"\[")
            else:
                body = pattern[i + 1: j]
                if body.startswith("!"):
                    body = "^" + re.escape(body[1:]).replace(r"\-", "-")
                elif body in ("*", "?", "#", "["):
                    body = re.escape(body)
                else:
                    body = re.escape(body).replace(r"\-", "-")
                out.append(f"[{body}]")
                i = j
        else:
            out.append(re.escape(ch))
        i += 1
    rx = re.compile("".join(out) + r"\Z", re.IGNORECASE | re.DOTALL)
    _LIKE_CACHE[pattern] = rx
    return rx


def like(x, pattern):
    if is_series(pattern):
        vals = [None if (_isna_scalar(a) or _isna_scalar(b)) else bool(like_regex(str(b)).match(V.CStr(a)))
                for a, b in zip(broadcast(x, pattern), pattern)]
        return pd.Series(vals, index=pattern.index, dtype="boolean")
    if _isna_scalar(pattern):
        return None
    rx = like_regex(str(pattern))
    if is_series(x):
        res = x.map(lambda v: pd.NA if _isna_scalar(v) else bool(rx.match(V.CStr(v))))
        return res.astype("boolean")
    return None if _isna_scalar(x) else bool(rx.match(V.CStr(x)))


def _text(v):
    return "" if _isna_scalar(v) else V.CStr(v)


def concat(a, b):
    """Operador & : Null vira texto vazio; números seguem o formato regional."""
    if is_series(a) or is_series(b):
        idx = a.index if is_series(a) else b.index
        A_ = a if is_series(a) else [a] * len(idx)
        B_ = b if is_series(b) else [b] * len(idx)
        both_null = [(_isna_scalar(x) and _isna_scalar(y)) for x, y in zip(A_, B_)]
        vals = [None if n else _text(x) + _text(y) for x, y, n in zip(A_, B_, both_null)]
        return pd.Series(vals, index=idx, dtype=object)
    if _isna_scalar(a) and _isna_scalar(b):
        return None
    return _text(a) + _text(b)


def _vb_int(x):
    return np.round(x) if is_series(x) else (None if _isna_scalar(x) else round(V._num(x)))


def idiv(a, b):
    """Operador \\ : arredonda os operandos (banker's) e trunca o quociente."""
    ia, ib = _vb_int(a), _vb_int(b)
    if is_series(ia) or is_series(ib):
        return np.trunc(ia / ib)
    if ia is None or ib is None:
        return None
    return int(ia / ib)


def mod(a, b):
    ia, ib = _vb_int(a), _vb_int(b)
    if is_series(ia) or is_series(ib):
        return np.fmod(ia, ib)
    if ia is None or ib is None:
        return None
    return int(math.fmod(ia, ib))


def fail(msg):
    raise AccessError(msg)


# ================================================================ subconsultas
def scalar(df: pd.DataFrame):
    if df.empty:
        return None
    v = df.iloc[0, 0]
    return None if _isna_scalar(v) else v


def exists(df: pd.DataFrame) -> bool:
    return len(df) > 0


def first_col(df: pd.DataFrame) -> pd.Series:
    return df.iloc[:, 0]


def per_row(f: pd.DataFrame, fn, kind: str) -> pd.Series:
    """Subconsulta correlacionada: executa uma vez por linha (lento em tabelas grandes)."""
    vals = []
    for _, row in f.iterrows():
        res = fn(row)
        vals.append(scalar(res) if kind == "scalar" else exists(res))
    return pd.Series(vals, index=f.index)


def per_row_in(f: pd.DataFrame, x, fn) -> pd.Series:
    xs = broadcast(x, f)
    vals = []
    for (_, row), xv in zip(f.iterrows(), xs):
        vals.append(None if _isna_scalar(xv) else bool(isin(xv, first_col(fn(row)))))
    return pd.Series(vals, index=f.index, dtype="boolean")


# ================================================================ exportação
_SHEET_BAD = re.compile(r"[\[\]:*?/\\]")


def _sheet_name(name: str, used: set) -> str:
    base = _SHEET_BAD.sub("_", name)[:31] or "Consulta"
    s, k = base, 2
    while s.lower() in used:
        suf = f" ({k})"
        s = base[: 31 - len(suf)] + suf
        k += 1
    used.add(s.lower())
    return s


def export_results(db: "Database", destino, consultas: list, parametros: Optional[dict] = None,
                   resumo: bool = True) -> pd.DataFrame:
    """Executa várias consultas de seleção e grava tudo numa saída só.

    * ``destino`` terminado em .xlsx: uma aba por consulta + aba "Resumo".
    * qualquer outro caminho: pasta com um .csv por consulta + resumo.csv.

    ``parametros`` pode ser comum a todas ({"Data inicial": ...}) e/ou por
    consulta ({"Vendas por Mês": {"Ano": 2026}}); os por consulta têm prioridade.
    Uma consulta que falha não interrompe as outras: o erro vai para o resumo.
    """
    from pathlib import Path
    parametros = dict(parametros or {})
    per_query = {k.lower(): v for k, v in parametros.items() if isinstance(v, dict)}
    common = {k: v for k, v in parametros.items() if not isinstance(v, dict)}
    results, rows = {}, []
    for name in consultas:
        p = dict(common)
        p.update(per_query.get(name.lower(), {}))
        try:
            df = db.run(name, p)
            if not isinstance(df, pd.DataFrame):
                raise AccessError("é consulta de ação (não gera tabela); não foi executada")
            results[name] = df
            rows.append({"Consulta": name, "Situação": "ok", "Linhas": len(df), "Erro": ""})
        except Exception as e:  # noqa: BLE001
            rows.append({"Consulta": name, "Situação": "erro", "Linhas": 0, "Erro": str(e)})
    summary = pd.DataFrame(rows, columns=["Consulta", "Situação", "Linhas", "Erro"])
    destino = Path(destino)
    if destino.suffix.lower() in (".xlsx", ".xlsm"):
        destino.parent.mkdir(parents=True, exist_ok=True)
        used: set = set()
        with pd.ExcelWriter(destino, engine="openpyxl") as w:
            if resumo:
                summary.to_excel(w, sheet_name=_sheet_name("Resumo", used), index=False)
            for name, df in results.items():
                df.to_excel(w, sheet_name=_sheet_name(name, used), index=False)
    else:
        destino.mkdir(parents=True, exist_ok=True)
        for name, df in results.items():
            df.to_csv(destino / (re.sub(r'[<>:"/\\|?*]', "_", name) + ".csv"), index=False,
                      encoding="utf-8-sig")
        if resumo:
            summary.to_csv(destino / "resumo.csv", index=False, encoding="utf-8-sig")
    return summary


def run_sequence(db: "Database", steps: list, parametros: Optional[dict] = None, destino=None,
                 intermediarias: bool = False, log: Callable = print) -> pd.DataFrame:
    """Executa consultas EM ORDEM, como uma rotina do Access (etapa 1, depois 2, …).

    * Consultas de ação (criar tabela, atualizar, acrescentar, excluir) alteram as
      tabelas do ``db`` em memória, e a etapa seguinte já enxerga essas mudanças.
      O arquivo .accdb de origem nunca é alterado.
    * Se uma etapa falha, a sequência PARA (as seguintes dependeriam dela).
    * ``destino`` (.xlsx ou pasta): grava a saída final (resultado da última etapa
      de seleção) + aba "Etapas" com o log; ``intermediarias=True`` grava também
      o resultado de cada etapa de seleção.
    Devolve o log das etapas.
    """
    import time
    from pathlib import Path
    parametros = dict(parametros or {})
    per_query = {k.lower(): v for k, v in parametros.items() if isinstance(v, dict)}
    common = {k: v for k, v in parametros.items() if not isinstance(v, dict)}
    rows, results, final_name = [], {}, None
    for i, name in enumerate(steps, 1):
        p = dict(common)
        p.update(per_query.get(name.lower(), {}))
        t0 = time.perf_counter()
        try:
            res = db.run(name, p)
            secs = round(time.perf_counter() - t0, 3)
            if isinstance(res, pd.DataFrame):
                results[name] = res
                final_name = name
                rows.append({"Etapa": i, "Consulta": name, "Tipo": "seleção", "Linhas": len(res),
                             "Segundos": secs, "Situação": "ok", "Erro": ""})
                log(f"  [{i}/{len(steps)}] {name}: {len(res)} linha(s)")
            else:
                n = int(res) if res is not None else 0
                rows.append({"Etapa": i, "Consulta": name, "Tipo": "ação", "Linhas": n,
                             "Segundos": secs, "Situação": "ok", "Erro": ""})
                log(f"  [{i}/{len(steps)}] {name}: {n} linha(s) afetada(s)")
        except Exception as e:  # noqa: BLE001
            rows.append({"Etapa": i, "Consulta": name, "Tipo": "", "Linhas": 0,
                         "Segundos": round(time.perf_counter() - t0, 3), "Situação": "ERRO", "Erro": str(e)})
            log(f"  [{i}/{len(steps)}] {name}: ERRO — {e}")
            for j, rest in enumerate(steps[i:], i + 1):
                rows.append({"Etapa": j, "Consulta": rest, "Tipo": "", "Linhas": 0, "Segundos": 0,
                             "Situação": "não executada", "Erro": "etapa anterior falhou"})
            break
    summary = pd.DataFrame(rows)
    if destino is not None:
        out = {}
        if intermediarias:
            out = dict(results)
        elif final_name is not None and summary["Situação"].eq("ok").all():
            out = {f"Final - {final_name}": results[final_name]}
        destino = Path(destino)
        if destino.suffix.lower() in (".xlsx", ".xlsm"):
            destino.parent.mkdir(parents=True, exist_ok=True)
            used: set = set()
            with pd.ExcelWriter(destino, engine="openpyxl") as w:
                for name, df in out.items():
                    df.to_excel(w, sheet_name=_sheet_name(name, used), index=False)
                summary.to_excel(w, sheet_name=_sheet_name("Etapas", used), index=False)
        else:
            destino.mkdir(parents=True, exist_ok=True)
            for name, df in out.items():
                df.to_csv(destino / (re.sub(r'[<>:"/\\|?*]', "_", name) + ".csv"), index=False,
                          encoding="utf-8-sig")
            summary.to_csv(destino / "etapas.csv", index=False, encoding="utf-8-sig")
    return summary
