"""Funções do Access usadas dentro de consultas, vetorizadas sobre Series.

Cada função de ``vbafuncs`` ganha aqui uma versão que aceita escalares ou
Series (aplicada linha a linha, com Null -> None). IIf, Nz e as funções de
domínio (DLookup, DSum…) têm implementação própria.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from . import vbafuncs as V


def _py(v):
    if V.is_null(v):
        return None
    if isinstance(v, np.generic):
        return v.item()
    return v


def vectorize(fn):
    def wrapper(*args):
        idx = next((a.index for a in args if isinstance(a, pd.Series)), None)
        if idx is None:
            return fn(*[_py(a) for a in args])
        cols = [a.tolist() if isinstance(a, pd.Series) else [a] * len(idx) for a in args]
        vals = [fn(*[_py(x) for x in row]) for row in zip(*cols)] if cols else []
        return pd.Series(vals, index=idx, dtype=object).infer_objects()
    wrapper.__name__ = fn.__name__
    wrapper.__doc__ = fn.__doc__
    return wrapper


for _name in V.__all__:
    globals()[_name] = vectorize(getattr(V, _name))


def IIf(cond, a, b):  # noqa: F811 — versão vetorizada própria
    if not any(isinstance(x, pd.Series) for x in (cond, a, b)):
        return V.IIf(_py(cond), a, b)
    idx = next(x.index for x in (cond, a, b) if isinstance(x, pd.Series))
    if isinstance(cond, pd.Series):
        mask = cond.astype("boolean").fillna(False).astype(bool).to_numpy()
    else:
        mask = np.full(len(idx), bool(_py(cond)) if _py(cond) is not None else False)
    av = a.to_numpy(dtype=object) if isinstance(a, pd.Series) else np.full(len(idx), a, dtype=object)
    bv = b.to_numpy(dtype=object) if isinstance(b, pd.Series) else np.full(len(idx), b, dtype=object)
    return pd.Series(np.where(mask, av, bv), index=idx, dtype=object).infer_objects()


def Nz(v, default=None):  # noqa: F811
    if isinstance(v, pd.Series):
        fill = "" if default is None else default
        if isinstance(fill, pd.Series):
            return v.where(v.notna(), fill).infer_objects()
        return v.astype(object).where(v.notna(), fill).infer_objects()
    return V.Nz(_py(v), default)


def IsNull(v):  # noqa: F811
    return v.isna() if isinstance(v, pd.Series) else V.is_null(v)


# ------------------------------------------------------------- domínio
def _bracket(name: str) -> str:
    name = str(name).strip()
    return name if name.startswith("[") or " " not in name and name.isidentifier() else f"[{name}]"


def _domain(agg, db, expr, domain, criteria=None):
    expr = str(expr)
    target = "*" if expr.strip() == "*" else expr
    sql = f"SELECT {agg}({target}) AS v FROM {_bracket(domain)}"
    if criteria not in (None, ""):
        sql += f" WHERE {criteria}"
    return _scalar(db.sql(sql))


def _scalar(df):
    if df.empty:
        return None
    v = df.iloc[0, 0]
    return _py(v)


def _make_domain(agg):
    def fn(db, expr, domain, criteria=None):
        scalar_fn = lambda e, d, c=None: _domain(agg, db, e, d, c)
        return vectorize(scalar_fn)(expr, domain, criteria) if any(
            isinstance(x, pd.Series) for x in (expr, domain, criteria)) else scalar_fn(expr, domain, criteria)
    fn.__name__ = "D" + agg
    return fn


DLookup = _make_domain("First")
DFirst = _make_domain("First")
DLast = _make_domain("Last")
DSum = _make_domain("Sum")
DAvg = _make_domain("Avg")
DMin = _make_domain("Min")
DMax = _make_domain("Max")
DStDev = _make_domain("StDev")
DStDevP = _make_domain("StDevP")
DVar = _make_domain("Var")
DVarP = _make_domain("VarP")


def DCount(db, expr, domain, criteria=None):
    def one(e, d, c=None):
        v = _domain("Count", db, e, d, c)
        return 0 if v is None else int(v)
    if any(isinstance(x, pd.Series) for x in (expr, domain, criteria)):
        return vectorize(one)(expr, domain, criteria)
    return one(expr, domain, criteria)


# ------------------------------------------------------------- VBA do usuário
def udf(db, name, *args):
    """Função pública de módulo VBA chamada dentro de uma consulta."""
    fn = db.functions.get(name)
    if fn is None:
        raise NameError(f"Função {name!r} usada na consulta não existe nem no Access nem nos "
                        f"módulos VBA transpilados (registre em db.functions).")
    return vectorize(fn)(*args)
