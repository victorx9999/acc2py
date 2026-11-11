"""Converte uma expressão da AST de volta para texto SQL (usado em chaves e comentários)."""
from __future__ import annotations

from datetime import datetime

from . import ast as A


def q(name: str) -> str:
    return name if name.replace("_", "a").isalnum() and not name[0].isdigit() else f"[{name}]"


def expr_sql(e: A.Expr) -> str:
    if isinstance(e, A.Col):
        return ".".join(q(p) for p in e.parts)
    if isinstance(e, A.Star):
        return f"{q(e.table)}.*" if e.table else "*"
    if isinstance(e, A.Lit):
        if e.kind == "null":
            return "Null"
        if e.kind == "str":
            return '"' + e.value.replace('"', '""') + '"'
        if e.kind == "date":
            d: datetime = e.value
            return "#" + d.strftime("%m/%d/%Y" + (" %H:%M:%S" if d.time() != d.min.time() else "")) + "#"
        if e.kind == "bool":
            return "True" if e.value else "False"
        return repr(e.value)
    if isinstance(e, A.Unary):
        return f"{e.op} {expr_sql(e.operand)}" if e.op == "NOT" else f"-{expr_sql(e.operand)}"
    if isinstance(e, A.Bin):
        return f"({expr_sql(e.left)} {e.op} {expr_sql(e.right)})"
    if isinstance(e, A.Func):
        d = "DISTINCT " if e.distinct else ""
        return f"{e.name}({d}{', '.join(expr_sql(a) for a in e.args)})"
    if isinstance(e, A.Between):
        n = "NOT " if e.negated else ""
        return f"{expr_sql(e.expr)} {n}BETWEEN {expr_sql(e.low)} AND {expr_sql(e.high)}"
    if isinstance(e, A.InList):
        n = "NOT " if e.negated else ""
        return f"{expr_sql(e.expr)} {n}IN ({', '.join(expr_sql(i) for i in e.items)})"
    if isinstance(e, A.IsNull):
        return f"{expr_sql(e.expr)} IS {'NOT ' if e.negated else ''}NULL"
    if isinstance(e, A.Like):
        return f"{expr_sql(e.expr)} {'NOT ' if e.negated else ''}LIKE {expr_sql(e.pattern)}"
    if isinstance(e, (A.InSelect, A.Exists, A.SubqueryExpr)):
        n = "NOT " if getattr(e, "negated", False) else ""
        if isinstance(e, A.InSelect):
            return f"{expr_sql(e.expr)} {n}IN (SELECT …)"
        if isinstance(e, A.Exists):
            return f"{n}EXISTS (SELECT …)"
        return "(SELECT …)"
    return repr(e)


def key(e: A.Expr) -> str:
    """Chave canônica (case-insensitive) para comparar expressões."""
    return expr_sql(e).lower()
