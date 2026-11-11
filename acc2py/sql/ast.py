"""Árvore sintática (AST) do SQL do Access (dialeto Jet/ACE)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


# ---------------------------------------------------------------- expressões
@dataclass
class Expr:
    pass


@dataclass
class Col(Expr):
    """Referência a coluna / parâmetro: ``[t].[x]``, ``x``, ``Forms!F!ctl``."""
    parts: list[str]


@dataclass
class Star(Expr):
    table: Optional[str] = None


@dataclass
class Lit(Expr):
    value: Any
    kind: str  # 'num' | 'str' | 'date' | 'null' | 'bool'


@dataclass
class Unary(Expr):
    op: str  # '-', 'NOT'
    operand: Expr


@dataclass
class Bin(Expr):
    op: str  # '+','-','*','/','\\','MOD','^','&','=','<>','<','>','<=','>=','AND','OR','XOR','EQV','IMP'
    left: Expr
    right: Expr


@dataclass
class Func(Expr):
    name: str
    args: list[Expr]
    distinct: bool = False


@dataclass
class Between(Expr):
    expr: Expr
    low: Expr
    high: Expr
    negated: bool = False


@dataclass
class InList(Expr):
    expr: Expr
    items: list[Expr]
    negated: bool = False


@dataclass
class InSelect(Expr):
    expr: Expr
    query: "Statement"
    negated: bool = False


@dataclass
class Exists(Expr):
    query: "Statement"
    negated: bool = False


@dataclass
class IsNull(Expr):
    expr: Expr
    negated: bool = False


@dataclass
class Like(Expr):
    expr: Expr
    pattern: Expr
    negated: bool = False


@dataclass
class SubqueryExpr(Expr):
    query: "Statement"


# ---------------------------------------------------------------- FROM
@dataclass
class Source:
    pass


@dataclass
class TableRef(Source):
    name: str
    alias: Optional[str] = None
    database: Optional[str] = None  # IN 'c:\outro.mdb'

    @property
    def key(self) -> str:
        return self.alias or self.name


@dataclass
class SubSource(Source):
    query: "Statement"
    alias: Optional[str] = None

    @property
    def key(self) -> str:
        return self.alias or "_sub"


@dataclass
class Join(Source):
    kind: str  # 'INNER' | 'LEFT' | 'RIGHT' | 'CROSS'
    left: Source
    right: Source
    on: Optional[Expr] = None


# ---------------------------------------------------------------- comandos
@dataclass
class Statement:
    parameters: list[tuple[str, str]] = field(default_factory=list, kw_only=True)


@dataclass
class Select(Statement):
    columns: list[tuple[Expr, Optional[str]]] = field(default_factory=list)
    distinct: str = ""  # '' | 'DISTINCT' | 'DISTINCTROW'
    top: Optional[float] = None
    top_percent: bool = False
    from_: Optional[Source] = None
    where: Optional[Expr] = None
    group_by: list[Expr] = field(default_factory=list)
    having: Optional[Expr] = None
    order_by: list[tuple[Expr, bool]] = field(default_factory=list)  # (expr, desc)
    into: Optional[str] = None  # SELECT ... INTO nova_tabela


@dataclass
class Union(Statement):
    parts: list[Select]
    all_flags: list[bool]  # all_flags[i] -> UNION ALL entre parts[i] e parts[i+1]
    order_by: list[tuple[Expr, bool]] = field(default_factory=list)


@dataclass
class Transform(Statement):
    value: Expr  # agregação do TRANSFORM
    select: Select
    pivot: Expr
    pivot_in: list[Expr] = field(default_factory=list)


@dataclass
class Update(Statement):
    source: Source
    sets: list[tuple[Col, Expr]]
    where: Optional[Expr] = None


@dataclass
class Delete(Statement):
    target: Optional[str]
    source: Source
    where: Optional[Expr] = None


@dataclass
class Insert(Statement):
    table: str
    columns: list[str]
    select: Optional[Statement] = None
    values: Optional[list[Expr]] = None


AGGREGATES = {"sum", "count", "avg", "min", "max", "first", "last",
              "stdev", "stdevp", "var", "varp"}


def walk(e):
    """Percorre recursivamente uma expressão (sem entrar em subconsultas)."""
    if e is None:
        return
    yield e
    if isinstance(e, Unary):
        yield from walk(e.operand)
    elif isinstance(e, Bin):
        yield from walk(e.left)
        yield from walk(e.right)
    elif isinstance(e, Func):
        for a in e.args:
            yield from walk(a)
    elif isinstance(e, Between):
        yield from walk(e.expr)
        yield from walk(e.low)
        yield from walk(e.high)
    elif isinstance(e, InList):
        yield from walk(e.expr)
        for i in e.items:
            yield from walk(i)
    elif isinstance(e, (InSelect,)):
        yield from walk(e.expr)
    elif isinstance(e, (IsNull, )):
        yield from walk(e.expr)
    elif isinstance(e, Like):
        yield from walk(e.expr)
        yield from walk(e.pattern)


def has_aggregate(e) -> bool:
    return any(isinstance(n, Func) and n.name.lower() in AGGREGATES for n in walk(e))
