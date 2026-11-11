"""AST do VBA (subconjunto usado em bancos Access)."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional


# ------------------------------------------------------------------ expressões
@dataclass
class Expr:
    pass


@dataclass
class Name(Expr):
    id: str


@dataclass
class Member(Expr):
    obj: Optional[Expr]  # None => ".x" dentro de With
    name: str


@dataclass
class Bang(Expr):
    obj: Expr
    name: str


@dataclass
class Arg:
    value: Optional[Expr]  # None => argumento omitido  f(a, , c)
    name: Optional[str] = None  # argumento nomeado  x:=1


@dataclass
class Call(Expr):
    target: Expr
    args: list[Arg]
    parens: bool = True


@dataclass
class Lit(Expr):
    value: Any
    kind: str  # num | str | date | bool | null | nothing | empty


@dataclass
class Unary(Expr):
    op: str
    operand: Expr


@dataclass
class Bin(Expr):
    op: str
    left: Expr
    right: Expr


@dataclass
class New(Expr):
    cls: str


@dataclass
class TypeOf(Expr):
    expr: Expr
    type_name: str


@dataclass
class Me(Expr):
    pass


# ------------------------------------------------------------------ declarações
@dataclass
class VarDecl:
    name: str
    type: Optional[str] = None
    dims: Optional[list[tuple[Optional[Expr], Expr]]] = None  # None = escalar; [] = array dinâmico
    new: bool = False
    is_array: bool = False


@dataclass
class Param:
    name: str
    type: Optional[str] = None
    byval: bool = False
    optional: bool = False
    default: Optional[Expr] = None
    paramarray: bool = False
    is_array: bool = False


# ------------------------------------------------------------------ comandos
@dataclass
class Stmt:
    line: int = field(default=0, kw_only=True)
    comment: Optional[str] = field(default=None, kw_only=True)


@dataclass
class Comment(Stmt):
    text: str


@dataclass
class Dim(Stmt):
    vars: list[VarDecl]
    static: bool = False


@dataclass
class ConstDecl(Stmt):
    items: list[tuple[str, Optional[str], Expr]]


@dataclass
class ReDim(Stmt):
    preserve: bool
    name: Expr
    dims: list[tuple[Optional[Expr], Expr]]
    type: Optional[str] = None


@dataclass
class Assign(Stmt):
    target: Expr
    value: Expr
    is_set: bool = False


@dataclass
class CallStmt(Stmt):
    expr: Expr


@dataclass
class If(Stmt):
    branches: list[tuple[Expr, list[Stmt]]]
    else_body: Optional[list[Stmt]] = None


@dataclass
class CaseClause:
    kind: str  # value | range | is
    a: Expr
    b: Optional[Expr] = None
    op: Optional[str] = None


@dataclass
class SelectCase(Stmt):
    expr: Expr
    cases: list[tuple[list[CaseClause], list[Stmt]]]
    else_body: Optional[list[Stmt]] = None


@dataclass
class For(Stmt):
    var: Expr
    start: Expr
    end: Expr
    step: Optional[Expr]
    body: list[Stmt]


@dataclass
class ForEach(Stmt):
    var: Expr
    iterable: Expr
    body: list[Stmt]


@dataclass
class DoLoop(Stmt):
    body: list[Stmt]
    pre: Optional[tuple[str, Expr]] = None   # ('while'|'until', cond) no início
    post: Optional[tuple[str, Expr]] = None  # no fim


@dataclass
class Exit(Stmt):
    kind: str  # sub | function | property | for | do


@dataclass
class With(Stmt):
    obj: Expr
    body: list[Stmt]


@dataclass
class OnError(Stmt):
    mode: str  # goto | resume_next | goto0
    label: Optional[str] = None


@dataclass
class Label(Stmt):
    name: str


@dataclass
class GoTo(Stmt):
    label: str


@dataclass
class Resume(Stmt):
    target: Optional[str]  # None => Resume, "next" => Resume Next, outro => rótulo


@dataclass
class Erase(Stmt):
    names: list[Expr]


@dataclass
class DebugPrint(Stmt):
    items: list[Expr]


@dataclass
class End(Stmt):
    pass


@dataclass
class RaiseEvent(Stmt):
    name: str
    args: list[Arg]


@dataclass
class Unsupported(Stmt):
    text: str
    reason: str


# ------------------------------------------------------------------ módulo
@dataclass
class Procedure:
    kind: str  # sub | function | property_get | property_let | property_set
    name: str
    params: list[Param]
    body: list[Stmt]
    return_type: Optional[str] = None
    public: bool = True
    static: bool = False
    line: int = 0
    comments_before: list[str] = field(default_factory=list)


@dataclass
class UDT:
    name: str
    fields: list[VarDecl]
    public: bool = True


@dataclass
class EnumDecl:
    name: str
    members: list[tuple[str, Optional[Expr]]]
    public: bool = True


@dataclass
class ModuleVar:
    decl: VarDecl
    public: bool


@dataclass
class Module:
    name: str
    kind: str = "standard"  # standard | class | form | report
    options: dict = field(default_factory=dict)
    variables: list[ModuleVar] = field(default_factory=list)
    constants: list[tuple[str, Optional[str], Expr, bool]] = field(default_factory=list)
    types: list[UDT] = field(default_factory=list)
    enums: list[EnumDecl] = field(default_factory=list)
    procedures: list[Procedure] = field(default_factory=list)
    declares: list[str] = field(default_factory=list)
    header_comments: list[str] = field(default_factory=list)
    unsupported: list[Unsupported] = field(default_factory=list)
