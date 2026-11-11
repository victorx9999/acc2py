"""Gera código Python a partir da AST do VBA.

Decisões principais (detalhadas no README):
* Nomes mantêm a grafia da declaração; a resolução é sem diferenciar maiúsculas.
* ``Function F`` usa uma variável ``_ret``; ``Exit Function`` vira ``return _ret``.
* Comparações de texto com ``Option Compare Database/Text`` usam ``vb.Eq``/``vb.Lt``…
  (sem diferenciar maiúsculas, como no Access).
* ``On Error GoTo Rotulo`` no padrão clássico (corpo / saída / tratador) vira
  ``try/except``; ``On Error Resume Next`` envolve cada comando seguinte.
* ByRef: o Python não devolve alterações de parâmetros; isso é sinalizado.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

from . import ast as A
from ..diagnostics import Diagnostics
from ..runtime.vba import BUILTINS, CONSTANTS, OBJECTS

PY_RESERVED = {"False", "None", "True", "and", "as", "assert", "async", "await", "break",
               "class", "continue", "def", "del", "elif", "else", "except", "finally", "for",
               "from", "global", "if", "import", "in", "is", "lambda", "nonlocal", "not", "or",
               "pass", "raise", "return", "try", "while", "with", "yield", "match", "case",
               "vb", "datetime", "Decimal", "range", "self", "_ret", "print", "type"}

TYPE_HINT = {"integer": "int", "long": "int", "byte": "int", "longlong": "int", "longptr": "int",
             "double": "float", "single": "float", "string": "str", "boolean": "bool",
             "date": "datetime", "currency": "float", "decimal": "Decimal"}
TYPE_DEFAULT = {"integer": "0", "long": "0", "byte": "0", "longlong": "0", "longptr": "0",
                "double": "0.0", "single": "0.0", "string": '""', "boolean": "False",
                "date": "datetime(1899, 12, 30)", "currency": "0.0", "decimal": 'Decimal("0")'}
NEW_CLASSES = {"collection": "vb.Collection", "dictionary": "vb.Dictionary",
               "scripting.dictionary": "vb.Dictionary", "filesystemobject": "vb.FileSystemObject",
               "scripting.filesystemobject": "vb.FileSystemObject"}
NOARG_BUILTINS = {"date", "now", "time", "timer", "currentdb", "codedb", "rnd", "curdir"}
NOARG_METHODS = {"movenext", "moveprevious", "movefirst", "movelast", "close", "clone", "requery",
                 "update", "edit", "addnew", "delete", "cancelupdate", "refresh", "undo",
                 "removeall", "clear", "setfocus", "keys", "items"}
STRING_FUNCS = {"left", "right", "mid", "trim", "ltrim", "rtrim", "ucase", "lcase", "format",
                "cstr", "replace", "space", "string", "strreverse", "chr", "chrw", "str", "join"}
CMP_OPS = {"=": ("==", "Eq"), "<>": ("!=", "Ne"), "<": ("<", "Lt"), ">": (">", "Gt"),
           "<=": ("<=", "Le"), ">=": (">=", "Ge")}


def py_name(name: str) -> str:
    s = name
    if not s.isidentifier():
        s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
        s = re.sub(r"\W+", "_", s).strip("_") or "x"
        if s[0].isdigit():
            s = "_" + s
    if s in PY_RESERVED or s.lower() in {"vb", "self", "range"}:
        s += "_"
    return s


def py_str(s: str) -> str:
    out = s.replace("\\", "\\\\").replace('"', '\\"').replace("\r", "\\r").replace("\n", "\\n") \
        .replace("\t", "\\t").replace("\0", "\\0")
    return f'"{out}"'


# ======================================================================= símbolos
@dataclass
class Sym:
    name: str          # grafia original
    kind: str          # var | const | proc | enum | enummember | type | class | module
    module: str
    public: bool = True
    is_array: bool = False
    vtype: Optional[str] = None
    params: list = field(default_factory=list)
    proc_kind: str = ""


class Project:
    """Símbolos de todos os módulos (para resolver chamadas entre módulos)."""

    def __init__(self, modules: list[A.Module]):
        self.modules = {m.name.lower(): m for m in modules}
        self.public: dict[str, Sym] = {}
        self.module_syms: dict[str, dict[str, Sym]] = {}
        for m in modules:
            syms: dict[str, Sym] = {}
            is_cls = m.kind != "standard"
            for v in m.variables:
                syms[v.decl.name.lower()] = Sym(v.decl.name, "var", m.name, v.public,
                                                v.decl.is_array, v.decl.type)
            for name, typ, _, public in m.constants:
                syms[name.lower()] = Sym(name, "const", m.name, public, vtype=typ)
            for e in m.enums:
                syms[e.name.lower()] = Sym(e.name, "enum", m.name, e.public)
                for mn, _ in e.members:
                    syms.setdefault(mn.lower(), Sym(mn, "enummember", m.name, e.public))
            for t in m.types:
                syms[t.name.lower()] = Sym(t.name, "type", m.name, t.public)
            for p in m.procedures:
                key = p.name.lower()
                if key in syms and syms[key].kind == "proc":
                    syms[key].proc_kind += "," + p.kind
                    if p.kind in ("property_get", "function"):
                        syms[key].params = p.params
                    continue
                syms[key] = Sym(p.name, "proc", m.name, p.public, vtype=p.return_type,
                                params=p.params, proc_kind=p.kind)
            self.module_syms[m.name.lower()] = syms
            if is_cls:
                self.public[m.name.lower()] = Sym(m.name, "class", m.name)
            else:
                self.public[m.name.lower()] = Sym(m.name, "module", m.name)
                for k, s in syms.items():
                    if s.public:
                        self.public.setdefault(k, s)


# ======================================================================= gerador
class Emitter:
    def __init__(self):
        self.lines: list[str] = []
        self.level = 0

    def emit(self, line=""):
        self.lines.append(("    " * self.level + line) if line else "")

    def block(self):
        return _Indent(self)


class _Indent:
    def __init__(self, em):
        self.em = em

    def __enter__(self):
        self.em.level += 1

    def __exit__(self, *a):
        self.em.level -= 1


@dataclass
class ProcCtx:
    proc: Optional[A.Procedure]
    locals: dict = field(default_factory=dict)   # lower -> Sym
    globals_used: set = field(default_factory=set)
    with_stack: list = field(default_factory=list)
    loop_stack: list = field(default_factory=list)  # (kind, continue_label)
    tmp: int = 0
    exit_label: Optional[str] = None
    in_try: bool = False
    resume_next: bool = False


class ModuleGenerator:
    def __init__(self, module: A.Module, project: Project, diags: Diagnostics):
        self.m = module
        self.project = project
        self.diags = diags
        self.syms = project.module_syms[module.name.lower()]
        self.is_class = module.kind != "standard"
        self.is_form = module.kind in ("form", "report")
        self.compare_text = module.options.get("compare", "binary").lower() in ("database", "text")
        self.em = Emitter()
        self.imports: set[str] = set()
        self.ctx = ProcCtx(None)
        self.statics: list[str] = []

    # ------------------------------------------------------------ util
    def where(self, line=None):
        p = self.ctx.proc
        loc = f"{self.m.name}" + (f".{p.name}" if p else "")
        return loc + (f" (linha {line})" if line else "")

    def warn(self, msg, line=None):
        self.diags.warn(self.where(line), msg)

    def info(self, msg, line=None):
        self.diags.info(self.where(line), msg)

    def error(self, msg, line=None):
        self.diags.error(self.where(line), msg)

    def tmp(self, base):
        self.ctx.tmp += 1
        return f"_{base}{self.ctx.tmp}"

    # ------------------------------------------------------------ módulo
    def generate(self) -> str:
        m = self.m
        em = self.em
        body = Emitter()
        self.em = body
        # tipos, enums e constantes (nível de módulo)
        for t in m.types:
            self.gen_type(t)
        for e in m.enums:
            self.gen_enum(e)
        for name, typ, val, public in m.constants:
            self.em.emit(f"{py_name(name)} = {self.ex(val)}")
        if m.constants:
            self.em.emit("")
        if self.is_class:
            self.gen_class()
        else:
            for v in m.variables:
                self.em.emit(self.decl_line(v.decl, prefix=""))
            if m.variables:
                self.em.emit("")
            for p in m.procedures:
                self.gen_procedure(p)
        for d in m.declares:
            self.warn(f"Declare (API do Windows) não convertido: {d[:80]}")
        for u in m.unsupported:
            self.warn(f"{u.reason}: {u.text[:80]}", u.line)
        # cabeçalho
        self.em = em
        em.emit(f'"""Módulo VBA "{m.name}" transpilado por acc2py.')
        em.emit("")
        em.emit("Procure por ATENÇÃO para os pontos que precisam de revisão manual.")
        em.emit('"""')
        for c in m.header_comments:
            em.emit("# " + c.strip())
        em.emit("from datetime import datetime")
        em.emit("from decimal import Decimal")
        em.emit("from dataclasses import dataclass, field")
        em.emit("")
        em.emit("from acc2py.runtime import vba as vb")
        for mod in sorted(self.imports):
            em.emit(f"from . import {py_name(mod)}")
        if self.statics:
            em.emit("")
            for s in self.statics:
                em.emit(s)
        em.emit("")
        em.lines.extend(body.lines)
        while em.lines and not em.lines[-1].strip():
            em.lines.pop()
        return "\n".join(em.lines) + "\n"

    def gen_type(self, t: A.UDT):
        em = self.em
        em.emit("@dataclass")
        em.emit(f"class {py_name(t.name)}:  # Type do VBA")
        with em.block():
            if not t.fields:
                em.emit("pass")
            for f in t.fields:
                hint = TYPE_HINT.get((f.type or "").lower(), "object")
                default = self.default_value(f)
                if "(" in default or default.startswith("vb."):
                    default = f"field(default_factory=lambda: {default})"
                    self.need_field = True
                em.emit(f"{py_name(f.name)}: {hint} = {default}")
        em.emit("")
        em.emit("")

    def gen_enum(self, e: A.EnumDecl):
        em = self.em
        em.emit(f"class {py_name(e.name)}:  # Enum do VBA")
        val = -1
        names = []
        with em.block():
            for mn, mv in e.members:
                if mv is not None:
                    txt = self.ex(mv)
                    try:
                        val = int(eval(txt, {}))
                        em.emit(f"{py_name(mn)} = {txt}")
                    except Exception:
                        em.emit(f"{py_name(mn)} = {txt}")
                else:
                    val += 1
                    em.emit(f"{py_name(mn)} = {val}")
                names.append(mn)
        em.emit("")
        for mn in names:
            em.emit(f"{py_name(mn)} = {py_name(e.name)}.{py_name(mn)}")
        em.emit("")

    def default_value(self, d: A.VarDecl) -> str:
        if d.is_array:
            if d.dims:
                base = self.m.options.get("base", "0")
                bounds = ", ".join(f"({self.ex(lo) if lo else base}, {self.ex(hi)})" for lo, hi in d.dims)
                dflt = TYPE_DEFAULT.get((d.type or "").lower())
                return f"vb.DimArray({bounds}" + (f", default={dflt})" if dflt else ")")
            return "vb.VBArray()"
        if d.new and d.type:
            return self.new_expr(d.type) + "()"
        if d.type and d.type.lower() in TYPE_DEFAULT:
            return TYPE_DEFAULT[d.type.lower()]
        if d.type:
            sym = self.lookup_global(d.type)
            if sym and sym.kind == "type":
                return self.sym_ref(sym) + "()"
        return "None"

    def decl_line(self, d: A.VarDecl, prefix="") -> str:
        hint = TYPE_HINT.get((d.type or "").lower())
        val = self.default_value(d)
        name = prefix + py_name(d.name)
        if hint and not d.is_array and not prefix:
            return f"{name}: {hint} = {val}"
        return f"{name} = {val}"

    def new_expr(self, type_name: str) -> str:
        low = type_name.lower()
        if low in NEW_CLASSES:
            return NEW_CLASSES[low]
        sym = self.lookup_global(type_name.split(".")[-1])
        if sym and sym.kind in ("class", "type"):
            return self.sym_ref(sym)
        self.warn(f"New {type_name}: classe desconhecida")
        return py_name(type_name.split(".")[-1])

    # ------------------------------------------------------------ classe
    def gen_class(self):
        m, em = self.m, self.em
        base = "(vb.FormBase)" if self.is_form else ""
        em.emit(f"class {py_name(m.name)}{base}:")
        with em.block():
            em.emit(f'"""Módulo de {"formulário" if self.is_form else "classe"} {m.name}."""')
            em.emit("")
            em.emit("def __init__(self, **controles):" if self.is_form else "def __init__(self):")
            with em.block():
                if self.is_form:
                    em.emit("super().__init__(**controles)")
                for v in m.variables:
                    em.emit(self.decl_line(v.decl, prefix="self."))
                if any(p.name.lower() == "class_initialize" for p in m.procedures):
                    em.emit("self.Class_Initialize()")
                if not m.variables and not self.is_form and not any(
                        p.name.lower() == "class_initialize" for p in m.procedures):
                    em.emit("pass")
            em.emit("")
            props: dict[str, dict[str, A.Procedure]] = {}
            for p in m.procedures:
                if p.kind.startswith("property") and not (p.kind == "property_get" and p.params) \
                        and not (p.kind != "property_get" and len(p.params) > 1):
                    props.setdefault(p.name.lower(), {})[p.kind] = p
                    continue
                self.gen_procedure(p, method=True)
            for key, group in props.items():
                getter = group.get("property_get")
                setter = group.get("property_let") or group.get("property_set")
                if getter:
                    em.emit("@property")
                    self.gen_procedure(getter, method=True)
                else:
                    name = (setter.name if setter else key)
                    em.emit("@property")
                    em.emit(f"def {py_name(name)}(self):")
                    with em.block():
                        em.emit(f'raise AttributeError("Propriedade {name} só tem Let/Set")')
                    em.emit("")
                if setter:
                    em.emit(f"@{py_name(setter.name)}.setter")
                    self.gen_procedure(setter, method=True)

    # ------------------------------------------------------------ procedimentos
    def gen_procedure(self, p: A.Procedure, method=False):
        em = self.em
        self.ctx = ProcCtx(p)
        ctx = self.ctx
        for prm in p.params:
            ctx.locals[prm.name.lower()] = Sym(prm.name, "param", self.m.name, is_array=prm.is_array, vtype=prm.type)
        self.collect_locals(p.body)
        params = ["self"] if method else []
        for prm in p.params:
            n = py_name(prm.name)
            if prm.paramarray:
                params.append("*" + n)
            elif prm.optional or prm.default is not None:
                params.append(f"{n}={self.ex(prm.default) if prm.default is not None else 'None'}")
            else:
                hint = TYPE_HINT.get((prm.type or "").lower())
                params.append(f"{n}: {hint}" if hint else n)
        ret = ""
        if p.kind in ("function", "property_get"):
            hint = TYPE_HINT.get((p.return_type or "").lower())
            ret = f" -> {hint}" if hint else ""
        for c in p.comments_before:
            em.emit("# " + c.strip())
        em.emit(f"def {py_name(p.name)}({', '.join(params)}){ret}:")
        body_em = Emitter()
        body_em.level = em.level + 1
        outer = self.em
        self.em = body_em
        # avisos de ByRef
        assigned = self.assigned_names(p.body)
        for prm in p.params:
            if not prm.byval and not prm.paramarray and prm.name.lower() in assigned:
                body_em.emit(f"# ATENÇÃO: parâmetro ByRef '{prm.name}' é alterado aqui; em Python a "
                             f"alteração NÃO volta para quem chamou.")
                self.warn(f"parâmetro ByRef '{prm.name}' é alterado; quem chama não verá a mudança "
                          f"(devolva o valor ou use um objeto mutável)", p.line)
        if p.kind in ("function", "property_get"):
            body_em.emit(f"_ret = {TYPE_DEFAULT.get((p.return_type or '').lower(), 'None')}")
        self.emit_body(p.body, top=True)
        if p.kind in ("function", "property_get"):
            body_em.emit("return _ret")
        if not any(ln.strip() and not ln.strip().startswith("#") for ln in body_em.lines):
            body_em.emit("pass")
        if ctx.globals_used and not method:
            em.emit("    global " + ", ".join(sorted(ctx.globals_used)))
        self.em = outer
        em.lines.extend(body_em.lines)
        em.emit("")
        if not method:
            em.emit("")
        if p.name.lower() == "class_terminate":
            self.info("Class_Terminate vira método comum (não é chamado automaticamente)")

    def collect_locals(self, body):
        for st in self.iter_stmts(body):
            if isinstance(st, A.Dim):
                for d in st.vars:
                    self.ctx.locals[d.name.lower()] = Sym(d.name, "var", self.m.name, is_array=d.is_array, vtype=d.type)
                    if st.static:
                        self.ctx.locals[d.name.lower()].kind = "static"
            elif isinstance(st, A.ConstDecl):
                for name, typ, _ in st.items:
                    self.ctx.locals[name.lower()] = Sym(name, "const", self.m.name, vtype=typ)
            elif isinstance(st, A.ReDim) and isinstance(st.name, A.Name):
                key = st.name.id.lower()
                if key not in self.ctx.locals and not self.syms.get(key):
                    self.ctx.locals[key] = Sym(st.name.id, "var", self.m.name, is_array=True, vtype=st.type)
                elif key in self.ctx.locals:
                    self.ctx.locals[key].is_array = True

    def iter_stmts(self, body):
        for st in body:
            yield st
            for sub in self.children(st):
                yield from self.iter_stmts(sub)

    @staticmethod
    def children(st):
        if isinstance(st, A.If):
            return [b for _, b in st.branches] + ([st.else_body] if st.else_body else [])
        if isinstance(st, A.SelectCase):
            return [b for _, b in st.cases] + ([st.else_body] if st.else_body else [])
        if isinstance(st, (A.For, A.ForEach, A.DoLoop, A.With)):
            return [st.body]
        return []

    def assigned_names(self, body) -> set:
        out = set()
        for st in self.iter_stmts(body):
            if isinstance(st, A.Assign) and isinstance(st.target, A.Name):
                out.add(st.target.id.lower())
            elif isinstance(st, A.For) and isinstance(st.var, A.Name):
                out.add(st.var.id.lower())
            elif isinstance(st, A.ReDim) and isinstance(st.name, A.Name):
                out.add(st.name.id.lower())
        return out

    # ------------------------------------------------------------ corpo
    def emit_body(self, body: list[A.Stmt], top=False):
        if top:
            body = self.structure_on_error(body)
        if not body:
            return
        for i, st in enumerate(body):
            if self.ctx.resume_next and not isinstance(st, (A.Comment, A.OnError, A.Label)):
                self.em.emit("try:")
                with self.em.block():
                    self.stmt(st)
                self.em.emit("except Exception as _e:")
                with self.em.block():
                    self.em.emit("vb.Err.set(_e)  # On Error Resume Next")
                continue
            self.stmt(st)

    def emit_block(self, body):
        start = len(self.em.lines)
        self.emit_body(body)
        if not any(ln.strip() and not ln.strip().startswith("#") for ln in self.em.lines[start:]):
            self.em.emit("pass")

    def structure_on_error(self, body: list[A.Stmt]) -> list[A.Stmt]:
        """Reconhece o padrão On Error GoTo Erro / Sair: / Exit Sub / Erro: … Resume Sair."""
        idx = next((i for i, s in enumerate(body) if isinstance(s, A.OnError) and s.mode == "goto"), None)
        if idx is None:
            return body
        label = body[idx].label.lower()
        j = next((k for k, s in enumerate(body) if isinstance(s, A.Label) and s.name.lower() == label and k > idx), None)
        if j is None:
            self.warn(f"On Error GoTo {body[idx].label}: rótulo não encontrado no nível principal", body[idx].line)
            return body
        k = next((x for x in range(idx + 1, j) if isinstance(body[x], A.Label)), None)
        try_body = body[idx + 1: k if k is not None else j]
        exit_sec = body[k + 1: j] if k is not None else []
        exit_label = body[k].name if k is not None else None
        handler = body[j + 1:]
        if exit_sec and isinstance(exit_sec[-1], A.Exit):
            exit_sec = exit_sec[:-1]
        elif not exit_sec and try_body and isinstance(try_body[-1], A.Exit):
            try_body = try_body[:-1]
        meaningful = [s for s in handler if not isinstance(s, A.Comment)]
        falls_through = False
        if meaningful and isinstance(meaningful[-1], A.Resume) and exit_label \
                and (meaningful[-1].target or "").lower() == exit_label.lower():
            handler = [s for s in handler if s is not meaningful[-1]]
            falls_through = True
        for s in self.iter_stmts(handler):
            if isinstance(s, A.Resume):
                self.warn(f"'Resume {s.target or ''}' dentro do tratador de erro não tem equivalente "
                          f"direto; revisado manualmente", s.line)
        # declarações (Dim) saem do try para valerem também no tratador e na saída
        dims = [s for s in try_body if isinstance(s, (A.Dim, A.ConstDecl))]
        try_body = [s for s in try_body if not isinstance(s, (A.Dim, A.ConstDecl))]
        return body[:idx] + dims + [_TryBlock(try_body, handler, exit_sec, exit_label, falls_through,
                                              line=body[idx].line)]

    # ------------------------------------------------------------ comandos
    def stmt(self, st: A.Stmt):
        em = self.em
        tail = f"  # {st.comment.strip()}" if st.comment else ""
        if isinstance(st, A.Comment):
            for ln in st.text.splitlines() or [""]:
                em.emit(("# " + ln.strip()).rstrip())
            return
        if isinstance(st, _TryBlock):
            return self.try_block(st)
        if isinstance(st, A.Dim):
            for d in st.vars:
                if st.static:
                    gname = f"_static_{py_name(self.ctx.proc.name)}_{py_name(d.name)}"
                    self.statics.append(f"{gname} = {self.default_value(d)}  # Static em {self.ctx.proc.name}")
                    self.ctx.locals[d.name.lower()].name = gname
                    self.ctx.locals[d.name.lower()].kind = "static"
                    self.ctx.globals_used.add(gname)
                    if self.is_class:
                        self.warn("Static em método de classe vira variável global do módulo", st.line)
                    continue
                em.emit(self.decl_line(d) + tail)
            return
        if isinstance(st, A.ConstDecl):
            for name, _, val in st.items:
                em.emit(f"{py_name(name)} = {self.ex(val)}{tail}")
            return
        if isinstance(st, A.ReDim):
            target = self.target(st.name)
            base = self.m.options.get("base", "0")
            bounds = ", ".join(f"({self.ex(lo) if lo else base}, {self.ex(hi)})" for lo, hi in st.dims)
            dflt = TYPE_DEFAULT.get((st.type or "").lower())
            extra = (", preserve=True" if st.preserve else "") + (f", default={dflt}" if dflt else "")
            em.emit(f"{target} = vb.ReDim({target}, {bounds}{extra}){tail}")
            return
        if isinstance(st, A.Assign):
            return self.assign(st, tail)
        if isinstance(st, A.CallStmt):
            em.emit(self.call_stmt(st.expr) + tail)
            return
        if isinstance(st, A.If):
            for n, (cond, body) in enumerate(st.branches):
                em.emit(f"{'if' if n == 0 else 'elif'} {self.cond(cond)}:" + (tail if n == 0 else ""))
                with em.block():
                    self.emit_block(body)
            if st.else_body:
                em.emit("else:")
                with em.block():
                    self.emit_block(st.else_body)
            return
        if isinstance(st, A.SelectCase):
            return self.select_case(st, tail)
        if isinstance(st, A.For):
            return self.for_loop(st, tail)
        if isinstance(st, A.ForEach):
            var = self.target(st.var)
            em.emit(f"for {var} in {self.ex(st.iterable)}:{tail}")
            self.loop_body(st.body, "for")
            return
        if isinstance(st, A.DoLoop):
            return self.do_loop(st, tail)
        if isinstance(st, A.Exit):
            return self.exit_stmt(st, tail)
        if isinstance(st, A.With):
            obj = st.obj
            if isinstance(obj, (A.Name, A.Me)):
                self.ctx.with_stack.append(self.ex(obj, as_value=False))
            else:
                w = self.tmp("w")
                em.emit(f"{w} = {self.ex(obj, as_value=False)}{tail}")
                self.ctx.with_stack.append(w)
            self.emit_body(st.body)
            self.ctx.with_stack.pop()
            return
        if isinstance(st, A.OnError):
            if st.mode == "resume_next":
                self.ctx.resume_next = True
                em.emit("vb.Err.Clear()  # On Error Resume Next: cada comando abaixo ignora erros")
                self.info("On Error Resume Next: cada comando seguinte foi envolvido em try/except", st.line)
            elif st.mode == "goto0":
                self.ctx.resume_next = False
                em.emit("# On Error GoTo 0")
            else:
                em.emit(f"# ATENÇÃO: On Error GoTo {st.label} fora do padrão reconhecido")
                self.warn(f"On Error GoTo {st.label} fora do padrão corpo/saída/tratador", st.line)
            return
        if isinstance(st, A.Label):
            em.emit(f"# {st.name}:")
            return
        if isinstance(st, A.GoTo):
            return self.goto(st, tail)
        if isinstance(st, A.Resume):
            em.emit(f"# ATENÇÃO: Resume {st.target or ''} não convertido")
            return
        if isinstance(st, A.Erase):
            for n in st.names:
                t = self.target(n)
                em.emit(f"{t} = vb.Erase({t}){tail}")
            return
        if isinstance(st, A.DebugPrint):
            em.emit(f"vb.DebugPrint({', '.join(self.ex(i) for i in st.items)}){tail}")
            return
        if isinstance(st, A.End):
            em.emit("raise SystemExit  # End")
            return
        if isinstance(st, A.RaiseEvent):
            em.emit(f"# ATENÇÃO: RaiseEvent {st.name} não convertido")
            self.warn(f"RaiseEvent {st.name} não convertido (eventos de classe)", st.line)
            return
        if isinstance(st, A.Unsupported):
            em.emit(f"# NÃO CONVERTIDO: {st.text}")
            em.emit(f"raise NotImplementedError({py_str(st.reason + ': ' + st.text[:60])})")
            self.error(f"{st.reason}: {st.text[:80]}", st.line)
            return
        raise NotImplementedError(type(st).__name__)

    def try_block(self, tb: "_TryBlock"):
        em = self.em
        prev = self.ctx.exit_label
        self.ctx.exit_label = tb.exit_label
        em.emit("try:")
        with em.block():
            self.emit_block(tb.body)
        needs_goto = any(isinstance(s, A.GoTo) and tb.exit_label and s.label.lower() == tb.exit_label.lower()
                         for s in self.iter_stmts(tb.body))
        if needs_goto:
            em.emit("except vb.GoToExit:")
            with em.block():
                em.emit(f"pass  # GoTo {tb.exit_label}")
        em.emit("except Exception as _e:")
        self.ctx.exit_label = prev
        with em.block():
            em.emit("vb.Err.set(_e)")
            self.emit_body(tb.handler)
            if not tb.falls_through and tb.exit_sec:
                em.emit(self._return())
        if tb.exit_sec:
            em.emit(f"# {tb.exit_label}:" if tb.exit_label else "# saída")
            self.emit_body(tb.exit_sec)

    def _return(self) -> str:
        p = self.ctx.proc
        return "return _ret" if p and p.kind in ("function", "property_get") else "return"

    def goto(self, st: A.GoTo, tail):
        lab = st.label.lower()
        for kind, cont in reversed(self.ctx.loop_stack):
            if cont and cont.lower() == lab:
                self.em.emit(f"continue  # GoTo {st.label}{tail}")
                return
        if self.ctx.exit_label and lab == self.ctx.exit_label.lower():
            self.em.emit(f"raise vb.GoToExit()  # GoTo {st.label}{tail}")
            return
        p = self.ctx.proc
        if p:
            body = p.body
            idx = next((i for i, s in enumerate(body) if isinstance(s, A.Label) and s.name.lower() == lab), None)
            if idx is not None and all(isinstance(s, (A.Exit, A.Comment)) for s in body[idx + 1:]):
                self.em.emit(f"{self._return()}  # GoTo {st.label}{tail}")
                return
        self.em.emit(f"raise NotImplementedError(\"GoTo {st.label}\")  # ATENÇÃO: GoTo não convertido")
        self.error(f"GoTo {st.label} não convertido (reestruture o fluxo)", st.line)

    def exit_stmt(self, st: A.Exit, tail):
        k = st.kind
        if k in ("sub", "function", "property"):
            if self.ctx.exit_label and self.ctx.in_try:
                pass
            self.em.emit(self._return() + tail)
            return
        want = "for" if k == "for" else "do"
        stack = self.ctx.loop_stack
        if stack and stack[-1][0] != want:
            self.em.emit(f"break  # ATENÇÃO: Exit {k.capitalize()} saindo de laço interno de outro tipo{tail}")
            self.warn(f"Exit {k.capitalize()} dentro de laço aninhado de outro tipo: confira o break", st.line)
        else:
            self.em.emit("break" + tail)

    def loop_body(self, body, kind):
        cont = None
        if body:
            last = [s for s in body if not isinstance(s, A.Comment)]
            if last and isinstance(last[-1], A.Label):
                cont = last[-1].name
        self.ctx.loop_stack.append((kind, cont))
        with self.em.block():
            self.emit_block(body)
        self.ctx.loop_stack.pop()

    def for_loop(self, st: A.For, tail):
        var = self.target(st.var)
        a, b = self.ex(st.start), self.ex(st.end)
        step = st.step
        vsym = self.resolve(st.var.id) if isinstance(st.var, A.Name) else None
        float_var = vsym is not None and (vsym.vtype or "").lower() in ("double", "single", "currency")
        if not float_var and (step is None or self._int_lit(step) is not None):
            s = 1 if step is None else self._int_lit(step)
            if not (self._looks_int(st.start) and self._looks_int(st.end)):
                a, b = f"vb.CInt({_strip_parens(a)})", f"vb.CInt({_strip_parens(b)})"
            stop = self._plus(b, 1 if s > 0 else -1)
            rng = f"range({a}, {stop})" if s == 1 else f"range({a}, {stop}, {s})"
        else:
            rng = f"vb.ForRange({a}, {b}, {self.ex(step) if step else 1})"
        self.em.emit(f"for {var} in {rng}:{tail}")
        self.loop_body(st.body, "for")

    @staticmethod
    def _int_lit(e):
        if isinstance(e, A.Lit) and isinstance(e.value, int):
            return e.value
        if isinstance(e, A.Unary) and e.op == "-" and isinstance(e.operand, A.Lit) and isinstance(e.operand.value, int):
            return -e.operand.value
        return None

    def _looks_int(self, e) -> bool:
        if self._int_lit(e) is not None:
            return True
        if isinstance(e, A.Name):
            s = self.resolve(e.id)
            return s is not None and (s.vtype or "").lower() in ("integer", "long", "byte")
        if isinstance(e, A.Call) and isinstance(e.target, A.Name) and e.target.id.lower() in (
                "ubound", "lbound", "len", "clng", "cint", "int", "fix", "instr", "year", "month", "day"):
            return True
        if isinstance(e, A.Member) and e.name.lower() in ("count", "recordcount"):
            return True
        if isinstance(e, A.Bin) and e.op in ("+", "-", "*"):
            return self._looks_int(e.left) and self._looks_int(e.right)
        return False

    @staticmethod
    def _plus(expr: str, n: int) -> str:
        m = re.fullmatch(r"\((.*) - 1\)", expr)
        if n == 1 and m:
            return m.group(1)
        if re.fullmatch(r"-?\d+", expr):
            return str(int(expr) + n)
        return f"{expr} + {n}" if n > 0 else f"{expr} - {-n}"

    def do_loop(self, st: A.DoLoop, tail):
        em = self.em
        if st.pre:
            kind, cond = st.pre
            c = self.cond(cond)
            em.emit(f"while {c if kind == 'while' else self._negate(c)}:{tail}")
            self.loop_body(st.body, "do")
        else:
            em.emit(f"while True:{tail}")
            self.ctx.loop_stack.append(("do", None))
            with em.block():
                self.emit_body(st.body)
                if st.post:
                    kind, cond = st.post
                    c = self.cond(cond)
                    em.emit(f"if {self._negate(c) if kind == 'while' else c}:")
                    with em.block():
                        em.emit("break")
                elif not st.body:
                    em.emit("pass")
            self.ctx.loop_stack.pop()

    @staticmethod
    def _negate(c: str) -> str:
        if c.startswith("not "):
            inner = c[4:]
            if not re.search(r"\b(and|or)\b", inner):
                return inner
        return f"not ({c})" if " " in c else f"not {c}"

    def select_case(self, st: A.SelectCase, tail):
        em = self.em
        if isinstance(st.expr, (A.Name, A.Lit)):
            subj = self.ex(st.expr)
        else:
            subj = self.tmp("sel")
            em.emit(f"{subj} = {self.ex(st.expr)}{tail}")
            tail = ""
        stringy = self.stringy(st.expr) or any(
            isinstance(c.a, A.Lit) and c.a.kind == "str" for cl, _ in st.cases for c in cl)
        first = True
        for clauses, body in st.cases:
            parts = []
            for c in clauses:
                if c.kind == "value":
                    parts.append(self.compare("=", subj, self.ex(c.a), stringy))
                elif c.kind == "range":
                    if stringy and self.compare_text:
                        parts.append(f"vb.Ge({subj}, {self.ex(c.a)}) and vb.Le({subj}, {self.ex(c.b)})")
                    else:
                        parts.append(f"{self.ex(c.a)} <= {subj} <= {self.ex(c.b)}")
                else:
                    parts.append(self.compare(c.op, subj, self.ex(c.a), stringy))
            if len(parts) > 1 and all(c.kind == "value" for c in clauses) and not (stringy and self.compare_text):
                cond = f"{subj} in ({', '.join(self.ex(c.a) for c in clauses)})"
            else:
                cond = " or ".join(f"({p})" if " or " in p or " and " in p else p for p in parts)
            em.emit(f"{'if' if first else 'elif'} {cond}:" + (tail if first else ""))
            first = False
            with em.block():
                self.emit_block(body)
        if st.else_body is not None:
            if first:
                self.emit_body(st.else_body)
            else:
                em.emit("else:")
                with em.block():
                    self.emit_block(st.else_body)

    def compare(self, op, a, b, stringy) -> str:
        pyop, fn = CMP_OPS[op]
        if stringy and self.compare_text:
            return f"vb.{fn}({a}, {b})"
        return f"{a} {pyop} {b}"

    # ------------------------------------------------------------ atribuição / chamada
    def assign(self, st: A.Assign, tail):
        val = self.ex(st.value, as_value=not st.is_set)
        t = st.target
        if isinstance(t, A.Name):
            sym = self.resolve(t.id)
            if sym and sym.kind == "proc" and self.ctx.proc and sym.name.lower() == self.ctx.proc.name.lower():
                self.em.emit(f"_ret = {val}{tail}")
                return
        self.em.emit(f"{self.target(t)} = {_strip_parens(val)}{tail}")

    def target(self, t: A.Expr) -> str:
        if isinstance(t, A.Name):
            sym = self.resolve(t.id)
            if sym is None:
                if self.is_form and self.m.options.get("explicit") is not None \
                        and t.id.lower() not in BUILTINS:
                    # com Option Explicit, nome não declarado num formulário é um controle
                    self.info(f"'{t.id}' tratado como controle do formulário (Me!{t.id})")
                    return f"self[{py_str(t.id)}]"
                self.ctx.locals[t.id.lower()] = Sym(t.id, "var", self.m.name)
                if self.m.options.get("explicit") is not None:
                    self.warn(f"variável não declarada: {t.id}")
                return py_name(t.id)
            if sym.kind == "var" and sym.module.lower() == self.m.name.lower() and t.id.lower() not in self.ctx.locals:
                if self.is_class:
                    return f"self.{py_name(sym.name)}"
                self.ctx.globals_used.add(py_name(sym.name))
                return py_name(sym.name)
            if sym.kind == "static":
                return sym.name
            if sym.kind == "proc" and self.ctx.proc and sym.name.lower() == self.ctx.proc.name.lower():
                return "_ret"
            return self.ex(t)
        if isinstance(t, A.Call) and isinstance(t.target, A.Member) and t.target.name.lower() == "fields":
            return self.ex(t, as_value=False) + ".Value"
        if isinstance(t, A.Bang):
            return self.ex(t, as_value=True)
        if isinstance(t, A.Member) and isinstance(t.obj, A.Me) and self.is_form \
                and t.name.lower() not in self.syms:
            return self.ex(t, as_value=True)
        return self.ex(t, as_value=False)

    def call_stmt(self, e: A.Expr) -> str:
        if isinstance(e, A.Call):
            tgt = e.target
            if isinstance(tgt, A.Name):
                sym = self.resolve(tgt.id)
                if sym and sym.kind in ("var", "param") and not e.args:
                    return self.ex(tgt) + "()"
            return _strip_parens(self.render_call(tgt, e.args, force_call=True))
        return self.ex(e) + "()"

    # ------------------------------------------------------------ expressões
    def cond(self, e: A.Expr) -> str:
        return _strip_parens(self.ex(e))

    def stringy(self, e) -> bool:
        if isinstance(e, A.Lit):
            return e.kind == "str"
        if isinstance(e, A.Bin) and e.op == "&":
            return True
        if isinstance(e, A.Name):
            s = self.resolve(e.id)
            return s is not None and (s.vtype or "").lower() == "string"
        if isinstance(e, A.Call) and isinstance(e.target, A.Name):
            low = e.target.id.lower().rstrip("$")
            if low in STRING_FUNCS:
                return True
            s = self.resolve(e.target.id)
            return s is not None and s.kind == "proc" and (s.vtype or "").lower() == "string"
        return False

    def lookup_global(self, name: str) -> Optional[Sym]:
        low = name.lower()
        if low in self.syms:
            return self.syms[low]
        return self.project.public.get(low)

    def resolve(self, name: str) -> Optional[Sym]:
        low = name.lower()
        if low in self.ctx.locals:
            return self.ctx.locals[low]
        return self.lookup_global(low)

    def sym_ref(self, sym: Sym) -> str:
        """Nome Python para um símbolo, importando o módulo se for de outro."""
        if sym.kind == "class":
            if sym.module.lower() != self.m.name.lower():
                self.imports.add(sym.module)
                return f"{py_name(sym.module)}.{py_name(sym.name)}"
            return py_name(sym.name)
        if sym.kind == "module":
            self.imports.add(sym.module)
            return py_name(sym.module)
        if sym.module.lower() != self.m.name.lower():
            self.imports.add(sym.module)
            return f"{py_name(sym.module)}.{py_name(sym.name)}"
        if self.is_class and sym.kind in ("var", "proc") and sym.kind != "enummember":
            return f"self.{py_name(sym.name)}"
        return py_name(sym.name)

    def ex(self, e: A.Expr, as_value: bool = True) -> str:
        if isinstance(e, A.Lit):
            return self.lit(e)
        if isinstance(e, A.Name):
            return self.name(e, as_value)
        if isinstance(e, A.Me):
            return "self"
        if isinstance(e, A.Member):
            return self.member(e, as_value)
        if isinstance(e, A.Bang):
            obj = self.ex(e.obj, as_value=False)
            if as_value:
                return f"{obj}[{py_str(e.name)}]"
            return f"{obj}.item({py_str(e.name)})"
        if isinstance(e, A.Call):
            s = self.render_call(e.target, e.args)
            if as_value and isinstance(e.target, A.Member) and e.target.name.lower() == "fields":
                s += ".Value"
            return s
        if isinstance(e, A.Unary):
            if e.op == "NOT":
                inner = self.ex(e.operand)
                if isinstance(e.operand, A.Lit) and isinstance(e.operand.value, int) and not isinstance(e.operand.value, bool):
                    return f"(~{inner})"
                return f"(not {inner})"
            return f"(-{self.ex(e.operand)})"
        if isinstance(e, A.Bin):
            return self.binop(e)
        if isinstance(e, A.New):
            return self.new_expr(e.cls) + "()"
        if isinstance(e, A.TypeOf):
            return f"vb.TypeOf({self.ex(e.expr)}, {py_str(e.type_name)})"
        raise NotImplementedError(type(e).__name__)

    def lit(self, e: A.Lit) -> str:
        if e.kind == "str":
            return py_str(e.value)
        if e.kind == "num":
            return repr(e.value)
        if e.kind == "bool":
            return "True" if e.value else "False"
        if e.kind == "date":
            d: datetime = e.value
            parts = [d.year, d.month, d.day] + ([d.hour, d.minute, d.second] if d.time() != d.min.time() else [])
            return f"datetime({', '.join(map(str, parts))})"
        return "None"

    def name(self, e: A.Name, as_value: bool) -> str:
        low = e.id.lower()
        sym = self.resolve(e.id)
        if sym is not None:
            if sym.kind in ("param", "var", "const") and low in self.ctx.locals:
                return py_name(sym.name)
            if sym.kind == "static":
                return sym.name
            if sym.kind == "proc":
                if self.ctx.proc and sym.name.lower() == self.ctx.proc.name.lower() \
                        and self.ctx.proc.kind in ("function", "property_get"):
                    return "_ret"
                ref = self.sym_ref(sym)
                if "property_get" in sym.proc_kind and self.is_class and sym.module.lower() == self.m.name.lower():
                    return ref
                return ref + "()"
            if sym.kind == "var" and sym.module.lower() == self.m.name.lower() and not self.is_class:
                return py_name(sym.name)
            return self.sym_ref(sym)
        if low in OBJECTS:
            return f"vb.{OBJECTS[low]}"
        if low in NOARG_BUILTINS:
            return f"vb.{BUILTINS[low]}()"
        if low in BUILTINS:
            return f"vb.{BUILTINS[low]}"
        const = next((k for k in CONSTANTS if k.lower() == low), None)
        if const:
            return f"vb.{const}"
        if low == "debug":
            return "vb.Debug"
        if self.is_form:
            self.info(f"'{e.id}' tratado como controle do formulário (Me!{e.id})")
            return f"self[{py_str(e.id)}]" if as_value else f"self.item({py_str(e.id)})"
        if re.match(r"(?i)^(ac|db|vb|ad)[A-Z]", e.id):
            self.warn(f"constante '{e.id}' não mapeada no runtime (acrescente em vb.CONSTANTS)")
            return f"vb.{e.id}"
        self.warn(f"nome não declarado: '{e.id}'")
        return py_name(e.id)

    def member(self, e: A.Member, as_value: bool) -> str:
        if e.obj is None:
            if not self.ctx.with_stack:
                self.error(f".{e.name} fora de bloco With")
                return f"_with_.{e.name}"
            base = self.ctx.with_stack[-1]
            if e.name == "__with__":
                return base
            return self._member_access(base, e.name, as_value, obj_node=None)
        if isinstance(e.obj, A.Name):
            sym = self.resolve(e.obj.id)
            if sym and sym.kind == "enum":
                return f"{self.sym_ref(sym)}.{py_name(e.name)}"
            if sym and sym.kind == "module":
                inner = self.project.module_syms.get(sym.module.lower(), {}).get(e.name.lower())
                if inner:
                    self.imports.add(sym.module)
                    ref = f"{py_name(sym.module)}.{py_name(inner.name)}"
                    return ref + "()" if inner.kind == "proc" else ref
        if isinstance(e.obj, A.Me) and self.is_form:
            key = e.name.lower()
            if key not in self.syms:
                return f"self[{py_str(e.name)}]" if as_value else f"self.item({py_str(e.name)})"
        base = self.ex(e.obj, as_value=False)
        return self._member_access(base, e.name, as_value, e.obj)

    def _member_access(self, base, name, as_value, obj_node):
        if name.lower() in NOARG_METHODS:
            return f"{base}.{name}()"
        return f"{base}.{py_name(name) if not name.isidentifier() else name}"

    def render_call(self, target: A.Expr, args: list[A.Arg], force_call=False) -> str:
        argv = self.render_args(args, target)
        if isinstance(target, A.Name):
            low = target.id.lower()
            sym = self.resolve(target.id)
            if sym is not None:
                if sym.kind in ("var", "param", "const", "static") and args:
                    base = self.name(target, as_value=False)
                    return f"{base}[{', '.join(a for a in argv)}]"
                if sym.kind == "proc":
                    if self.ctx.proc and sym.name.lower() == self.ctx.proc.name.lower():
                        return f"{self.sym_ref(sym)}({', '.join(argv)})"
                    return f"{self.sym_ref(sym)}({', '.join(argv)})"
                if sym.kind == "class":
                    return f"{self.sym_ref(sym)}({', '.join(argv)})"
            if low in OBJECTS:
                return f"vb.{OBJECTS[low]}({', '.join(argv)})"
            if low.rstrip("$") in BUILTINS:
                return f"vb.{BUILTINS[low.rstrip('$')]}({', '.join(argv)})"
            if self.is_form and not args:
                return f"self.{target.id}()"
            self.warn(f"chamada a procedimento desconhecido: {target.id}")
            return f"{py_name(target.id)}({', '.join(argv)})"
        if isinstance(target, A.Member):
            if isinstance(target.obj, A.Name):
                sym = self.resolve(target.obj.id)
                if sym and sym.kind == "module":
                    inner = self.project.module_syms.get(sym.module.lower(), {}).get(target.name.lower())
                    if inner:
                        self.imports.add(sym.module)
                        return f"{py_name(sym.module)}.{py_name(inner.name)}({', '.join(argv)})"
            if target.obj is None and target.name != "__with__":
                if not self.ctx.with_stack:
                    self.error(f".{target.name} fora de With")
                    base = "_with_"
                else:
                    base = self.ctx.with_stack[-1]
            else:
                base = self.ex(target.obj, as_value=False)
            if isinstance(target.obj, A.Me) and self.is_form and target.name.lower() not in self.syms and args:
                return f"self.item({py_str(target.name)})({', '.join(argv)})"
            return f"{base}.{target.name}({', '.join(argv)})"
        return f"{self.ex(target, as_value=False)}({', '.join(argv)})"

    def render_args(self, args: list[A.Arg], target) -> list[str]:
        params = None
        if isinstance(target, A.Name):
            sym = self.resolve(target.id)
            if sym and sym.kind == "proc":
                params = sym.params
        out = []
        keyword_mode = False
        for i, a in enumerate(args):
            if a.name:
                out.append(f"{py_name(a.name)}={self.ex(a.value)}")
                keyword_mode = True
                continue
            if a.value is None:
                if params is not None and i < len(params):
                    keyword_mode = True  # pula: o parâmetro mantém o valor padrão
                    continue
                if any(x.value is not None for x in args[i + 1:]):
                    out.append("None")
                continue
            val = self.ex(a.value)
            if keyword_mode and params is not None and i < len(params):
                out.append(f"{py_name(params[i].name)}={val}")
            else:
                out.append(_strip_parens(val))
        return out

    def binop(self, e: A.Bin) -> str:
        op = e.op
        if op == "&":
            parts = []

            def flat(x):
                if isinstance(x, A.Bin) and x.op == "&":
                    flat(x.left)
                    flat(x.right)
                else:
                    parts.append(_strip_parens(self.ex(x)))
            flat(e)
            return f"vb.Concat({', '.join(parts)})"
        if op in CMP_OPS:
            if any(isinstance(x, A.Lit) and x.kind == "null" for x in (e.left, e.right)):
                self.warn("comparação '= Null' é sempre Null no VBA (nunca verdadeira); use IsNull()")
                return "None"
            stringy = self.stringy(e.left) or self.stringy(e.right)
            return "(" + self.compare(op, self.ex(e.left), self.ex(e.right), stringy) + ")"
        a, b = self.ex(e.left), self.ex(e.right)
        if op in ("+", "-", "*", "/"):
            return f"({a} {op} {b})"
        if op == "^":
            return f"({a} ** {b})"
        if op == "\\":
            return f"vb.IntDiv({_strip_parens(a)}, {_strip_parens(b)})"
        if op == "MOD":
            return f"vb.Mod({_strip_parens(a)}, {_strip_parens(b)})"
        if op in ("AND", "OR"):
            bitwise = any(isinstance(x, A.Lit) and x.kind == "num" for x in (e.left, e.right))
            sym = ("&" if op == "AND" else "|") if bitwise else op.lower()
            return f"({a} {sym} {b})"
        if op == "XOR":
            return f"vb.Xor({_strip_parens(a)}, {_strip_parens(b)})"
        if op == "EQV":
            return f"(not vb.Xor({_strip_parens(a)}, {_strip_parens(b)}))"
        if op == "IMP":
            return f"((not {a}) or {b})"
        if op == "IS":
            return f"({a} is {b})"
        if op == "LIKE":
            return f"vb.Like({_strip_parens(a)}, {_strip_parens(b)})"
        raise NotImplementedError(op)


@dataclass
class _TryBlock(A.Stmt):
    body: list
    handler: list
    exit_sec: list
    exit_label: Optional[str]
    falls_through: bool


def _strip_parens(s: str) -> str:
    """Remove um par de parênteses externos redundante (ignorando texto entre aspas)."""
    if s.startswith("(") and s.endswith(")"):
        depth = 0
        quote = None
        for i, ch in enumerate(s):
            if quote:
                if ch == quote and s[i - 1] != "\\":
                    quote = None
                continue
            if ch in "\"'":
                quote = ch
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0 and i != len(s) - 1:
                    return s
        return s[1:-1]
    return s


def transpile_modules(modules: list[A.Module], diags: Optional[Diagnostics] = None) -> dict[str, str]:
    diags = diags or Diagnostics()
    project = Project(modules)
    out = {}
    for m in modules:
        out[m.name] = ModuleGenerator(m, project, diags).generate()
    return out
