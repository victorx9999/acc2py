"""Lexer e parser do VBA (módulos padrão, de classe e de formulário/relatório)."""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Optional

from . import ast as A
from ..sql.parser import parse_access_date, SQLSyntaxError


class VBASyntaxError(Exception):
    pass


@dataclass
class Tok:
    kind: str  # IDENT, STR, NUM, DATE, OP, NL, COLON, COMMENT, EOF
    value: object
    line: int
    bracket: bool = False
    space_before: bool = False

    def kw(self, *words) -> bool:
        return self.kind == "IDENT" and not self.bracket and str(self.value).lower() in words


_IDCHARS = re.compile(r"[^\W\d]\w*", re.UNICODE)
_NUM = re.compile(r"(?:&[Hh][0-9A-Fa-f]+|&[Oo][0-7]+|(?:\d+\.\d*|\.\d+|\d+)(?:[eEdD][+-]?\d+)?)[%&!#@^]?")
_DATE = re.compile(r"#\s*[0-9][0-9/\-.: ]*(?:[AaPp][Mm])?\s*#|#\s*[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}\s*#")


def _strip_header(src: str) -> tuple[str, list[str]]:
    """Remove cabeçalho de exportação (VERSION/BEGIN…END/Attribute VB_*)."""
    lines = src.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out, meta = [], []
    in_begin = 0
    for ln in lines:
        s = ln.strip()
        if s.upper().startswith("VERSION ") and not out:
            meta.append(s)
            out.append("")
            continue
        if (s.upper() == "BEGIN" or s.upper().startswith("BEGIN ")) and not any(o.strip() for o in out):
            in_begin += 1
            out.append("")
            continue
        if in_begin:
            if s.upper() == "END":
                in_begin -= 1
            out.append("")
            continue
        if s.startswith("Attribute VB_"):
            meta.append(s)
            out.append("")
            continue
        out.append(ln)
    return "\n".join(out), meta


def tokenize(src: str) -> list[Tok]:
    toks: list[Tok] = []
    i, line, n = 0, 1, len(src)
    space = False
    while i < n:
        c = src[i]
        if c in " \t":
            # continuação de linha: " _" seguido de quebra
            m = re.match(r"[ \t]+_[ \t]*\n", src[i:])
            if m:
                i += m.end()
                line += 1
            else:
                i += 1
            space = True
            continue
        if c == "\n":
            toks.append(Tok("NL", None, line))
            line += 1
            i += 1
            space = False
            continue
        start_space, space = space, False
        if c == "'":
            j = src.find("\n", i)
            j = n if j == -1 else j
            text = src[i + 1: j]
            # comentário com continuação " _"
            while text.rstrip().endswith(" _") and j < n:
                k = src.find("\n", j + 1)
                k = n if k == -1 else k
                text = text.rstrip()[:-1] + src[j + 1: k]
                line += 1
                j = k
            toks.append(Tok("COMMENT", text, line))
            i = j
            continue
        if c == '"':
            j = i + 1
            buf = []
            while j < n:
                if src[j] == '"':
                    if j + 1 < n and src[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                if src[j] == "\n":
                    raise VBASyntaxError(f"linha {line}: texto sem aspas de fechamento")
                buf.append(src[j])
                j += 1
            toks.append(Tok("STR", "".join(buf), line, space_before=start_space))
            i = j + 1
            continue
        if c == "#":
            m = _DATE.match(src, i)
            if m:
                try:
                    d = parse_access_date(m.group()[1:-1].strip())
                    toks.append(Tok("DATE", d, line, space_before=start_space))
                    i = m.end()
                    continue
                except SQLSyntaxError:
                    pass
        if c == "[":
            j = src.find("]", i)
            if j != -1:
                toks.append(Tok("IDENT", src[i + 1: j], line, bracket=True, space_before=start_space))
                i = j + 1
                continue
        if c.isdigit() or (c == "." and i + 1 < n and src[i + 1].isdigit()
                           and (not toks or toks[-1].kind not in ("IDENT",) or start_space)) \
                or (c == "&" and i + 1 < n and src[i + 1] in "HhOo" and re.match(r"&[Hh][0-9A-Fa-f]|&[Oo][0-7]", src[i:i + 3])):
            m = _NUM.match(src, i)
            if m:
                txt = m.group().rstrip("%&!#@^")
                if txt[:2].lower() == "&h":
                    v = int(txt[2:], 16)
                elif txt[:2].lower() == "&o":
                    v = int(txt[2:], 8)
                elif re.search(r"[.eEdD]", txt):
                    v = float(txt.replace("d", "e").replace("D", "e"))
                else:
                    v = int(txt)
                toks.append(Tok("NUM", v, line, space_before=start_space))
                i = m.end()
                continue
        m = _IDCHARS.match(src, i)
        if m:
            name = m.group()
            j = m.end()
            # sufixo de tipo: Left$, contador%, total&, x!, d#, c@
            if j < n and src[j] in "$%&!#@" and (j + 1 >= n or not (src[j + 1].isalnum() or src[j + 1] in "_[")):
                if not (src[j] == "!" and j + 1 < n and src[j + 1] == "["):
                    j += 1
            toks.append(Tok("IDENT", name, line, space_before=start_space))
            i = j
            continue
        two = src[i: i + 2]
        if two in (":=", "<>", "<=", ">=", "=<", "=>"):
            toks.append(Tok("OP", {"=<": "<=", "=>": ">="}.get(two, two), line, space_before=start_space))
            i += 2
            continue
        if c == ":":
            toks.append(Tok("COLON", ":", line))
            i += 1
            continue
        if c in "=<>+-*/\\^&(),.!;#":
            toks.append(Tok("OP", c, line, space_before=start_space))
            i += 1
            continue
        raise VBASyntaxError(f"linha {line}: caractere inesperado {c!r}")
    toks.append(Tok("NL", None, line))
    toks.append(Tok("EOF", None, line))
    return toks


_BLOCK_END = {"end", "else", "elseif", "next", "loop", "wend", "case"}


class Parser:
    def __init__(self, src: str, name: str = "Modulo", kind: str = "standard"):
        clean, self.meta = _strip_header(src)
        self.toks = tokenize(clean)
        self.i = 0
        self.module = A.Module(name=name, kind=kind)
        self.pending_comments: list[str] = []

    # ------------------------------------------------------------ util
    @property
    def t(self) -> Tok:
        return self.toks[self.i]

    def peek(self, k=1) -> Tok:
        return self.toks[min(self.i + k, len(self.toks) - 1)]

    def next(self) -> Tok:
        tok = self.toks[self.i]
        self.i += 1
        return tok

    def accept_kw(self, *w) -> bool:
        if self.t.kw(*w):
            self.i += 1
            return True
        return False

    def expect_kw(self, w):
        if not self.accept_kw(w):
            self.error(f"esperado '{w}'")

    def accept_op(self, op) -> bool:
        if self.t.kind == "OP" and self.t.value == op:
            self.i += 1
            return True
        return False

    def expect_op(self, op):
        if not self.accept_op(op):
            self.error(f"esperado '{op}'")

    def at_eos(self) -> bool:
        return self.t.kind in ("NL", "COLON", "COMMENT", "EOF")

    def error(self, msg):
        raise VBASyntaxError(f"{self.module.name}, linha {self.t.line}: {msg} (encontrado {self.t.value!r})")

    def ident(self) -> str:
        if self.t.kind == "IDENT":
            return str(self.next().value)
        self.error("esperado identificador")

    def skip_eos(self):
        """Pula separadores; guarda comentários soltos."""
        while self.t.kind in ("NL", "COLON", "COMMENT"):
            tok = self.next()
            if tok.kind == "COMMENT":
                self.pending_comments.append(tok.value)

    def raw_until_eol(self) -> str:
        parts = []
        while self.t.kind not in ("NL", "EOF", "COMMENT"):
            v = self.next().value
            parts.append(f'"{v}"' if isinstance(v, str) and self.toks[self.i - 1].kind == "STR" else str(v))
        return " ".join(parts)

    # ------------------------------------------------------------ módulo
    def parse(self) -> A.Module:
        m = self.module
        for meta in self.meta:
            if "VB_PredeclaredId" in meta or "VERSION" in meta:
                if m.kind == "standard" and not m.name.lower().startswith(("form_", "report_")):
                    m.kind = "class"
        while True:
            self.skip_eos()
            if self.t.kind == "EOF":
                break
            line = self.t.line
            comments, self.pending_comments = self.pending_comments, []
            if self.t.kw("option"):
                self.next()
                key = self.ident().lower()
                val = self.ident() if self.t.kind == "IDENT" else (str(self.next().value) if not self.at_eos() else "")
                m.options[key] = val
                continue
            public = True
            static = False
            mod_scope = False
            while self.t.kw("public", "private", "global", "friend", "static", "dim"):
                w = self.next().value.lower()
                if w == "private":
                    public = False
                if w == "dim":
                    public = False
                if w == "static":
                    static = True
                mod_scope = True
            if self.t.kw("sub", "function", "property"):
                proc = self.procedure(public, static)
                proc.line = line
                proc.comments_before = comments
                m.procedures.append(proc)
                continue
            m.header_comments.extend(comments)
            if self.t.kw("const"):
                self.next()
                for name, typ, val in self.const_items():
                    m.constants.append((name, typ, val, public))
                continue
            if self.t.kw("type"):
                self.next()
                m.types.append(self.udt(public))
                continue
            if self.t.kw("enum"):
                self.next()
                m.enums.append(self.enum(public))
                continue
            if self.t.kw("declare"):
                text = "Declare " + self.raw_until_eol()
                m.declares.append(text)
                continue
            if self.t.kw("event", "implements", "defint", "deflng", "defstr", "defbool", "defdbl", "defvar"):
                text = self.raw_until_eol()
                m.unsupported.append(A.Unsupported(text, "declaração de módulo não suportada", line=line))
                continue
            if mod_scope or self.t.kw("withevents"):
                self.accept_kw("withevents")
                for d in self.var_list():
                    m.variables.append(A.ModuleVar(d, public))
                continue
            text = self.raw_until_eol()
            m.unsupported.append(A.Unsupported(text, "código fora de procedimento", line=line))
        return m

    def const_items(self):
        items = []
        while True:
            name = self.ident()
            typ = None
            if self.accept_kw("as"):
                typ = self.type_name()
            self.expect_op("=")
            items.append((name, typ, self.expr()))
            if not self.accept_op(","):
                return items

    def type_name(self) -> str:
        self.accept_kw("new")
        name = self.ident()
        while self.accept_op("."):
            name += "." + self.ident()
        if self.accept_op("*"):  # String * 10
            self.next()
        return name

    def udt(self, public) -> A.UDT:
        name = self.ident()
        fields = []
        while True:
            self.skip_eos()
            if self.t.kw("end"):
                self.next()
                self.expect_kw("type")
                return A.UDT(name, fields, public)
            fields.extend(self.var_list())

    def enum(self, public) -> A.EnumDecl:
        name = self.ident()
        members = []
        while True:
            self.skip_eos()
            if self.t.kw("end"):
                self.next()
                self.expect_kw("enum")
                return A.EnumDecl(name, members, public)
            mname = self.ident()
            val = self.expr() if self.accept_op("=") else None
            members.append((mname, val))

    def var_list(self) -> list[A.VarDecl]:
        out = []
        while True:
            self.accept_kw("withevents")
            name = self.ident()
            d = A.VarDecl(name)
            if self.accept_op("("):
                d.is_array = True
                d.dims = self.dims_until_paren()
            if self.accept_kw("as"):
                d.new = self.t.kw("new")
                d.type = self.type_name()
            out.append(d)
            if not self.accept_op(","):
                return out

    def dims_until_paren(self):
        dims = []
        if self.accept_op(")"):
            return dims
        while True:
            a = self.expr()
            if self.accept_kw("to"):
                dims.append((a, self.expr()))
            else:
                dims.append((None, a))
            if self.accept_op(")"):
                return dims
            self.expect_op(",")

    def procedure(self, public, static) -> A.Procedure:
        w = self.next().value.lower()
        kind = w
        if w == "property":
            kind = "property_" + self.ident().lower()
        name = self.ident()
        params = []
        if self.accept_op("("):
            if not self.accept_op(")"):
                while True:
                    params.append(self.param())
                    if self.accept_op(")"):
                        break
                    self.expect_op(",")
        rtype = None
        if self.accept_kw("as"):
            rtype = self.type_name()
            if self.accept_op("("):
                self.expect_op(")")
        end_kw = "property" if kind.startswith("property") else kind
        body = self.block({"end " + end_kw})
        self.next()  # End
        self.next()  # Sub/Function/Property
        return A.Procedure(kind, name, params, body, rtype, public, static)

    def param(self) -> A.Param:
        p = A.Param("")
        while self.t.kw("optional", "byval", "byref", "paramarray"):
            w = self.next().value.lower()
            if w == "optional":
                p.optional = True
            elif w == "byval":
                p.byval = True
            elif w == "paramarray":
                p.paramarray = True
        p.name = self.ident()
        if self.accept_op("("):
            self.expect_op(")")
            p.is_array = True
        if self.accept_kw("as"):
            p.type = self.type_name()
        if self.accept_op("="):
            p.default = self.expr()
        return p

    # ------------------------------------------------------------ blocos
    def _matches(self, terms: set) -> Optional[str]:
        t = self.t
        if t.kind != "IDENT" or t.bracket:
            return None
        w = t.value.lower()
        if w == "end":
            nxt = self.peek()
            two = "end " + (str(nxt.value).lower() if nxt.kind == "IDENT" else "")
            if two in terms:
                return two
            return None
        if w in terms:
            return w
        if w == "loop" and "loop" in terms:
            return w
        return None

    def block(self, terms: set) -> list[A.Stmt]:
        body: list[A.Stmt] = []
        while True:
            self.flush_comments(body)
            if self.t.kind == "EOF":
                self.error(f"fim do arquivo antes de {' / '.join(sorted(terms))}")
            if self.t.kind in ("NL", "COLON"):
                self.next()
                continue
            if self.t.kind == "COMMENT":
                body.append(A.Comment(self.next().value, line=self.t.line))
                continue
            if self._matches(terms):
                return body
            st = self.statement(body)
            if st is not None:
                if self.t.kind == "COMMENT":
                    st.comment = self.next().value
                body.append(st)

    def flush_comments(self, body):
        for c in self.pending_comments:
            body.append(A.Comment(c))
        self.pending_comments = []

    def inline_block(self, stop_words: set) -> list[A.Stmt]:
        """Comandos do If de uma linha só (separados por ':')."""
        body = []
        while True:
            if self.t.kind in ("NL", "EOF"):
                return body
            if self.t.kind == "COLON":
                self.next()
                continue
            if self.t.kind == "COMMENT":
                body.append(A.Comment(self.next().value))
                continue
            if self.t.kw(*stop_words):
                return body
            st = self.statement(body)
            if st is not None:
                body.append(st)

    # ------------------------------------------------------------ comandos
    def statement(self, body) -> Optional[A.Stmt]:
        t = self.t
        line = t.line
        # rótulo  "Sai:"  ou número de linha
        if t.kind == "IDENT" and not t.bracket and self.peek().kind == "COLON" \
                and t.value.lower() not in ("else",) and (self.i == 0 or self.toks[self.i - 1].kind == "NL"):
            self.next()
            self.next()
            return A.Label(t.value, line=line)
        if t.kind == "NUM" and (self.i == 0 or self.toks[self.i - 1].kind == "NL"):
            self.next()
            return A.Label(f"L{t.value}", line=line)
        st = self._statement()
        if st is not None:
            st.line = line
        return st

    def _statement(self) -> Optional[A.Stmt]:
        t = self.t
        w = str(t.value).lower() if (t.kind == "IDENT" and not t.bracket) else ""
        if w == "rem":
            j = self.i
            self.raw_until_eol()
            return A.Comment(" ".join(str(x.value) for x in self.toks[j + 1: self.i]))
        if w in ("dim", "static"):
            self.next()
            return A.Dim(self.var_list(), static=(w == "static"))
        if w in ("private", "public") and self.peek().kw("const"):
            self.next()
        if w == "const" or self.t.kw("const"):
            self.next()
            return A.ConstDecl(self.const_items())
        if w == "redim":
            self.next()
            preserve = self.accept_kw("preserve")
            target = self.postfix(no_call=True)
            self.expect_op("(")
            dims = self.dims_until_paren()
            typ = self.type_name() if self.accept_kw("as") else None
            return A.ReDim(preserve, target, dims, typ)
        if w in ("set", "let"):
            self.next()
            target = self.postfix()
            self.expect_op("=")
            return A.Assign(target, self.expr(), is_set=(w == "set"))
        if w == "call":
            self.next()
            e = self.postfix()
            if not isinstance(e, A.Call):
                e = A.Call(e, [], parens=False)
            return A.CallStmt(e)
        if w == "if":
            return self.if_stmt()
        if w == "select":
            return self.select_case()
        if w == "for":
            return self.for_stmt()
        if w == "do":
            return self.do_loop()
        if w == "while":
            self.next()
            cond = self.expr()
            body = self.block({"wend"})
            self.next()
            return A.DoLoop(body, pre=("while", cond))
        if w == "exit":
            self.next()
            return A.Exit(self.ident().lower())
        if w == "with":
            self.next()
            obj = self.expr()
            body = self.block({"end with"})
            self.next(); self.next()
            return A.With(obj, body)
        if w == "on" and self.peek().kw("error"):
            self.next(); self.next()
            if self.accept_kw("resume"):
                self.expect_kw("next")
                return A.OnError("resume_next")
            self.expect_kw("goto")
            if self.t.kind == "NUM" and self.t.value == 0:
                self.next()
                return A.OnError("goto0")
            if self.accept_op("-"):
                self.next()
                return A.OnError("goto0")
            return A.OnError("goto", self.ident() if self.t.kind == "IDENT" else f"L{self.next().value}")
        if w == "goto":
            self.next()
            return A.GoTo(self.ident() if self.t.kind == "IDENT" else f"L{self.next().value}")
        if w == "resume":
            self.next()
            if self.at_eos():
                return A.Resume(None)
            if self.accept_kw("next"):
                return A.Resume("next")
            return A.Resume(self.ident() if self.t.kind == "IDENT" else f"L{self.next().value}")
        if w == "erase":
            self.next()
            names = [self.postfix()]
            while self.accept_op(","):
                names.append(self.postfix())
            return A.Erase(names)
        if w == "debug" and self.peek().kind == "OP" and self.peek().value == "." \
                and str(self.peek(2).value).lower() == "print":
            self.next(); self.next(); self.next()
            items = []
            while not self.at_eos():
                if self.accept_op(";") or self.accept_op(","):
                    continue
                items.append(self.expr())
            return A.DebugPrint(items)
        if w == "end" and (self.peek().kind in ("NL", "COLON", "COMMENT", "EOF")):
            self.next()
            return A.End()
        if w == "stop":
            self.next()
            return A.Comment("Stop (ponto de parada do depurador) removido")
        if w == "raiseevent":
            self.next()
            name = self.ident()
            args = self.paren_args() if self.accept_op("(") else []
            return A.RaiseEvent(name, args)
        if w in ("open", "close", "print", "write", "input", "line", "get", "put", "seek", "lock",
                 "unlock", "reset", "lset", "rset", "mid", "sendkeys", "appactivate", "kill",
                 "mkdir", "rmdir", "chdir", "filecopy", "name", "beep", "date", "time", "randomize",
                 "load", "unload", "doevents") and self._looks_special(w):
            return self.special(w)
        # atribuição ou chamada
        target = self.postfix()
        if self.accept_op("="):
            return A.Assign(target, self.expr())
        if isinstance(target, A.Call) and self.at_eos():
            return A.CallStmt(target)
        args = []
        if not self.at_eos():
            if isinstance(target, A.Call) and target.parens and len(target.args) == 1 and self.t.kind == "OP" \
                    and self.t.value == ",":
                # "Foo (a), b"  => o parêntese era só do primeiro argumento
                args = [target.args[0]]
                target = target.target
                self.next()
            args += self.bare_args()
            return A.CallStmt(A.Call(target, args, parens=False))
        return A.CallStmt(A.Call(target, [], parens=False))

    def _looks_special(self, w) -> bool:
        nxt = self.peek()
        if nxt.kind == "OP" and nxt.value in ("=", ".", "!"):
            return False
        if w in ("mid", "date", "time") and not (nxt.kind == "OP" and nxt.value == "("):
            return False
        if w == "name" and nxt.kind in ("NL", "COLON"):
            return False
        return True

    def special(self, w) -> A.Stmt:
        if w in ("beep", "doevents"):
            self.next()
            return A.Comment(f"{w.capitalize()} removido (sem equivalente)")
        if w == "randomize":
            self.next()
            arg = None if self.at_eos() else self.expr()
            return A.CallStmt(A.Call(A.Name("Randomize"), [A.Arg(arg)] if arg else []))
        if w == "kill":
            self.next()
            return A.CallStmt(A.Call(A.Name("Kill"), [A.Arg(self.expr())]))
        if w in ("mkdir", "rmdir"):
            self.next()
            return A.CallStmt(A.Call(A.Name("MkDir" if w == "mkdir" else "RmDir"), [A.Arg(self.expr())]))
        if w == "filecopy":
            self.next()
            a = self.expr()
            self.expect_op(",")
            return A.CallStmt(A.Call(A.Name("FileCopy"), [A.Arg(a), A.Arg(self.expr())]))
        text = self.raw_until_eol()
        reason = "E/S de arquivo do VBA (Open/Print #/Input #…) não é convertida; use open() do Python" \
            if w in ("open", "close", "print", "write", "input", "line", "get", "put", "seek", "lock", "unlock", "reset") \
            else f"comando '{w}' não suportado"
        return A.Unsupported(text, reason)

    def if_stmt(self) -> A.Stmt:
        self.expect_kw("if")
        cond = self.expr()
        self.expect_kw("then")
        if self.t.kind not in ("NL", "COMMENT", "EOF"):
            # If de uma linha
            body = self.inline_block({"else"})
            else_body = None
            if self.accept_kw("else"):
                else_body = self.inline_block(set())
            self.accept_kw("end") and self.accept_kw("if")
            return A.If([(cond, body)], else_body)
        branches = []
        body = self.block({"elseif", "else", "end if"})
        branches.append((cond, body))
        else_body = None
        while True:
            if self.accept_kw("elseif"):
                c = self.expr()
                self.expect_kw("then")
                branches.append((c, self.block({"elseif", "else", "end if"})))
            elif self.accept_kw("else"):
                else_body = self.block({"end if"})
            else:
                self.expect_kw("end")
                self.expect_kw("if")
                return A.If(branches, else_body)

    def select_case(self) -> A.Stmt:
        self.expect_kw("select")
        self.expect_kw("case")
        subject = self.expr()
        cases, else_body = [], None
        while True:
            self.skip_eos()
            self.pending_comments = []
            if self.accept_kw("end"):
                self.expect_kw("select")
                return A.SelectCase(subject, cases, else_body)
            self.expect_kw("case")
            if self.accept_kw("else"):
                else_body = self.block({"case", "end select"})
                continue
            clauses = []
            while True:
                if self.accept_kw("is"):
                    op = self.next().value
                    clauses.append(A.CaseClause("is", self.expr(), op=op))
                else:
                    a = self.expr()
                    if self.accept_kw("to"):
                        clauses.append(A.CaseClause("range", a, self.expr()))
                    else:
                        clauses.append(A.CaseClause("value", a))
                if not self.accept_op(","):
                    break
            cases.append((clauses, self.block({"case", "end select"})))

    def for_stmt(self) -> A.Stmt:
        self.expect_kw("for")
        if self.accept_kw("each"):
            var = self.postfix()
            self.expect_kw("in")
            it = self.expr()
            body = self.block({"next"})
            self._end_next()
            return A.ForEach(var, it, body)
        var = self.postfix(no_call=True)
        self.expect_op("=")
        start = self.expr()
        self.expect_kw("to")
        end = self.expr()
        step = self.expr() if self.accept_kw("step") else None
        body = self.block({"next"})
        self._end_next()
        return A.For(var, start, end, step, body)

    def _end_next(self):
        self.expect_kw("next")
        if self.t.kind == "IDENT":
            self.next()
            # "Next j, i" fecha dois laços: deixa um "Next" virtual para o laço externo
            if self.t.kind == "OP" and self.t.value == ",":
                self.toks[self.i] = Tok("IDENT", "Next", self.t.line)

    def do_loop(self) -> A.Stmt:
        self.expect_kw("do")
        pre = None
        if self.t.kw("while", "until"):
            kind = self.next().value.lower()
            pre = (kind, self.expr())
        body = self.block({"loop"})
        self.expect_kw("loop")
        post = None
        if self.t.kw("while", "until"):
            kind = self.next().value.lower()
            post = (kind, self.expr())
        return A.DoLoop(body, pre, post)

    # ------------------------------------------------------------ argumentos
    def bare_args(self) -> list[A.Arg]:
        args = []
        while True:
            if self.t.kind == "OP" and self.t.value == ",":
                args.append(A.Arg(None))
                self.next()
                continue
            args.append(self.arg())
            if not self.accept_op(","):
                return args
            if self.at_eos():
                args.append(A.Arg(None))
                return args

    def arg(self) -> A.Arg:
        if self.t.kind == "IDENT" and self.peek().kind == "OP" and self.peek().value == ":=":
            name = self.next().value
            self.next()
            return A.Arg(self.expr(), name)
        self.accept_kw("byval")
        return A.Arg(self.expr())

    def paren_args(self) -> list[A.Arg]:
        args = []
        if self.accept_op(")"):
            return args
        while True:
            if self.t.kind == "OP" and self.t.value == ",":
                args.append(A.Arg(None))
            elif self.t.kind == "OP" and self.t.value == ")":
                args.append(A.Arg(None))
            else:
                args.append(self.arg())
            if self.accept_op(")"):
                return args
            self.expect_op(",")

    # ------------------------------------------------------------ expressões
    def expr(self) -> A.Expr:
        return self._binary(0)

    _LEVELS = [
        ("imp",), ("eqv",), ("xor",), ("or",), ("and",),
    ]

    def _binary(self, level):
        if level < len(self._LEVELS):
            op = self._LEVELS[level][0]
            left = self._binary(level + 1)
            while self.accept_kw(op):
                left = A.Bin(op.upper(), left, self._binary(level + 1))
            return left
        return self.not_expr()

    def not_expr(self):
        if self.accept_kw("not"):
            return A.Unary("NOT", self.not_expr())
        return self.cmp_expr()

    def cmp_expr(self):
        left = self.concat_expr()
        while True:
            t = self.t
            if t.kind == "OP" and t.value in ("=", "<>", "<", ">", "<=", ">="):
                self.next()
                left = A.Bin(t.value, left, self.concat_expr())
            elif t.kw("is"):
                self.next()
                left = A.Bin("IS", left, self.concat_expr())
            elif t.kw("like"):
                self.next()
                left = A.Bin("LIKE", left, self.concat_expr())
            else:
                return left

    def concat_expr(self):
        left = self.add_expr()
        while self.t.kind == "OP" and self.t.value == "&":
            self.next()
            left = A.Bin("&", left, self.add_expr())
        return left

    def add_expr(self):
        left = self.mod_expr()
        while self.t.kind == "OP" and self.t.value in ("+", "-"):
            op = self.next().value
            left = A.Bin(op, left, self.mod_expr())
        return left

    def mod_expr(self):
        left = self.idiv_expr()
        while self.accept_kw("mod"):
            left = A.Bin("MOD", left, self.idiv_expr())
        return left

    def idiv_expr(self):
        left = self.mul_expr()
        while self.accept_op("\\"):
            left = A.Bin("\\", left, self.mul_expr())
        return left

    def mul_expr(self):
        left = self.unary_expr()
        while self.t.kind == "OP" and self.t.value in ("*", "/"):
            op = self.next().value
            left = A.Bin(op, left, self.unary_expr())
        return left

    def unary_expr(self):
        if self.accept_op("-"):
            return A.Unary("-", self.unary_expr())
        if self.accept_op("+"):
            return self.unary_expr()
        return self.pow_expr()

    def pow_expr(self):
        left = self.postfix()
        while self.accept_op("^"):
            right = A.Unary("-", self.postfix()) if self.accept_op("-") else self.postfix()
            left = A.Bin("^", left, right)
        return left

    def postfix(self, no_call=False) -> A.Expr:
        e = self.primary()
        while True:
            t = self.t
            if t.kind == "OP" and t.value == "." and not t.space_before:
                self.next()
                e = A.Member(e, self.ident())
            elif t.kind == "OP" and t.value == "!" and not t.space_before:
                self.next()
                e = A.Bang(e, self.ident())
            elif t.kind == "OP" and t.value == "(" and not no_call:
                self.next()
                e = A.Call(e, self.paren_args())
            else:
                return e

    def primary(self) -> A.Expr:
        t = self.t
        if t.kind == "NUM":
            self.next()
            return A.Lit(t.value, "num")
        if t.kind == "STR":
            self.next()
            return A.Lit(t.value, "str")
        if t.kind == "DATE":
            self.next()
            return A.Lit(t.value, "date")
        if t.kind == "OP" and t.value == "(":
            self.next()
            e = self.expr()
            self.expect_op(")")
            return e
        if t.kind == "OP" and t.value == ".":  # .Campo dentro de With
            self.next()
            return A.Member(None, self.ident())
        if t.kind == "OP" and t.value == "!":  # !Campo dentro de With
            self.next()
            return A.Bang(A.Member(None, "__with__"), self.ident())
        if t.kind == "IDENT":
            w = str(t.value).lower() if not t.bracket else ""
            if w in ("true", "false"):
                self.next()
                return A.Lit(w == "true", "bool")
            if w == "null":
                self.next()
                return A.Lit(None, "null")
            if w == "nothing":
                self.next()
                return A.Lit(None, "nothing")
            if w == "empty":
                self.next()
                return A.Lit(None, "empty")
            if w == "me":
                self.next()
                return A.Me()
            if w == "new":
                self.next()
                return A.New(self.type_name())
            if w == "typeof":
                self.next()
                e = self.postfix()
                self.expect_kw("is")
                return A.TypeOf(e, self.type_name())
            if w == "addressof":
                self.next()
                return A.Name(self.ident())
            self.next()
            return A.Name(str(t.value))
        self.error("expressão inválida")


def parse_module(src: str, name: str = "Modulo", kind: Optional[str] = None) -> A.Module:
    if kind is None:
        low = name.lower()
        kind = "form" if low.startswith("form_") else "report" if low.startswith("report_") else "standard"
    return Parser(src, name, kind).parse()
