"""Lexer e parser recursivo para o SQL do Access (Jet/ACE)."""
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime

from . import ast as A


class SQLSyntaxError(Exception):
    pass


# ---------------------------------------------------------------- lexer
@dataclass
class Tok:
    kind: str   # IDENT, QIDENT ([...]), STR, NUM, DATE, OP, EOF
    value: str
    pos: int

    def is_kw(self, *words: str) -> bool:
        return self.kind == "IDENT" and self.value.upper() in words


_TOKEN_RE = re.compile(r"""
    (?P<ws>\s+)
  | (?P<qident>\[[^\]]*\])
  | (?P<str>'(?:[^']|'')*'|"(?:[^"]|"")*")
  | (?P<date>\#[^#\r\n]+\#)
  | (?P<num>(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?)
  | (?P<ident>[^\W\d][\w]*)
  | (?P<op><=|>=|<>|[=<>+\-*/\\&(),.!;^?])
""", re.VERBOSE | re.UNICODE)


def tokenize(sql: str) -> list[Tok]:
    toks: list[Tok] = []
    pos = 0
    while pos < len(sql):
        m = _TOKEN_RE.match(sql, pos)
        if not m:
            raise SQLSyntaxError(f"Caractere inesperado {sql[pos]!r} na posição {pos}")
        kind = m.lastgroup
        text = m.group()
        if kind == "ws":
            pass
        elif kind == "qident":
            toks.append(Tok("QIDENT", text[1:-1], pos))
        elif kind == "str":
            q = text[0]
            toks.append(Tok("STR", text[1:-1].replace(q + q, q), pos))
        elif kind == "date":
            toks.append(Tok("DATE", text[1:-1].strip(), pos))
        elif kind == "num":
            toks.append(Tok("NUM", text, pos))
        elif kind == "ident":
            toks.append(Tok("IDENT", text, pos))
        else:
            toks.append(Tok("OP", text, pos))
        pos = m.end()
    toks.append(Tok("EOF", "", len(sql)))
    return toks


def parse_access_date(text: str) -> datetime:
    """Datas literais no SQL do Access usam formato americano (m/d/a) ou ISO."""
    t = text.strip()
    fmts = ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%m/%d/%Y", "%m/%d/%Y %H:%M:%S",
            "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M", "%m-%d-%Y", "%m/%d/%y",
            "%H:%M:%S", "%H:%M", "%I:%M:%S %p", "%I:%M %p", "%d-%b-%Y", "%d %b %Y", "%b %d, %Y"]
    for f in fmts:
        try:
            d = datetime.strptime(t, f)
            if f.startswith(("%H", "%I")):  # só hora: Access usa 30/12/1899
                d = d.replace(year=1899, month=12, day=30)
            return d
        except ValueError:
            continue
    raise SQLSyntaxError(f"Data literal não reconhecida: #{text}#")


# ---------------------------------------------------------------- parser
_STOP_ALIAS = {"WHERE", "GROUP", "HAVING", "ORDER", "INNER", "LEFT", "RIGHT", "JOIN", "ON",
               "UNION", "PIVOT", "SET", "FROM", "IN", "OUTER", "FULL", "CROSS", "SELECT",
               "INTO", "VALUES", "WITH"}


class Parser:
    def __init__(self, sql: str):
        self.sql = sql
        self.toks = tokenize(sql)
        self.i = 0

    # utilidades -----------------------------------------------------
    @property
    def t(self) -> Tok:
        return self.toks[self.i]

    def peek(self, k: int = 1) -> Tok:
        return self.toks[min(self.i + k, len(self.toks) - 1)]

    def next(self) -> Tok:
        tok = self.t
        self.i += 1
        return tok

    def accept_kw(self, *words: str) -> bool:
        if self.t.is_kw(*words):
            self.i += 1
            return True
        return False

    def expect_kw(self, word: str):
        if not self.accept_kw(word):
            self.error(f"esperado {word}")

    def accept_op(self, op: str) -> bool:
        if self.t.kind == "OP" and self.t.value == op:
            self.i += 1
            return True
        return False

    def expect_op(self, op: str):
        if not self.accept_op(op):
            self.error(f"esperado '{op}'")

    def error(self, msg: str):
        ctx = self.sql[max(0, self.t.pos - 25): self.t.pos + 25].replace("\n", " ")
        raise SQLSyntaxError(f"{msg} perto de: …{ctx}… (encontrado {self.t.value!r})")

    def ident(self) -> str:
        tok = self.t
        if tok.kind in ("IDENT", "QIDENT"):
            self.i += 1
            return tok.value
        self.error("esperado identificador")

    # entrada --------------------------------------------------------
    def parse(self) -> A.Statement:
        params = self.parse_parameters()
        stmt = self.statement()
        self.accept_op(";")
        if self.t.kind != "EOF":
            self.error("texto inesperado após o fim do comando")
        stmt.parameters = params
        return stmt

    def parse_parameters(self) -> list[tuple[str, str]]:
        params = []
        if self.accept_kw("PARAMETERS"):
            while True:
                name = self.ident()
                while self.t.kind == "OP" and self.t.value in (".", "!"):
                    self.next()
                    name += "!" + self.ident()
                typ = ""
                if self.t.kind == "IDENT":
                    typ = self.next().value
                    if self.accept_op("("):
                        typ += "(" + self.next().value + ")"
                        self.expect_op(")")
                params.append((name, typ))
                if not self.accept_op(","):
                    break
            self.expect_op(";")
        return params

    def statement(self) -> A.Statement:
        if self.t.is_kw("TRANSFORM"):
            return self.transform()
        if self.t.is_kw("UPDATE"):
            return self.update()
        if self.t.is_kw("DELETE"):
            return self.delete()
        if self.t.is_kw("INSERT"):
            return self.insert()
        return self.select_or_union()

    def select_or_union(self) -> A.Statement:
        parts = [self.select_core()]
        flags = []
        while self.accept_kw("UNION"):
            flags.append(self.accept_kw("ALL"))
            if not self.accept_kw("DISTINCT"):
                pass
            if self.t.is_kw("TABLE"):  # UNION TABLE x
                self.next()
                name = self.ident()
                parts.append(A.Select(columns=[(A.Star(), None)], from_=A.TableRef(name)))
                continue
            paren = self.accept_op("(")
            parts.append(self.select_core())
            if paren:
                self.expect_op(")")
        if len(parts) == 1:
            sel = parts[0]
            return sel
        # o ORDER BY da última parte vale para a união inteira
        order = parts[-1].order_by
        parts[-1].order_by = []
        return A.Union(parts=parts, all_flags=flags, order_by=order)

    def select_core(self) -> A.Select:
        if self.accept_op("("):
            inner = self.select_core()
            self.expect_op(")")
            return inner
        if self.t.is_kw("TABLE"):  # TABLE x  ==  SELECT * FROM x
            self.next()
            return A.Select(columns=[(A.Star(), None)], from_=A.TableRef(self.ident()))
        self.expect_kw("SELECT")
        sel = A.Select()
        if self.accept_kw("DISTINCT"):
            sel.distinct = "DISTINCT"
        elif self.accept_kw("DISTINCTROW"):
            sel.distinct = "DISTINCTROW"
        else:
            self.accept_kw("ALL")
        if self.accept_kw("TOP"):
            sel.top = float(self.next().value)
            sel.top_percent = self.accept_kw("PERCENT")
        sel.columns = self.select_list()
        if self.accept_kw("INTO"):
            sel.into = self.ident()
        if self.accept_kw("FROM"):
            sel.from_ = self.from_clause()
        if self.accept_kw("WHERE"):
            sel.where = self.expr()
        if self.accept_kw("GROUP"):
            self.expect_kw("BY")
            sel.group_by = self.expr_list()
        if self.accept_kw("HAVING"):
            sel.having = self.expr()
        if self.t.is_kw("ORDER"):
            sel.order_by = self.order_by()
        self.accept_kw("WITH") and self.accept_kw("OWNERACCESS") and self.accept_kw("OPTION")
        return sel

    def order_by(self):
        self.expect_kw("ORDER")
        self.expect_kw("BY")
        items = []
        while True:
            e = self.expr()
            desc = False
            if self.accept_kw("DESC"):
                desc = True
            else:
                self.accept_kw("ASC")
            items.append((e, desc))
            if not self.accept_op(","):
                break
        return items

    def select_list(self):
        cols = []
        while True:
            if self.t.kind == "OP" and self.t.value == "*":
                self.next()
                cols.append((A.Star(), None))
            else:
                e = self.expr()
                alias = None
                if self.accept_kw("AS"):
                    alias = self.ident()
                cols.append((e, alias))
            if not self.accept_op(","):
                break
        return cols

    def expr_list(self):
        items = [self.expr()]
        while self.accept_op(","):
            items.append(self.expr())
        return items

    # FROM -----------------------------------------------------------
    def from_clause(self) -> A.Source:
        src = self.join_source()
        while self.accept_op(","):
            src = A.Join("CROSS", src, self.join_source())
        return src

    def join_source(self) -> A.Source:
        left = self.primary_source()
        while True:
            kind = None
            if self.t.is_kw("INNER"):
                kind = "INNER"
            elif self.t.is_kw("LEFT"):
                kind = "LEFT"
            elif self.t.is_kw("RIGHT"):
                kind = "RIGHT"
            elif self.t.is_kw("JOIN"):
                kind = "INNER"
            if kind is None:
                return left
            if not self.t.is_kw("JOIN"):
                self.next()
            self.accept_kw("OUTER")
            self.expect_kw("JOIN")
            right = self.primary_source()
            self.expect_kw("ON")
            on = self.expr()
            left = A.Join(kind, left, right, on)

    def primary_source(self) -> A.Source:
        if self.accept_op("("):
            if self.t.is_kw("SELECT", "TRANSFORM"):
                q = self.select_or_union()
                self.expect_op(")")
                return A.SubSource(q, self.opt_alias())
            src = self.from_clause()
            self.expect_op(")")
            alias = self.opt_alias()
            if alias and isinstance(src, A.TableRef):
                src.alias = alias
            return src
        name = self.ident()
        while self.t.kind == "OP" and self.t.value == "." and self.peek().kind in ("IDENT", "QIDENT"):
            self.next()
            name = name + "." + self.ident()  # owner.tabela (raro)
        database = None
        if self.t.is_kw("IN") and self.peek().kind == "STR":
            self.next()
            database = self.next().value
        return A.TableRef(name, self.opt_alias(), database)

    def opt_alias(self):
        if self.accept_kw("AS"):
            return self.ident()
        if self.t.kind == "QIDENT" or (self.t.kind == "IDENT" and self.t.value.upper() not in _STOP_ALIAS):
            return self.next().value
        return None

    # comandos de ação -----------------------------------------------
    def transform(self) -> A.Transform:
        self.expect_kw("TRANSFORM")
        value = self.expr()
        if self.accept_kw("AS"):
            self.ident()
        sel = self.select_core()
        self.expect_kw("PIVOT")
        pivot = self.expr()
        items = []
        if self.accept_kw("IN"):
            self.expect_op("(")
            items = self.expr_list()
            self.expect_op(")")
        return A.Transform(value=value, select=sel, pivot=pivot, pivot_in=items)

    def update(self) -> A.Update:
        self.expect_kw("UPDATE")
        src = self.from_clause()
        self.expect_kw("SET")
        sets = []
        while True:
            target = self.primary()
            if not isinstance(target, A.Col):
                self.error("esperado nome de coluna no SET")
            self.expect_op("=")
            sets.append((target, self.expr()))
            if not self.accept_op(","):
                break
        where = self.expr() if self.accept_kw("WHERE") else None
        return A.Update(source=src, sets=sets, where=where)

    def delete(self) -> A.Delete:
        self.expect_kw("DELETE")
        self.accept_kw("DISTINCTROW")
        target = None
        if not self.t.is_kw("FROM"):
            # DELETE t.* FROM ...   ou   DELETE * FROM ...
            if self.accept_op("*"):
                pass
            else:
                target = self.ident()
                if self.accept_op("."):
                    self.expect_op("*")
        self.expect_kw("FROM")
        src = self.from_clause()
        where = self.expr() if self.accept_kw("WHERE") else None
        if target is None and isinstance(src, A.TableRef):
            target = src.key
        return A.Delete(target=target, source=src, where=where)

    def insert(self) -> A.Insert:
        self.expect_kw("INSERT")
        self.expect_kw("INTO")
        table = self.ident()
        if self.t.is_kw("IN") and self.peek().kind == "STR":
            self.next(); self.next()
        cols = []
        if self.t.kind == "OP" and self.t.value == "(" and not self.peek().is_kw("SELECT"):
            self.next()
            cols.append(self._col_name())
            while self.accept_op(","):
                cols.append(self._col_name())
            self.expect_op(")")
        if self.accept_kw("VALUES"):
            self.expect_op("(")
            vals = self.expr_list()
            self.expect_op(")")
            return A.Insert(table=table, columns=cols, values=vals)
        return A.Insert(table=table, columns=cols, select=self.select_or_union())

    def _col_name(self) -> str:
        name = self.ident()
        while self.accept_op("."):
            name = self.ident()
        return name

    # expressões (precedência VBA/Jet) ------------------------------
    def expr(self) -> A.Expr:
        return self.imp_expr()

    def imp_expr(self):
        left = self.eqv_expr()
        while self.accept_kw("IMP"):
            left = A.Bin("IMP", left, self.eqv_expr())
        return left

    def eqv_expr(self):
        left = self.xor_expr()
        while self.accept_kw("EQV"):
            left = A.Bin("EQV", left, self.xor_expr())
        return left

    def xor_expr(self):
        left = self.or_expr()
        while self.accept_kw("XOR"):
            left = A.Bin("XOR", left, self.or_expr())
        return left

    def or_expr(self):
        left = self.and_expr()
        while self.accept_kw("OR"):
            left = A.Bin("OR", left, self.and_expr())
        return left

    def and_expr(self):
        left = self.not_expr()
        while self.accept_kw("AND"):
            left = A.Bin("AND", left, self.not_expr())
        return left

    def not_expr(self):
        if self.accept_kw("NOT"):
            return A.Unary("NOT", self.not_expr())
        return self.cmp_expr()

    def cmp_expr(self):
        left = self.concat_expr()
        while True:
            t = self.t
            if t.kind == "OP" and t.value in ("=", "<>", "<", ">", "<=", ">="):
                self.next()
                left = A.Bin(t.value, left, self.concat_expr())
                continue
            neg = False
            save = self.i
            if t.is_kw("NOT"):
                self.next()
                neg = True
            if self.accept_kw("LIKE"):
                left = A.Like(left, self.concat_expr(), neg)
            elif self.accept_kw("BETWEEN"):
                low = self.concat_expr()
                self.expect_kw("AND")
                left = A.Between(left, low, self.concat_expr(), neg)
            elif self.accept_kw("IN"):
                self.expect_op("(")
                if self.t.is_kw("SELECT"):
                    q = self.select_or_union()
                    self.expect_op(")")
                    left = A.InSelect(left, q, neg)
                else:
                    items = self.expr_list()
                    self.expect_op(")")
                    left = A.InList(left, items, neg)
            elif not neg and self.accept_kw("IS"):
                n = self.accept_kw("NOT")
                self.expect_kw("NULL")
                left = A.IsNull(left, n)
            else:
                self.i = save
                return left

    def concat_expr(self):
        left = self.add_expr()
        while self.accept_op("&"):
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
        while self.accept_kw("MOD"):
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
        left = self.primary()
        while self.accept_op("^"):
            if self.accept_op("-"):
                right = A.Unary("-", self.primary())
            else:
                right = self.primary()
            left = A.Bin("^", left, right)
        return left

    def primary(self) -> A.Expr:
        t = self.t
        if t.kind == "NUM":
            self.next()
            v = float(t.value) if any(c in t.value for c in ".eE") else int(t.value)
            return A.Lit(v, "num")
        if t.kind == "STR":
            self.next()
            return A.Lit(t.value, "str")
        if t.kind == "DATE":
            self.next()
            return A.Lit(parse_access_date(t.value), "date")
        if t.kind == "OP" and t.value == "(":
            self.next()
            if self.t.is_kw("SELECT"):
                q = self.select_or_union()
                self.expect_op(")")
                return A.SubqueryExpr(q)
            e = self.expr()
            self.expect_op(")")
            return e
        if t.kind == "OP" and t.value == "*":
            self.next()
            return A.Star()
        if t.kind == "IDENT":
            up = t.value.upper()
            if up == "NULL":
                self.next()
                return A.Lit(None, "null")
            if up in ("TRUE", "YES", "ON") and not self._followed_by_ref():
                self.next()
                return A.Lit(True, "bool")
            if up in ("FALSE", "NO", "OFF") and not self._followed_by_ref():
                self.next()
                return A.Lit(False, "bool")
            if up == "EXISTS":
                self.next()
                self.expect_op("(")
                q = self.select_or_union()
                self.expect_op(")")
                return A.Exists(q)
            # chamada de função
            if self.peek().kind == "OP" and self.peek().value == "(":
                self.next()
                self.next()
                args = []
                distinct = False
                if not self.accept_op(")"):
                    distinct = self.accept_kw("DISTINCT")
                    args = self.expr_list()
                    self.expect_op(")")
                return A.Func(t.value.rstrip("$"), args, distinct)
        if t.kind in ("IDENT", "QIDENT"):
            parts = [self.next().value]
            while self.t.kind == "OP" and self.t.value in (".", "!"):
                self.next()
                if self.t.kind == "OP" and self.t.value == "*":
                    self.next()
                    return A.Star(".".join(parts))
                parts.append(self.ident())
            return A.Col(parts)
        self.error("expressão inválida")

    def _followed_by_ref(self) -> bool:
        p = self.peek()
        return p.kind == "OP" and p.value in (".", "!")


def parse(sql: str) -> A.Statement:
    return Parser(sql).parse()


def parse_expr(text: str) -> A.Expr:
    """Analisa só uma expressão (ex.: critério de DLookup ou FindFirst)."""
    p = Parser(text)
    e = p.expr()
    if p.t.kind != "EOF":
        p.error("texto inesperado após a expressão")
    return e
