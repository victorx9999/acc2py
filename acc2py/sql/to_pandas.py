"""Gera código Python/pandas a partir da AST do SQL do Access.

Cada consulta vira uma função ``q_<nome>(db, p=None)`` que recebe um
``acc2py.runtime.Database`` (tabelas como DataFrames) e um dicionário de
parâmetros, e devolve um DataFrame (consultas de seleção) ou o número de
linhas afetadas (consultas de ação).

Convenção usada no código gerado: dentro de cada etapa, o DataFrame de
trabalho se chama ``f`` e as colunas vindas do FROM são qualificadas como
``"alias.coluna"``. ``rt.ref`` resolve nomes sem distinguir maiúsculas,
cai para os parâmetros quando a coluna não existe (como o Access faz) e
aceita nomes não qualificados.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Optional

from . import ast as A
from .parser import parse
from .unparse import expr_sql, key
from ..diagnostics import Diagnostics
from ..runtime.vbafuncs import CANONICAL as FUNC_NAMES

DOMAIN_FUNCS = {n.lower(): n for n in ("DLookup", "DSum", "DCount", "DAvg", "DMin", "DMax",
                                       "DFirst", "DLast", "DStDev", "DStDevP", "DVar", "DVarP")}

AGG_NAMES_PT = {"sum": "SomaDe", "count": "ContarDe", "avg": "MédiaDe", "min": "MínDe",
                "max": "MáxDe", "first": "PrimeiroDe", "last": "ÚltimoDe", "stdev": "DesvPadDe",
                "stdevp": "DesvPadPDe", "var": "VarDe", "varp": "VarPDe"}
AGG_NAMES_EN = {"sum": "SumOf", "count": "CountOf", "avg": "AvgOf", "min": "MinOf",
                "max": "MaxOf", "first": "FirstOf", "last": "LastOf", "stdev": "StDevOf",
                "stdevp": "StDevPOf", "var": "VarOf", "varp": "VarPOf"}

PY_KEYWORDS = {"False", "None", "True", "and", "as", "assert", "async", "await", "break",
               "class", "continue", "def", "del", "elif", "else", "except", "finally", "for",
               "from", "global", "if", "import", "in", "is", "lambda", "nonlocal", "not", "or",
               "pass", "raise", "return", "try", "while", "with", "yield"}


def py_identifier(name: str, prefix: str = "") -> str:
    """Nome Access -> identificador Python legível (sem acentos, snake_case)."""
    s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    s = re.sub(r"[^0-9a-zA-Z]+", "_", s).strip("_").lower() or "consulta"
    s = prefix + s
    if s[0].isdigit():
        s = "_" + s
    if s in PY_KEYWORDS:
        s += "_"
    return s


class Emitter:
    def __init__(self):
        self.lines: list[str] = []
        self.level = 0
        self.counter = 0

    def emit(self, line: str = ""):
        self.lines.append(("    " * self.level + line) if line else "")

    def tmp(self, base: str) -> str:
        self.counter += 1
        return f"_{base}{self.counter}"

    def indent(self):
        self.level += 1

    def dedent(self):
        self.level -= 1


@dataclass
class Scope:
    """Contexto de compilação de uma expressão."""
    frame: str = "f"
    aliases: set = field(default_factory=set)          # aliases (minúsculos) desta consulta
    outer_aliases: set = field(default_factory=set)    # aliases da consulta externa
    group_map: dict = field(default_factory=dict)      # chave expr -> coluna temporária
    agg_map: dict = field(default_factory=dict)        # chave agregação -> coluna temporária
    grouped: bool = False
    correlated: bool = False


class SQLCompiler:
    def __init__(self, diags: Optional[Diagnostics] = None, locale: str = "pt"):
        self.diags = diags or Diagnostics()
        self.agg_names = AGG_NAMES_PT if locale == "pt" else AGG_NAMES_EN
        self.where = ""
        self.deps: set[str] = set()
        self.em = Emitter()

    # ================================================================ API
    def compile_function(self, name: str, sql: str, func_name: Optional[str] = None) -> str:
        """Compila uma consulta salva em uma função Python completa (texto)."""
        self.where = f"consulta '{name}'"
        self.em = Emitter()
        self.deps = set()
        func_name = func_name or py_identifier(name, "q_")
        em = self.em
        stmt = parse(sql)
        em.emit(f"def {func_name}(db, p=None):")
        em.indent()
        safe = sql.strip().replace("\\", "\\\\").replace('"""', "'''")
        em.emit(f'"""Consulta Access: {name}')
        em.emit("")
        em.emit("SQL original:")
        for ln in safe.splitlines():
            em.emit(("    " + ln.strip()) if ln.strip() else "")
        em.emit('"""')
        if stmt.parameters:
            em.emit("# Parâmetros declarados: " + ", ".join(f"{n} ({t or 'Variant'})" for n, t in stmt.parameters))
        em.emit("p = rt.Params(p)")
        result = self.statement(stmt, Scope(), action_ok=True)
        em.emit(f"return {result}")
        em.dedent()
        return "\n".join(em.lines)

    # ================================================================ comandos
    def statement(self, st: A.Statement, outer: Scope, action_ok=False) -> str:
        if isinstance(st, A.Select):
            return self.select(st, outer)
        if isinstance(st, A.Union):
            return self.union(st, outer)
        if isinstance(st, A.Transform):
            return self.transform(st, outer)
        if not action_ok:
            raise NotImplementedError("comando de ação dentro de subconsulta")
        if isinstance(st, A.Update):
            return self.update(st)
        if isinstance(st, A.Delete):
            return self.delete(st)
        if isinstance(st, A.Insert):
            return self.insert(st)
        raise NotImplementedError(type(st).__name__)

    # ---------------------------------------------------------------- SELECT
    def select(self, sel: A.Select, outer: Scope) -> str:
        em = self.em
        sc = Scope(outer_aliases=outer.aliases | outer.outer_aliases)
        conjuncts = _split_and(sel.where)
        if sel.from_ is None:
            em.emit("f = rt.single_row()")
        else:
            em.emit(f"# FROM {self._source_desc(sel.from_)}")
            var, aliases = self.source(sel.from_, sc, conjuncts)
            sc.aliases |= aliases
            if var != "f":
                em.emit(f"f = {var}")
        if conjuncts:
            cond = _join_and(conjuncts)
            em.emit(f"# WHERE {_short(expr_sql(cond))}")
            em.emit(f"f = rt.where(f, {self.expr(cond, sc)})")

        aggregated = bool(sel.group_by) or any(A.has_aggregate(e) for e, _ in sel.columns) \
            or (sel.having is not None) or any(A.has_aggregate(e) for e, _ in sel.order_by)
        if aggregated:
            self.group(sel.group_by, [e for e, _ in sel.columns] + [sel.having] + [e for e, _ in sel.order_by], sc)
            if sel.having is not None:
                em.emit(f"# HAVING {_short(expr_sql(sel.having))}")
                em.emit(f"f = rt.where(f, {self.expr(sel.having, sc)})")

        cols = self.projection(sel.columns, sc)
        opts = []
        if sel.order_by:
            orders = ", ".join(f"({self.expr(e, sc)}, {d})" for e, d in sel.order_by)
            opts.append(f"order_by=[{orders}]")
        em.emit("out = rt.project(f, [")
        em.indent()
        for c in cols:
            em.emit(c + ",")
        em.dedent()
        if sel.distinct == "DISTINCT":
            opts.append("distinct=True")
        if sel.top is not None:
            top = int(sel.top) if float(sel.top).is_integer() else sel.top
            opts.append(f"top={top}")
            if sel.top_percent:
                opts.append("top_percent=True")
        em.emit("]" + (", " + ", ".join(opts) if opts else "") + ")")
        if sel.distinct == "DISTINCTROW":
            self.diags.info(self.where, "DISTINCTROW ignorado (só tem efeito em junções com tabelas não exibidas)")
        if sel.into:
            em.emit(f"return rt.make_table(db, {sel.into!r}, out)")
            return "None  # (consulta criar-tabela)"
        if sc.correlated:
            outer.correlated = True
        return "out"

    def group(self, group_by, exprs, sc: Scope):
        em = self.em
        keys = []
        for i, g in enumerate(group_by):
            col = f"_g{i}"
            sc.group_map[key(g)] = col
            keys.append(f"{col!r}: {self.expr(g, sc)}")
        aggs = []
        for e in exprs:
            for node in A.walk(e):
                if isinstance(node, A.Func) and node.name.lower() in A.AGGREGATES:
                    k = key(node)
                    if k in sc.agg_map:
                        continue
                    col = f"_a{len(sc.agg_map)}"
                    sc.agg_map[k] = col
                    fn = node.name.lower()
                    if not node.args or isinstance(node.args[0], A.Star):
                        aggs.append(f"{col!r}: ({fn!r}, None)")
                    else:
                        if node.distinct:
                            fn += "_distinct"
                        aggs.append(f"{col!r}: ({fn!r}, {self.expr(node.args[0], sc)})")
        em.emit(f"# GROUP BY {', '.join(_short(expr_sql(g), 40) for g in group_by) or '(total geral)'}")
        em.emit("f = rt.group(f,")
        em.emit(f"    keys={{{', '.join(keys)}}},")
        em.emit(f"    aggs={{{', '.join(aggs)}}})")
        sc.grouped = True
        sc.frame = "f"

    def projection(self, columns, sc: Scope) -> list[str]:
        out = []
        used: dict[str, int] = {}
        expr_n = 1000
        for e, alias in columns:
            if isinstance(e, A.Star):
                out.append(f"rt.STAR({e.table!r})" if e.table else "rt.STAR()")
                continue
            if alias:
                name = alias
            elif isinstance(e, A.Col):
                name = e.parts[-1]
            elif isinstance(e, A.Func) and e.name.lower() in A.AGGREGATES and e.args \
                    and isinstance(e.args[0], A.Col):
                name = self.agg_names[e.name.lower()] + e.args[0].parts[-1]
            elif isinstance(e, A.Func) and e.name.lower() == "count" and (
                    not e.args or isinstance(e.args[0], A.Star)):
                name = f"Expr{expr_n}"
                expr_n += 1
            else:
                name = f"Expr{expr_n}"
                expr_n += 1
            low = name.lower()
            if low in used and isinstance(e, A.Col) and len(e.parts) > 1 and not alias:
                name = ".".join(e.parts)
            elif low in used:
                used[low] += 1
                name = f"{name}{used[low]}"
            used[name.lower()] = used.get(name.lower(), 0)
            out.append(f"({name!r}, {self.expr(e, sc)})")
        return out

    # ---------------------------------------------------------------- FROM
    def _source_desc(self, s: A.Source) -> str:
        if isinstance(s, A.TableRef):
            return s.name + (f" AS {s.alias}" if s.alias else "")
        if isinstance(s, A.SubSource):
            return f"(subconsulta) AS {s.alias}"
        return f"{self._source_desc(s.left)} {s.kind} JOIN {self._source_desc(s.right)}"

    def source(self, s: A.Source, sc: Scope, conjuncts: list, rowid=False) -> tuple[str, set]:
        em = self.em
        if isinstance(s, A.TableRef):
            v = em.tmp("t")
            if s.database:
                self.diags.warn(self.where, f"tabela externa IN '{s.database}' tratada como tabela local '{s.name}'")
            self.deps.add(s.name)
            extra = ", rowid=True" if rowid else ""
            em.emit(f"{v} = rt.source(db, {s.name!r}, {s.key!r}, p{extra})")
            return v, {s.key.lower()}
        if isinstance(s, A.SubSource):
            fn = self.subquery_function(s.query, sc)
            v = em.tmp("t")
            em.emit(f"{v} = rt.as_source({fn}(), {s.key!r})")
            return v, {s.key.lower()}
        if isinstance(s, A.Join):
            lv, la = self.source(s.left, sc, conjuncts, rowid)
            rv, ra = self.source(s.right, sc, conjuncts, rowid)
            kind = {"INNER": "inner", "LEFT": "left", "RIGHT": "right", "CROSS": "inner"}[s.kind]
            on_parts = _split_and(s.on) if s.on is not None else []
            if s.kind == "CROSS":
                # junção implícita (FROM a, b WHERE a.x = b.y): puxa as igualdades do WHERE
                for c in list(conjuncts):
                    if _equi_pair(c, la, ra):
                        on_parts.append(c)
                        conjuncts.remove(c)
            keys, rest = [], []
            for c in on_parts:
                pair = _equi_pair(c, la, ra)
                if pair:
                    keys.append(pair)
                else:
                    rest.append(c)
            v = em.tmp("j")
            args = [lv, rv, repr(kind)]
            if keys:
                args.append("on=[" + ", ".join(f"({a!r}, {b!r})" for a, b in keys) + "]")
            if rest:
                inner = Scope(frame="f", aliases=sc.aliases | la | ra, outer_aliases=sc.outer_aliases)
                args.append(f"cond=lambda f: {self.expr(_join_and(rest), inner)}")
            if not keys and not rest and s.kind != "CROSS":
                self.diags.warn(self.where, "JOIN sem condição reconhecida; virou produto cartesiano")
            em.emit(f"{v} = rt.join({', '.join(args)})")
            return v, la | ra
        raise NotImplementedError(type(s).__name__)

    def subquery_function(self, st: A.Statement, sc: Scope, correlated_arg=False) -> str:
        """Emite uma função interna para a subconsulta e devolve o nome dela."""
        em = self.em
        fn = em.tmp("sub")
        start = len(em.lines)
        em.emit(f"def {fn}(o=None):")
        em.indent()
        inner_scope = Scope(aliases=set(), outer_aliases=sc.aliases | sc.outer_aliases)
        res = self.statement(st, inner_scope)
        em.emit(f"return {res}")
        em.dedent()
        if inner_scope.correlated:
            em.lines[start] = em.lines[start].replace("(o=None)", "(o)")
            sc.correlated = sc.correlated  # o chamador decide como usar
        self._last_correlated = inner_scope.correlated
        return fn

    # ---------------------------------------------------------------- UNION / TRANSFORM
    def union(self, u: A.Union, outer: Scope) -> str:
        em = self.em
        fns = []
        for part in u.parts:
            fns.append(self.subquery_function(part, outer))
        flags = ", ".join(str(f) for f in u.all_flags)
        em.emit(f"out = rt.union([{', '.join(fn + '()' for fn in fns)}], all_flags=[{flags}])")
        if u.order_by:
            sc = Scope(frame="out")
            orders = ", ".join(f"({self.expr(e, sc)}, {d})" for e, d in u.order_by)
            em.emit(f"out = rt.order(out, [{orders}])")
        return "out"

    def transform(self, t: A.Transform, outer: Scope) -> str:
        em = self.em
        sel = t.select
        sc = Scope()
        conjuncts = _split_and(sel.where)
        em.emit(f"# FROM {self._source_desc(sel.from_)}")
        var, aliases = self.source(sel.from_, sc, conjuncts)
        sc.aliases |= aliases
        em.emit(f"f = {var}")
        if conjuncts:
            em.emit(f"f = rt.where(f, {self.expr(_join_and(conjuncts), sc)})")
        rows, extras = [], []
        group_keys = {key(g): g for g in sel.group_by}
        named = set()
        for e, alias in sel.columns:
            k = key(e)
            if k in group_keys:
                name = alias or (e.parts[-1] if isinstance(e, A.Col) else f"Expr{1000 + len(rows)}")
                rows.append(f"{name!r}: {self.expr(e, sc)}")
                named.add(k)
            elif isinstance(e, A.Func) and e.name.lower() in A.AGGREGATES:
                arg = e.args[0] if e.args else A.Star()
                name = alias or (self.agg_names[e.name.lower()] + arg.parts[-1] if isinstance(arg, A.Col) else "Total")
                val = "None" if isinstance(arg, A.Star) else self.expr(arg, sc)
                extras.append(f"{name!r}: ({e.name.lower()!r}, {val})")
            else:
                self.diags.warn(self.where, f"coluna de referência cruzada não suportada: {expr_sql(e)}")
        for k, g in group_keys.items():
            if k not in named:
                rows.append(f"'_g{len(rows)}': {self.expr(g, sc)}")
        if not (isinstance(t.value, A.Func) and t.value.name.lower() in A.AGGREGATES):
            raise NotImplementedError("TRANSFORM sem função de agregação simples")
        va = t.value.args[0] if t.value.args else A.Star()
        vexpr = "None" if isinstance(va, A.Star) else self.expr(va, sc)
        pivot = self.expr(t.pivot, sc)
        pin = ", ".join(self.expr(i, Scope()) for i in t.pivot_in)
        em.emit("# TRANSFORM … PIVOT (referência cruzada)")
        em.emit("out = rt.crosstab(f,")
        em.emit(f"    rows={{{', '.join(rows)}}},")
        em.emit(f"    pivot={pivot},")
        em.emit(f"    value=({t.value.name.lower()!r}, {vexpr}),")
        em.emit(f"    extras={{{', '.join(extras)}}},")
        em.emit(f"    pivot_in=[{pin}])")
        if sel.order_by:
            osc = Scope(frame="out")
            em.emit(f"out = rt.order(out, [{', '.join(f'({self.expr(e, osc)}, {d})' for e, d in sel.order_by)}])")
        return "out"

    # ---------------------------------------------------------------- ação
    def _target(self, src: A.Source, alias_hint: Optional[str]) -> A.TableRef:
        tables = _tables(src)
        if alias_hint:
            for t in tables:
                if t.key.lower() == alias_hint.lower() or t.name.lower() == alias_hint.lower():
                    return t
        return tables[0]

    def _action_from(self, src: A.Source, where, sc: Scope):
        em = self.em
        conjuncts = _split_and(where)
        var, aliases = self.source(src, sc, conjuncts, rowid=True)
        sc.aliases |= aliases
        em.emit(f"f = {var}")
        if conjuncts:
            em.emit(f"f = rt.where(f, {self.expr(_join_and(conjuncts), sc)})")

    def update(self, u: A.Update) -> str:
        em = self.em
        sc = Scope()
        self._action_from(u.source, u.where, sc)
        hint = next((c.parts[-2] for c, _ in u.sets if len(c.parts) > 1), None)
        target = self._target(u.source, hint)
        sets = ", ".join(f"{c.parts[-1]!r}: {self.expr(v, sc)}" for c, v in u.sets)
        em.emit(f"return rt.apply_update(db, {target.name!r}, f, {target.key!r}, {{{sets}}})")
        return "None  # (consulta atualização)"

    def delete(self, d: A.Delete) -> str:
        em = self.em
        sc = Scope()
        self._action_from(d.source, d.where, sc)
        target = self._target(d.source, d.target)
        em.emit(f"return rt.apply_delete(db, {target.name!r}, f, {target.key!r})")
        return "None  # (consulta exclusão)"

    def insert(self, ins: A.Insert) -> str:
        em = self.em
        cols = repr(ins.columns) if ins.columns else "None"
        if ins.values is not None:
            vals = ", ".join(self.expr(v, Scope()) for v in ins.values)
            em.emit(f"return rt.apply_insert_values(db, {ins.table!r}, {cols}, [{vals}])")
        else:
            res = self.statement(ins.select, Scope())
            em.emit(f"return rt.apply_insert(db, {ins.table!r}, {res}, {cols})")
        return "None  # (consulta acréscimo)"

    # ================================================================ expressões
    def expr(self, e: A.Expr, sc: Scope) -> str:
        if sc.grouped:
            k = key(e)
            if k in sc.agg_map:
                return f'{sc.frame}["{sc.agg_map[k]}"]'
            if k in sc.group_map:
                return f'{sc.frame}["{sc.group_map[k]}"]'
            if isinstance(e, A.Col):
                # "Nome" pode casar com "c.Nome" do GROUP BY (e vice-versa)
                for gk, col in sc.group_map.items():
                    if gk.split(".")[-1].strip("[]") == e.parts[-1].lower():
                        return f'{sc.frame}["{col}"]'
                self.diags.warn(self.where, f"coluna '{expr_sql(e)}' não está no GROUP BY nem em agregação")
        f = sc.frame
        if isinstance(e, A.Col):
            name = ".".join(e.parts)
            if len(e.parts) > 1 and e.parts[0].lower() not in sc.aliases \
                    and e.parts[0].lower() in sc.outer_aliases:
                sc.correlated = True
                return f"rt.ref(o, {name!r}, p)"
            return f"rt.ref({f}, {name!r}, p)"
        if isinstance(e, A.Lit):
            if e.kind == "date":
                return f'pd.Timestamp("{e.value.isoformat(sep=" ")}")'
            return repr(e.value)
        if isinstance(e, A.Star):
            raise NotImplementedError("'*' fora de SELECT/Count")
        if isinstance(e, A.Unary):
            if e.op == "NOT":
                return f"rt.not_({self.expr(e.operand, sc)})"
            return f"-{self._paren(e.operand, sc)}"
        if isinstance(e, A.Bin):
            a, b = self.expr(e.left, sc), self.expr(e.right, sc)
            op = e.op
            if op in ("+", "-", "*", "/"):
                return f"({a} {op} {b})"
            if op == "^":
                return f"({a} ** {b})"
            named = {"\\": "idiv", "MOD": "mod", "&": "concat", "=": "eq", "<>": "ne", "<": "lt",
                     ">": "gt", "<=": "le", ">=": "ge", "AND": "and_", "OR": "or_", "XOR": "xor_",
                     "EQV": "eqv", "IMP": "imp"}[op]
            return f"rt.{named}({a}, {b})"
        if isinstance(e, A.Func):
            return self.func(e, sc)
        if isinstance(e, A.Between):
            r = f"rt.between({self.expr(e.expr, sc)}, {self.expr(e.low, sc)}, {self.expr(e.high, sc)})"
            return f"rt.not_({r})" if e.negated else r
        if isinstance(e, A.InList):
            r = f"rt.isin({self.expr(e.expr, sc)}, [{', '.join(self.expr(i, sc) for i in e.items)}])"
            return f"rt.not_({r})" if e.negated else r
        if isinstance(e, A.IsNull):
            return f"rt.{'notnull' if e.negated else 'isnull'}({self.expr(e.expr, sc)})"
        if isinstance(e, A.Like):
            r = f"rt.like({self.expr(e.expr, sc)}, {self.expr(e.pattern, sc)})"
            return f"rt.not_({r})" if e.negated else r
        if isinstance(e, (A.InSelect, A.Exists, A.SubqueryExpr)):
            fn = self.subquery_function(e.query, sc)
            corr = self._last_correlated
            if isinstance(e, A.SubqueryExpr):
                return f"rt.per_row({f}, {fn}, 'scalar')" if corr else f"rt.scalar({fn}())"
            if isinstance(e, A.Exists):
                r = f"rt.per_row({f}, {fn}, 'exists')" if corr else f"rt.exists({fn}())"
                return f"rt.not_({r})" if e.negated else r
            x = self.expr(e.expr, sc)
            r = f"rt.per_row_in({f}, {x}, {fn})" if corr else f"rt.isin({x}, rt.first_col({fn}()))"
            return f"rt.not_({r})" if e.negated else r
        raise NotImplementedError(type(e).__name__)

    def _paren(self, e, sc):
        s = self.expr(e, sc)
        return s if s.startswith("(") or not any(c in s for c in " +-*/") else f"({s})"

    def func(self, e: A.Func, sc: Scope) -> str:
        low = e.name.lower()
        args = [self.expr(a, sc) for a in e.args if not isinstance(a, A.Star)]
        if low in A.AGGREGATES:
            self.diags.error(self.where, f"agregação {e.name}() fora de contexto de agrupamento")
            return f"rt.fail({('agregação ' + e.name)!r})"
        if low in DOMAIN_FUNCS:
            if len(e.args) >= 2 and isinstance(e.args[1], A.Lit) and e.args[1].kind == "str":
                self.deps.add(e.args[1].value)
            canonical = DOMAIN_FUNCS[low]
            return f"F.{canonical}(db, {', '.join(args)})"
        if low in FUNC_NAMES:
            return f"F.{FUNC_NAMES[low]}({', '.join(args)})"
        self.diags.info(self.where, f"função '{e.name}' não é nativa do Access: chamada via db.functions (VBA transpilado)")
        return f"F.udf(db, {e.name!r}{''.join(', ' + a for a in args)})"


# ---------------------------------------------------------------- utilitários
def _split_and(e) -> list:
    if e is None:
        return []
    if isinstance(e, A.Bin) and e.op == "AND":
        return _split_and(e.left) + _split_and(e.right)
    return [e]


def _join_and(items):
    e = items[0]
    for x in items[1:]:
        e = A.Bin("AND", e, x)
    return e


def _qualified(c) -> Optional[str]:
    if isinstance(c, A.Col) and len(c.parts) >= 2:
        return c.parts[-2].lower()
    return None


def _equi_pair(c, la: set, ra: set):
    """Se ``c`` é ``l.x = r.y`` com l à esquerda e r à direita, devolve as colunas."""
    if not (isinstance(c, A.Bin) and c.op == "="):
        return None
    ql, qr = _qualified(c.left), _qualified(c.right)
    if ql is None or qr is None:
        return None
    a = ".".join(c.left.parts[-2:])
    b = ".".join(c.right.parts[-2:])
    if ql in la and qr in ra:
        return (a, b)
    if ql in ra and qr in la:
        return (b, a)
    return None


def _tables(src) -> list[A.TableRef]:
    if isinstance(src, A.TableRef):
        return [src]
    if isinstance(src, A.Join):
        return _tables(src.left) + _tables(src.right)
    return []


def _short(s: str, n: int = 90) -> str:
    s = " ".join(s.split())
    return s if len(s) <= n else s[: n - 1] + "…"


def source_dependencies(sql: str) -> set[str]:
    """Nomes de tabelas/consultas usados por uma consulta (para fechar dependências)."""
    comp = SQLCompiler()
    try:
        comp.compile_function("_", sql)
    except Exception:
        return set()
    return comp.deps
