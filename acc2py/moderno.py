"""Gerador de código pandas "moderno" (idiomático e sem o runtime do acc2py).

A versão fiel (``consultas.py``) reproduz o Access nos detalhes usando o
runtime. Esta gera código que um time de engenharia mantém: uma função por
consulta, ``merge``/``groupby``/``loc``, parâmetros tipados e nomes legíveis.

O preço é que algumas regras do Access deixam de ser reproduzidas (listadas em
``DIFERENCAS``). Por isso o pacote moderno vem com ``comparar_com_fiel``, que
roda as duas versões com os mesmos dados e aponta qualquer divergência.

O que não dá para modernizar com segurança (subconsultas, DLookup, junções sem
chave, colunas de origem desconhecidas…) usa a versão fiel naquela etapa, e o
código diz isso explicitamente.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

from .sql import ast as A
from .sql.parser import parse
from .sql.to_pandas import AGG_NAMES_PT, AGG_NAMES_EN, PY_KEYWORDS
from .sql.unparse import expr_sql, key

DIFERENCAS = [
    "Texto: comparações (=, <>, IN, GROUP BY, JOIN, ORDER BY) DIFERENCIAM maiúsculas; no Access não.",
    "JOIN: chaves vazias (Null) casam entre si no pandas; no Access nunca casam.",
    "TOP n: corta em n linhas; o Access inclui os empates na última posição.",
    "Sum de um grupo só com valores vazios dá 0; no Access dá Null.",
    "First/Last ignoram valores vazios; no Access pegam o primeiro/último registro mesmo vazio.",
    "Concatenação (&) de números usa ponto decimal (1234.5); o Access usa a configuração regional (1234,5).",
    "ORDER BY com direções misturadas (ASC e DESC) põe vazios no fim; o Access põe vazios primeiro no ASC.",
    "Round: valores que caem exatamente no meio (ex.: 58,075) podem diferir na última casa, porque o "
    "pandas arredonda o número binário e o Access o decimal.",
]


class NaoModernizavel(Exception):
    """A consulta usa algo que o gerador moderno não converte com segurança."""


def ident(name: str, prefix: str = "") -> str:
    s = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^0-9a-zA-Z]+", "_", s).strip("_").lower() or "x"
    s = prefix + s
    if s[0].isdigit():
        s = "_" + s
    if s in PY_KEYWORDS or s in {"pd", "np", "df", "g", "t", "p", "resultado", "filtro"}:
        s += "_"
    return s


def q(v) -> str:
    """Literal Python com aspas duplas (str, listas, dicts)."""
    import json
    if isinstance(v, str):
        return json.dumps(v, ensure_ascii=False)
    if isinstance(v, (list, tuple)):
        return "[" + ", ".join(q(x) for x in v) + "]"
    if isinstance(v, dict):
        return "{" + ", ".join(f"{q(k)}: {q(x)}" for k, x in v.items()) + "}"
    return repr(v)


def func_name(nome: str) -> str:
    """'01 Cria base' -> 'cria_base' (a numeração da etapa sai do nome da função)."""
    sem_num = re.sub(r"^\s*\d+[\s._)-]*", "", nome)
    return ident(sem_num or nome, "" if sem_num else "etapa_")


def _lit_py(v):
    if isinstance(v, datetime):
        if v.time() == v.min.time():
            return f'pd.Timestamp("{v:%Y-%m-%d}")'
        return f'pd.Timestamp("{v:%Y-%m-%d %H:%M:%S}")'
    return q(v)


TIPO_KIND = {"texto": "str", "numero": "num", "inteiro": "num", "data": "date", "logico": "bool"}


@dataclass
class V:
    """Código de uma expressão: texto, se é coluna (Series) e o tipo provável."""
    code: str
    series: bool
    kind: str = "any"  # num | str | date | bool | any


@dataclass
class Fonte:
    alias: str
    nome: str            # tabela ou consulta
    var: str             # nome do parâmetro da função
    cols: list           # colunas
    kinds: dict          # coluna(lower) -> kind


@dataclass
class Info:
    nome: str
    func: str
    tipo: str                              # seleção | criar tabela | atualização | exclusão | acréscimo
    fontes: list = field(default_factory=list)   # nomes (tabelas/consultas) na ordem dos parâmetros
    params: list = field(default_factory=list)   # nomes originais de parâmetros
    cols: Optional[list] = None
    kinds: dict = field(default_factory=dict)
    alvo: Optional[str] = None             # tabela criada/alterada
    codigo: str = ""
    moderna: bool = True
    motivo: str = ""


class Ctx:
    """Resolução de colunas dentro de uma consulta."""

    def __init__(self, gen: "GeradorModerno", frame="df"):
        self.gen = gen
        self.frame = frame
        self.colmap: dict = {}      # (alias_lower, col_lower) -> nome da coluna no frame
        self.kinds: dict = {}       # nome da coluna no frame -> kind
        self.params: list = []
        self.agg: dict = {}         # chave da expressão -> coluna agregada
        self.pre_colmap: dict = {}
        self.group_cols: set = set()
        self.grouped = False

    def add_source(self, f: Fonte, renamed: dict):
        for c in f.cols:
            col = renamed.get(c, c)
            self.colmap[(f.alias.lower(), c.lower())] = col
            self.kinds[col] = f.kinds.get(c.lower(), "any")

    def frame_cols(self) -> set:
        return set(self.colmap.values())

    def resolve(self, parts: list) -> Optional[str]:
        if len(parts) >= 2:
            k = (parts[-2].lower(), parts[-1].lower())
            if k in self.colmap:
                return self.colmap[k]
            return None
        c = parts[0].lower()
        hits = sorted({v for (a, cl), v in self.colmap.items() if cl == c})
        if len(hits) > 1:
            raise NaoModernizavel(f"coluna ambígua '{parts[0]}' (qualifique com a tabela)")
        return hits[0] if hits else None

    def param(self, parts: list) -> V:
        nome = "!".join(parts)
        if nome.lower() not in [x.lower() for x in self.params]:
            self.params.append(nome)
        return V(ident(nome), False, self.gen.param_kind(nome))


class GeradorModerno:
    def __init__(self, sql: dict, esquemas: dict, udfs: Optional[dict] = None, locale: str = "pt"):
        self.sql = sql                                  # nome -> SQL (todas as consultas transpiladas)
        self.lower = {k.lower(): k for k in sql}
        self.esquemas = {k.lower(): v for k, v in esquemas.items()}   # tabela -> {col: tipo}
        self.esquema_nomes = {k.lower(): k for k in esquemas}
        self.udfs = {k.lower(): v for k, v in (udfs or {}).items()}   # função VBA -> "Modulo.Funcao"
        self.agg_names = AGG_NAMES_PT if locale.startswith("pt") else AGG_NAMES_EN
        self.infos: dict[str, Info] = {}
        self.param_tipos: dict = {}
        self.criadas: dict = {}                          # tabela criada (lower) -> consulta
        self.funcs_usadas: set = set()
        self.usa_np = False
        self.modulos_usados: set = set()

    # ------------------------------------------------------------- utilidades
    def param_kind(self, nome):
        t = self.param_tipos.get(nome.lower(), "")
        t = t.lower()
        if t.startswith("datetime"):
            return "date"
        if t in ("long", "short", "byte", "currency", "ieeedouble", "ieeesingle", "double", "decimal"):
            return "num"
        if t.startswith("text"):
            return "str"
        return "any"

    def fonte_info(self, nome: str):
        """(colunas, kinds) de uma tabela ou consulta; None se desconhecido."""
        low = nome.lower()
        if low in self.lower:
            info = self.info(self.lower[low])
            if info.cols is None:
                return None
            return info.cols, info.kinds
        if low in self.criadas:
            info = self.info(self.criadas[low])
            if info.cols is None:
                return None
            return info.cols, info.kinds
        esq = self.esquemas.get(low)
        if esq:
            return list(esq), {c.lower(): TIPO_KIND.get(t, "any") for c, t in esq.items()}
        return None

    # ------------------------------------------------------------- entrada
    def info(self, nome: str) -> Info:
        if nome in self.infos:
            return self.infos[nome]
        fn, k = func_name(nome), 2
        usados = {i.func for i in self.infos.values()}
        base = fn
        while fn in usados:
            fn, k = f"{base}_{k}", k + 1
        info = Info(nome, fn, "seleção")
        self.infos[nome] = info  # evita ciclo
        try:
            st = parse(self.sql[nome])
        except Exception as e:  # noqa: BLE001
            info.moderna, info.motivo = False, f"SQL não analisado: {e}"
            return info
        for p, t in st.parameters:
            self.param_tipos[p.lower()] = t
        try:
            self._gerar(info, st)
        except NaoModernizavel as e:
            info.moderna, info.motivo = False, str(e)
            info.codigo = ""
        except Exception as e:  # noqa: BLE001
            info.moderna, info.motivo = False, f"falha ao gerar: {e}"
            info.codigo = ""
        if not info.moderna and info.tipo == "seleção" and info.cols is None:
            pass
        return info

    def _gerar(self, info: Info, st):
        L = []
        if isinstance(st, A.Select):
            if st.into:
                info.tipo, info.alvo = "criar tabela", st.into
                self.criadas[st.into.lower()] = info.nome
            ctx = Ctx(self)
            body, cols, kinds = self._select(st, ctx, info)
            info.cols, info.kinds = cols, kinds
            info.params = ctx.params
            L += body
            L.append("return resultado")
        elif isinstance(st, A.Union):
            ctx = Ctx(self)
            L += self._union(st, ctx, info)
        elif isinstance(st, A.Update):
            info.tipo = "atualização"
            L += self._update(st, info)
        elif isinstance(st, A.Delete):
            info.tipo = "exclusão"
            L += self._delete(st, info)
        elif isinstance(st, A.Insert):
            info.tipo = "acréscimo"
            L += self._insert(st, info)
        else:
            raise NaoModernizavel(f"{type(st).__name__} (referência cruzada/outros) ainda não modernizado")
        info.codigo = self._funcao(info, _encadear(L))

    def _funcao(self, info: Info, body: list) -> str:
        args = [f"{ident(f)}: pd.DataFrame" for f in info.fontes]
        args += [f"{ident(p)}" for p in info.params]
        sql = self.sql[info.nome].strip().replace('"""', "'''").replace("\\", "\\\\")
        doc = [f'"""{info.nome} ({info.tipo}).', "", "SQL original:"]
        doc += ["    " + ln.strip() for ln in sql.splitlines() if ln.strip()]
        doc.append('"""')
        ret = "pd.DataFrame"
        head = f"def {info.func}({', '.join(args)}) -> {ret}:"
        if len(head) > 100:
            head = f"def {info.func}(\n" + "".join(f"    {a},\n" for a in args) + f") -> {ret}:"
        lines = [head] + ["    " + d if d else "" for d in doc] + ["    " + ln if ln else "" for ln in body]
        return "\n".join(lines)

    # ------------------------------------------------------------- FROM
    def _fontes(self, src, ctx: Ctx, info: Info, where_conj: list) -> list:
        """Gera o código que monta 'df' a partir do FROM."""
        L = []
        if isinstance(src, A.TableRef):
            f = self._fonte(src, info)
            ctx.add_source(f, {})
            L.append(f"df = {f.var}.copy()" if info.tipo != "seleção" else f"df = {f.var}")
            return L
        if isinstance(src, A.SubSource):
            raise NaoModernizavel("subconsulta no FROM")
        if not isinstance(src, A.Join):
            raise NaoModernizavel("FROM não suportado")
        # achata a árvore de junções da esquerda para a direita
        seq = []

        def walk(s):
            if isinstance(s, A.Join):
                walk(s.left)
                if not isinstance(s.right, A.TableRef):
                    raise NaoModernizavel("junção com subconsulta ou parênteses à direita")
                seq.append((s.kind, s.right, s.on))
            elif isinstance(s, A.TableRef):
                seq.append((None, s, None))
            else:
                raise NaoModernizavel("subconsulta no FROM")
        walk(src)
        _, first, _ = seq[0]
        f0 = self._fonte(first, info)
        ctx.add_source(f0, {})
        L.append(f"df = {f0.var}")
        aliases_left = {f0.alias.lower()}
        for kind, ref, on in seq[1:]:
            f = self._fonte(ref, info)
            conds = _split_and(on)
            if kind == "CROSS":
                for c in list(where_conj):
                    if _equi(c, aliases_left, {f.alias.lower()}):
                        conds.append(c)
                        where_conj.remove(c)
            pairs, extra = [], []
            for c in conds:
                pr = _equi(c, aliases_left, {f.alias.lower()})
                if pr:
                    pairs.append(pr)
                else:
                    extra.append(c)
            if not pairs:
                raise NaoModernizavel("junção sem chave de igualdade")
            if extra and kind in ("LEFT", "RIGHT"):
                raise NaoModernizavel("LEFT/RIGHT JOIN com condição que não é igualdade")
            # renomeia colunas do lado direito que já existem à esquerda
            existentes = {c.lower() for c in ctx.frame_cols()}
            ren = {}
            for c in f.cols:
                if c.lower() in existentes:
                    novo = f"{c}_{f.alias}"
                    while novo.lower() in existentes:
                        novo += "_"
                    ren[c] = novo
            left_on, right_on = [], []
            for (la, lc), (ra, rc) in pairs:
                lcol = ctx.resolve([la, lc])
                if lcol is None:
                    raise NaoModernizavel(f"coluna '{la}.{lc}' não encontrada")
                if rc.lower() not in {c.lower() for c in f.cols}:
                    raise NaoModernizavel(f"coluna '{ra}.{rc}' não encontrada em {f.nome}")
                rreal = next(c for c in f.cols if c.lower() == rc.lower())
                left_on.append(lcol)
                right_on.append(ren.get(rreal, rreal))
            right = f.var + (f".rename(columns={q(ren)})" if ren else "")
            how = {"INNER": "inner", "LEFT": "left", "RIGHT": "right", "CROSS": "inner"}[kind]
            on = (f'left_on={q(left_on[0])}, right_on={q(right_on[0])}' if len(left_on) == 1
                  else f"left_on={q(left_on)}, right_on={q(right_on)}")
            L.append(f'df = df.merge({right}, how="{how}", {on})')
            ctx.add_source(f, ren)
            aliases_left.add(f.alias.lower())
            for c in extra:
                v = self.ex(c, ctx)
                L.append(f"df = df[{v.code}]")
        return L

    def _fonte(self, ref: A.TableRef, info: Info) -> Fonte:
        if ref.database:
            raise NaoModernizavel(f"tabela de outro banco (IN '{ref.database}')")
        r = self.fonte_info(ref.name)
        if r is None:
            if ref.name.lower() in self.lower:
                dep = self.info(self.lower[ref.name.lower()])
                raise NaoModernizavel(f"depende de '{ref.name}', que não foi modernizada"
                                      + (f" ({dep.motivo})" if dep.motivo else ""))
            raise NaoModernizavel(f"colunas de '{ref.name}' desconhecidas")
        cols, kinds = r
        nome = self._nome_real(ref.name)
        if nome not in info.fontes:
            info.fontes.append(nome)
        return Fonte(ref.alias or ref.name, nome, ident(nome), list(cols), dict(kinds))

    def _nome_real(self, nome):
        low = nome.lower()
        if low in self.lower:
            return self.lower[low]
        return self.esquema_nomes.get(low, nome)

    # ------------------------------------------------------------- SELECT
    def _select(self, st: A.Select, ctx: Ctx, info: Info):
        if st.from_ is None:
            raise NaoModernizavel("SELECT sem FROM")
        conj = _split_and(st.where)
        L = self._fontes(st.from_, ctx, info, conj)
        if conj:
            cond = self.ex(_join_and(conj), ctx)
            L.append(f"df = df[{_strip(cond.code)}]")
        aggregated = bool(st.group_by) or st.having is not None or any(
            A.has_aggregate(e) for e, _ in st.columns)
        out_names, out_exprs = self._nomes_saida(st.columns, ctx)
        if aggregated:
            prefer = {key(e): n for n, e in zip(out_names, out_exprs)
                      if n is not None and isinstance(e, A.Func) and e.name.lower() in A.AGGREGATES}
            L += self._group(st, ctx, prefer)
        cols_code, kinds = [], {}
        for (name, e) in zip(out_names, out_exprs):
            if name is None:  # estrela
                for c in self._estrela(e, ctx):
                    cols_code.append((c, V(f'{ctx.frame}["{c}"]', True, ctx.kinds.get(c, "any"))))
                continue
            v = self.ex(e, ctx)
            cols_code.append((name, v))
        names = [n for n, _ in cols_code]
        for n, v in cols_code:
            kinds[n.lower()] = v.kind
        # ORDER BY: por coluna de saída quando possível, senão coluna auxiliar
        sort_by, asc, extra_cols = [], [], []
        for i, (e, desc) in enumerate(st.order_by):
            nome = self._ordem_por_saida(e, st.columns, out_names, names)
            if nome is None:
                aux = f"_ordem{i}"
                extra_cols.append((aux, self.ex(e, ctx)))
                nome = aux
            sort_by.append(nome)
            asc.append(not desc)
        simples = self._selecao_simples(cols_code + extra_cols, ctx.frame)
        frame_code = [simples] if simples else self._dataframe(cols_code + extra_cols, ctx.frame)
        if len(frame_code) == 1:
            L.append(f"resultado = {frame_code[0]}")
        else:
            L.append("resultado = " + frame_code[0])
            L += frame_code[1:]
        if st.distinct == "DISTINCT":
            L.append("resultado = resultado.drop_duplicates()")
        if sort_by:
            na = '"first"' if all(asc) else '"last"'
            by, a_ = (q(sort_by[0]), asc[0]) if len(sort_by) == 1 else (q(sort_by), q(asc))
            L.append(f"resultado = resultado.sort_values({by}, ascending={a_}, na_position={na}, "
                     f'kind="stable", ignore_index=True)')
        if extra_cols:
            L.append(f"resultado = resultado.drop(columns={q([c for c, _ in extra_cols])})")
        if st.top is not None:
            if st.top_percent:
                L.append(f"resultado = resultado.head(math.ceil(len(resultado) * {st.top} / 100))")
                self.funcs_usadas.add("math")
            else:
                L.append(f"resultado = resultado.head({int(st.top)})")
        if not sort_by:
            L.append("resultado = resultado.reset_index(drop=True)")
        return L, names, kinds

    @staticmethod
    def _selecao_simples(cols: list, frame: str) -> Optional[str]:
        """Só colunas existentes (talvez renomeadas): df[[...]].rename(...) em vez de pd.DataFrame."""
        srcs = []
        for n, v in cols:
            m = re.fullmatch(re.escape(frame) + r'\["([^"\\]*)"\]', v.code)
            if not m:
                return None
            srcs.append(m.group(1))
        if len(set(srcs)) != len(srcs):
            return None
        code = f"{frame}[{q(srcs)}]"
        ren = {a: n for (n, _), a in zip(cols, srcs) if a != n}
        if ren:
            code += f".rename(columns={q(ren)})"
        return code

    def _nomes_saida(self, columns, ctx):
        names, exprs, used = [], [], {}
        n = 1000
        for e, alias in columns:
            if isinstance(e, A.Star):
                names.append(None)
                exprs.append(e)
                continue
            if alias:
                name = alias
            elif isinstance(e, A.Col):
                name = e.parts[-1]
            elif isinstance(e, A.Func) and e.name.lower() in A.AGGREGATES and e.args and isinstance(e.args[0], A.Col):
                name = self.agg_names[e.name.lower()] + e.args[0].parts[-1]
            else:
                name = f"Expr{n}"
                n += 1
            low = name.lower()
            if low in used and isinstance(e, A.Col) and len(e.parts) > 1 and not alias:
                name = ".".join(e.parts)
            elif low in used:
                used[low] += 1
                name = f"{name}{used[low]}"
            used[name.lower()] = used.get(name.lower(), 0)
            names.append(name)
            exprs.append(e)
        return names, exprs

    def _estrela(self, e: A.Star, ctx: Ctx) -> list:
        if ctx.grouped:
            raise NaoModernizavel("* com GROUP BY")
        if e.table:
            return [v for (a, c), v in ctx.colmap.items() if a == e.table.lower()]
        seen = []
        for v in ctx.colmap.values():
            if v not in seen:
                seen.append(v)
        return seen

    def _ordem_por_saida(self, e, columns, out_names, names):
        if isinstance(e, A.Col) and len(e.parts) == 1:
            for n in names:
                if n.lower() == e.parts[0].lower():
                    return n
        k = key(e)
        for (ce, _), n in zip(columns, out_names):
            if n is not None and not isinstance(ce, A.Star) and key(ce) == k:
                return n
        return None

    def _dataframe(self, cols: list, frame: str) -> list:
        if not cols:
            raise NaoModernizavel("SELECT sem colunas")
        all_series = all(v.series for _, v in cols)
        idx = "" if all_series else f", index={frame}.index"
        lines = ["pd.DataFrame({"]
        for n, v in cols:
            lines.append(f"    {q(n)}: {_strip(v.code)},")
        lines.append(f"}}{idx})")
        if len(lines) <= 3 and len(lines[1]) < 70:
            return ["pd.DataFrame({" + lines[1].strip().rstrip(",") + "}" + idx + ")"]
        return lines

    # ------------------------------------------------------------- GROUP BY
    def _group(self, st: A.Select, ctx: Ctx, prefer: Optional[dict] = None) -> list:
        prefer = prefer or {}
        L = []
        keys, pre = [], []
        for i, gexpr in enumerate(st.group_by):
            v = self.ex(gexpr, ctx)
            if isinstance(gexpr, A.Col):
                col = ctx.resolve(gexpr.parts)
                if col is None:
                    raise NaoModernizavel(f"GROUP BY por parâmetro '{expr_sql(gexpr)}'")
                keys.append((key(gexpr), col))
            else:
                name = f"_grupo{i}"
                pre.append((name, v))
                keys.append((key(gexpr), name))
        aggs = []   # (nome, coluna de entrada, função)
        seen = {}
        exprs = [e for e, _ in st.columns] + [st.having] + [e for e, _ in st.order_by]
        n_in = 0
        for e in exprs:
            for node in A.walk(e):
                if not (isinstance(node, A.Func) and node.name.lower() in A.AGGREGATES):
                    continue
                k = key(node)
                if k in seen:
                    continue
                fn = node.name.lower()
                if node.distinct:
                    raise NaoModernizavel("agregação com DISTINCT")
                pfn = {"sum": "sum", "avg": "mean", "count": "count", "min": "min", "max": "max",
                       "first": "first", "last": "last", "stdev": "std", "var": "var"}.get(fn)
                if pfn is None:
                    raise NaoModernizavel(f"agregação {node.name}")
                arg = node.args[0] if node.args else A.Star()
                if isinstance(arg, A.Star):
                    src, pfn = (keys[0][1] if keys else None), "size"
                    if src is None:
                        src = next(iter(ctx.frame_cols()))
                elif isinstance(arg, A.Col) and ctx.resolve(arg.parts):
                    src = ctx.resolve(arg.parts)
                else:
                    src = f"_entrada{n_in}"
                    n_in += 1
                    pre.append((src, self.ex(arg, ctx)))
                name = prefer.get(k)
                ocupados = {c.lower() for _, c in keys} | {a[0].lower() for a in aggs}
                if not name or name.lower() in ocupados or name.startswith("_"):
                    name = f"_agg{len(aggs)}"
                aggs.append((name, src, pfn))
                seen[k] = name
        if pre:
            L.append("df = df.assign(**{")
            for n, v in pre:
                L.append(f"    {q(n)}: {_strip(v.code)},")
            L.append("})")
        if all(n.isidentifier() for n, _, _ in aggs):
            agg_kw = ", ".join(f"{n}=({q(s)}, {q(f)})" for n, s, f in aggs)
        else:
            agg_kw = "**{" + ", ".join(f"{q(n)}: ({q(s)}, {q(f)})" for n, s, f in aggs) + "}"
        if keys:
            kcols = [c for _, c in keys]
            by = q(kcols[0]) if len(kcols) == 1 else q(kcols)
            linha = f"g = df.groupby({by}, dropna=False, as_index=False).agg({agg_kw})"
            if not aggs:
                L.append(f"g = df[{q(kcols)}].drop_duplicates()")
            elif len(linha) <= 100:
                L.append(linha)
            else:
                L.append(f"g = df.groupby({by}, dropna=False, as_index=False).agg(")
                if agg_kw.startswith("**{"):
                    L.append("    **{")
                    L += [f"        {q(n)}: ({q(s_)}, {q(f)})," for n, s_, f in aggs]
                    L.append("    }")
                else:
                    L += [f"    {n}=({q(s_)}, {q(f)})," for n, s_, f in aggs]
                L.append(")")
        else:
            if any(f in ("first", "last") for _, _, f in aggs):
                raise NaoModernizavel("First/Last sem GROUP BY")
            row = ", ".join(f"{q(n)}: [len(df)]" if f == "size" else f"{q(n)}: [df[{q(s)}].{f}()]"
                            for n, s, f in aggs)
            L.append(f"g = pd.DataFrame({{{row}}})")
        # no frame agrupado, as expressões passam a apontar para as colunas de g
        kinds = {}
        for k, c in keys:
            kinds[c] = ctx.kinds.get(c, "any")
        for n, s, f in aggs:
            kinds[n] = "num" if f in ("sum", "mean", "count", "size", "std", "var") else ctx.kinds.get(s, "any")
        group_cols = {k: c for k, c in keys}
        ctx.agg = {**{k: c for k, c in seen.items()}, **group_cols}
        ctx.group_cols = set(group_cols.values())
        ctx.pre_colmap = dict(ctx.colmap)
        ctx.colmap = {kk: v for kk, v in ctx.colmap.items() if v in group_cols.values()}
        ctx.kinds = kinds
        ctx.frame = "g"
        ctx.grouped = True
        if st.having is not None:
            v = self.ex(st.having, ctx)
            L.append(f"g = g[{_strip(v.code)}]")
        return L

    # ------------------------------------------------------------- UNION
    def _union(self, u: A.Union, ctx: Ctx, info: Info) -> list:
        L, partes = [], []
        cols = kinds = None
        for i, part in enumerate(u.parts):
            c2 = Ctx(self)
            body, pc, pk = self._select(part, c2, info)
            for p in c2.params:
                if p not in ctx.params:
                    ctx.params.append(p)
            L.append(f"# parte {i + 1}")
            L += body
            L.append(f"parte{i + 1} = resultado")
            partes.append(f"parte{i + 1}")
            if cols is None:
                cols, kinds = pc, pk
            elif len(pc) != len(cols):
                raise NaoModernizavel("UNION com número diferente de colunas")
        info.cols, info.kinds = cols, kinds
        renomeadas = [partes[0]] + [f"{p}.set_axis({q(cols)}, axis=1)" for p in partes[1:]]
        L.append(f"resultado = pd.concat([{', '.join(renomeadas)}], ignore_index=True)")
        if not all(u.all_flags):
            if any(u.all_flags):
                raise NaoModernizavel("mistura de UNION e UNION ALL")
            L.append("resultado = resultado.drop_duplicates(ignore_index=True)")
        if u.order_by:
            by, asc = [], []
            for e, d in u.order_by:
                if not (isinstance(e, A.Col) and len(e.parts) == 1):
                    raise NaoModernizavel("ORDER BY de UNION por expressão")
                n = next((c for c in cols if c.lower() == e.parts[0].lower()), None)
                if n is None:
                    raise NaoModernizavel("ORDER BY de UNION por coluna fora da saída")
                by.append(n)
                asc.append(not d)
            by_, asc_ = (q(by[0]), asc[0]) if len(by) == 1 else (q(by), q(asc))
            L.append(f"resultado = resultado.sort_values({by_}, ascending={asc_}, "
                     f'na_position={q("first" if all(asc) else "last")}, kind="stable", ignore_index=True)')
        L.append("return resultado")
        info.params = ctx.params
        return L

    # ------------------------------------------------------------- ações
    def _tabela_unica(self, src) -> A.TableRef:
        if not isinstance(src, A.TableRef):
            raise NaoModernizavel("atualização/exclusão com JOIN")
        return src

    def _update(self, st: A.Update, info: Info) -> list:
        ref = self._tabela_unica(st.source)
        ctx = Ctx(self)
        info.alvo = self._nome_real(ref.name)
        L = self._fontes(ref, ctx, info, [])
        L.append("")
        if st.where is not None:
            cond = self.ex(st.where, ctx)
            L.append(f"filtro = {_strip(cond.code)}")
        novos = []
        for col, val in st.sets:
            c = ctx.resolve(col.parts)
            if c is None:
                raise NaoModernizavel(f"coluna '{expr_sql(col)}' não existe em {ref.name}")
            novos.append((c, self.ex(val, ctx)))
        if len(novos) > 1:  # o Access calcula todos os SET com os valores antigos
            L.append("novos = {")
            for c, v in novos:
                L.append(f"    {q(c)}: {_strip(v.code)},")
            L.append("}")
            for c, v in novos:
                rhs = f"novos[{q(c)}]" + ("[filtro]" if st.where is not None and v.series else "")
                L.append(f"df.loc[{'filtro' if st.where is not None else ':'}, {q(c)}] = {rhs}")
        else:
            c, v = novos[0]
            rhs = _strip(v.code)
            if st.where is not None and v.series:
                rhs = f"({rhs})[filtro]" if not rhs.startswith("df[") or " " in rhs else f"{rhs}[filtro]"
            L.append(f"df.loc[{'filtro' if st.where is not None else ':'}, {q(c)}] = {rhs}")
        L.append("return df")
        info.params = ctx.params
        return L

    def _delete(self, st: A.Delete, info: Info) -> list:
        ref = self._tabela_unica(st.source)
        ctx = Ctx(self)
        info.alvo = self._nome_real(ref.name)
        L = self._fontes(ref, ctx, info, [])
        if st.where is None:
            L.append("return df.iloc[0:0]")
        else:
            cond = self.ex(st.where, ctx)
            L.append(f"filtro = {_strip(cond.code)}")
            L.append("return df[~filtro].reset_index(drop=True)")
        info.params = ctx.params
        return L

    def _insert(self, st: A.Insert, info: Info) -> list:
        alvo = self._nome_real(st.table)
        info.alvo = alvo
        r = self.fonte_info(alvo)
        L = []
        if st.values is not None:
            ctx = Ctx(self)
            cols = st.columns or (r[0][: len(st.values)] if r else None)
            if not cols:
                raise NaoModernizavel("INSERT VALUES sem lista de colunas")
            vals = [self.ex(v, ctx) for v in st.values]
            if any(v.series for v in vals):
                raise NaoModernizavel("INSERT VALUES com coluna")
            if alvo not in info.fontes:
                info.fontes.insert(0, alvo)
            row = ", ".join(f"{q(c)}: [{v.code}]" for c, v in zip(cols, vals))
            L.append(f"novas = pd.DataFrame({{{row}}})")
            info.params = ctx.params
        else:
            if not isinstance(st.select, A.Select):
                raise NaoModernizavel("INSERT a partir de UNION")
            ctx = Ctx(self)
            body, cols, _ = self._select(st.select, ctx, info)
            L += body
            if st.columns:
                ren = dict(zip(cols, st.columns))
                ren = {a: b for a, b in ren.items() if a != b}
                L.append(f"novas = resultado.rename(columns={q(ren)})" if ren else "novas = resultado")
            else:
                L.append("novas = resultado")
            if alvo not in info.fontes:
                info.fontes.insert(0, alvo)
            info.params = ctx.params
        L.append(f"return pd.concat([{ident(alvo)}, novas], ignore_index=True)")
        return L

    # ------------------------------------------------------------- expressões
    def ex(self, e, ctx: Ctx) -> V:
        if ctx.grouped:
            k = key(e)
            if k in ctx.agg:
                c = ctx.agg[k]
                return V(f'g["{c}"]', True, ctx.kinds.get(c, "any"))
            if isinstance(e, A.Col):
                for gk, c in ctx.agg.items():
                    if c in group_cols_set(ctx) and gk.split(".")[-1].strip("[]") == e.parts[-1].lower():
                        return V(f'g["{c}"]', True, ctx.kinds.get(c, "any"))
        if isinstance(e, A.Col):
            if len(e.parts) >= 3 or (len(e.parts) == 2 and e.parts[0].lower() in ("forms", "reports")):
                return ctx.param(e.parts)
            col = ctx.resolve(e.parts)
            if col is None:
                if ctx.grouped:
                    antes = Ctx(self)
                    antes.colmap = ctx.pre_colmap
                    if antes.resolve(e.parts) is not None:
                        raise NaoModernizavel(f"coluna '{expr_sql(e)}' fora do GROUP BY")
                if len(e.parts) == 2:
                    raise NaoModernizavel(f"coluna '{expr_sql(e)}' não encontrada")
                return ctx.param(e.parts)
            return V(f'{ctx.frame}["{col}"]', True, ctx.kinds.get(col, "any"))
        if isinstance(e, A.Lit):
            if e.kind == "null":
                return V("None", False, "any")
            kind = {"num": "num", "str": "str", "date": "date", "bool": "bool"}[e.kind]
            return V(_lit_py(e.value), False, kind)
        if isinstance(e, A.Unary):
            v = self.ex(e.operand, ctx)
            if e.op == "NOT":
                return V(f"~{_wrap(v.code)}" if v.series else f"(not {v.code})", v.series, "bool")
            return V(f"-{_wrap(v.code)}", v.series, v.kind)
        if isinstance(e, A.Bin):
            return self._bin(e, ctx)
        if isinstance(e, A.Between):
            x, lo, hi = self.ex(e.expr, ctx), self.ex(e.low, ctx), self.ex(e.high, ctx)
            if not x.series:
                raise NaoModernizavel("BETWEEN sobre valor que não é coluna")
            code = f"{_wrap(x.code)}.between({lo.code}, {hi.code})"
            return V(f"~{code}" if e.negated else code, True, "bool")
        if isinstance(e, A.InList):
            x = self.ex(e.expr, ctx)
            items = [self.ex(i, ctx) for i in e.items]
            if not x.series or any(i.series for i in items):
                raise NaoModernizavel("IN com colunas na lista")
            code = f"{_wrap(x.code)}.isin([{', '.join(i.code for i in items)}])"
            return V(f"~{code}" if e.negated else code, True, "bool")
        if isinstance(e, A.IsNull):
            x = self.ex(e.expr, ctx)
            if not x.series:
                return V(f"{'not ' if e.negated else ''}pd.isna({x.code})", False, "bool")
            return V(f"{_wrap(x.code)}.{'notna' if e.negated else 'isna'}()", True, "bool")
        if isinstance(e, A.Like):
            return self._like(e, ctx)
        if isinstance(e, A.Func):
            return self._func(e, ctx)
        raise NaoModernizavel(f"expressão {type(e).__name__} (subconsulta/EXISTS)")

    def _bin(self, e: A.Bin, ctx: Ctx) -> V:
        a, b = self.ex(e.left, ctx), self.ex(e.right, ctx)
        s = a.series or b.series
        op = e.op
        if op in ("+", "-", "*", "/"):
            if op == "+" and (a.kind == "str" or b.kind == "str"):
                return V(f"({a.code} + {b.code})", s, "str")
            kind = "date" if (op in "+-" and "date" in (a.kind, b.kind) and op == "+") else "num"
            return V(f"({a.code} {op} {b.code})", s, kind)
        if op == "^":
            return V(f"({a.code} ** {b.code})", s, "num")
        if op == "\\":
            raise NaoModernizavel("divisão inteira (\\) do Access")
        if op == "MOD":
            self.usa_np = True
            return V(f"np.fmod({_strip(a.code)}, {_strip(b.code)})", s, "num")
        if op == "&":
            folhas = []

            def achatar(x):
                if isinstance(x, A.Bin) and x.op == "&":
                    achatar(x.left)
                    achatar(x.right)
                else:
                    folhas.append(self._texto(self.ex(x, ctx)))
            achatar(e)
            return V("(" + " + ".join(folhas) + ")", s, "str")
        if op in ("=", "<>", "<", ">", "<=", ">="):
            if (isinstance(e.left, A.Lit) and e.left.kind == "null") or (isinstance(e.right, A.Lit) and e.right.kind == "null"):
                raise NaoModernizavel("comparação com Null (= Null)")
            pyop = {"=": "==", "<>": "!="}.get(op, op)
            code = f"({a.code} {pyop} {b.code})"
            if op == "<>" and s:
                # no Access, Null <> x é Null (a linha não entra); no pandas daria True
                guards = [f"{_wrap(x.code)}.notna()" for x in (a, b) if x.series]
                code = f"(({a.code} != {b.code}) & {' & '.join(guards)})"
            return V(code, s, "bool")
        if op in ("AND", "OR"):
            sym = "&" if op == "AND" else "|"
            if not s:
                return V(f"({a.code} {'and' if op == 'AND' else 'or'} {b.code})", False, "bool")
            return V(f"({_wrap(a.code)} {sym} {_wrap(b.code)})", True, "bool")
        if op == "XOR":
            return V(f"({_wrap(a.code)} ^ {_wrap(b.code)})", s, "bool")
        raise NaoModernizavel(f"operador {op}")

    def _texto(self, v: V) -> str:
        if not v.series:
            if v.kind == "str":
                return v.code
            return f'("" if pd.isna({v.code}) else str({v.code}))'
        if v.kind == "str":
            return f'{_wrap(v.code)}.fillna("")'
        return f'{_wrap(v.code)}.astype("string").fillna("")'

    def _like(self, e: A.Like, ctx: Ctx) -> V:
        x = self.ex(e.expr, ctx)
        if not (isinstance(e.pattern, A.Lit) and e.pattern.kind == "str") or not x.series:
            raise NaoModernizavel("LIKE com padrão que não é texto fixo")
        pat = e.pattern.value
        base = f"{_wrap(x.code)}.str.lower()" if x.kind == "str" else f'{_wrap(x.code)}.astype("string").str.lower()'
        core = pat.strip("*")
        simple = not re.search(r"[?#\[\]]", pat) and "*" not in core
        low = core.lower()
        if simple and pat.startswith("*") and pat.endswith("*") and core:
            code = f"{base}.str.contains({q(low)}, regex=False, na=False)"
        elif simple and pat.endswith("*") and not pat.startswith("*"):
            code = f"{base}.str.startswith({q(low)}, na=False)"
        elif simple and pat.startswith("*") and not pat.endswith("*"):
            code = f"{base}.str.endswith({q(low)}, na=False)"
        elif not re.search(r"[*?#\[\]]", pat):
            code = f"({base} == {q(pat.lower())})"
        else:
            from .runtime.core import like_regex
            rx = like_regex(pat).pattern
            code = f"{_wrap(x.code)}.str.fullmatch({q(rx[:-2])}, case=False, na=False)"
        return V(f"~{code}" if e.negated else code, True, "bool")

    def _func(self, e: A.Func, ctx: Ctx) -> V:
        n = e.name.lower()
        if n in A.AGGREGATES:
            raise NaoModernizavel(f"agregação {e.name} fora de GROUP BY")
        if n.startswith("d") and n[1:] in ("lookup", "sum", "count", "avg", "min", "max", "first", "last"):
            raise NaoModernizavel(f"{e.name} (função de domínio)")
        args = [self.ex(a, ctx) for a in e.args]
        s = any(a.series for a in args)
        a0 = args[0] if args else None

        def ser(v: V, as_str=False):
            if as_str and v.kind != "str":
                return f'{_wrap(v.code)}.astype("string")'
            return _wrap(v.code)

        if n == "iif":
            c, x, y = args
            if not s:
                return V(f"({x.code} if {c.code} else {y.code})", False, x.kind)
            self.usa_np = True
            cond = c.code if c.series else f"np.full(len({ctx.frame}), bool({c.code}))"
            return V(f"pd.Series(np.where({_strip(cond)}, {x.code}, {y.code}), index={ctx.frame}.index)", True,
                     x.kind if x.kind == y.kind else "any")
        if n == "switch":
            if len(args) % 2:
                raise NaoModernizavel("Switch com número ímpar de argumentos")
            self.usa_np = True
            conds = ", ".join(_strip(args[i].code) for i in range(0, len(args), 2))
            vals = ", ".join(args[i].code for i in range(1, len(args), 2))
            return V(f"pd.Series(np.select([{conds}], [{vals}], default=None), index={ctx.frame}.index)", True)
        if n == "nz":
            d = args[1].code if len(args) > 1 else '""'
            if not a0.series:
                return V(f"({d} if pd.isna({a0.code}) else {a0.code})", False, a0.kind)
            return V(f"{ser(a0)}.fillna({d})", True, a0.kind if len(args) > 1 else "any")
        if n == "isnull":
            return V(f"{ser(a0)}.isna()" if a0.series else f"pd.isna({a0.code})", a0.series, "bool")
        if not s:
            raise NaoModernizavel(f"{e.name} só com valores fixos/parâmetros")
        if n in ("year", "month", "day", "hour", "minute", "second"):
            attr = {"year": "year", "month": "month", "day": "day", "hour": "hour", "minute": "minute",
                    "second": "second"}[n]
            return V(f"{ser(a0)}.dt.{attr}", True, "num")
        if n == "weekday" and len(args) == 1:
            return V(f"({ser(a0)}.dt.dayofweek + 1) % 7 + 1", True, "num")
        if n in ("ucase", "lcase"):
            return V(f"{ser(a0, True)}.str.{'upper' if n == 'ucase' else 'lower'}()", True, "str")
        if n in ("trim", "ltrim", "rtrim"):
            m = {"trim": "strip", "ltrim": "lstrip", "rtrim": "rstrip"}[n]
            return V(f'{ser(a0, True)}.str.{m}(" ")', True, "str")
        if n == "len":
            return V(f"{ser(a0, True)}.str.len()", True, "num")
        if n in ("left", "right") and not args[1].series:
            k = args[1].code
            sl = f"[:{k}]" if n == "left" else f"[-{k}:]"
            return V(f"{ser(a0, True)}.str{sl}", True, "str")
        if n == "mid" and not any(x.series for x in args[1:]):
            start = args[1].code
            if re.fullmatch(r"\d+", start):
                st0 = int(start) - 1
                if len(args) > 2 and re.fullmatch(r"\d+", args[2].code):
                    return V(f"{ser(a0, True)}.str[{st0}:{st0 + int(args[2].code)}]", True, "str")
                if len(args) == 2:
                    return V(f"{ser(a0, True)}.str[{st0}:]", True, "str")
            raise NaoModernizavel("Mid com posição não numérica")
        if n == "instr" and len(args) == 2 and not args[1].series:
            return V(f"{ser(a0, True)}.str.lower().str.find(str({args[1].code}).lower()) + 1", True, "num")
        if n == "replace" and len(args) == 3 and not args[1].series and not args[2].series:
            return V(f"{ser(a0, True)}.str.replace({args[1].code}, {args[2].code}, regex=False)", True, "str")
        if n == "round":
            nd = args[1].code if len(args) > 1 else "0"
            return V(f"{ser(a0)}.round({nd})", True, "num")
        if n == "abs":
            return V(f"{ser(a0)}.abs()", True, "num")
        if n in ("int", "fix"):
            self.usa_np = True
            return V(f"np.{'floor' if n == 'int' else 'trunc'}({_strip(a0.code)})", True, "num")
        if n in ("cdbl", "csng", "ccur", "cdec", "val"):
            return V(f'pd.to_numeric({_strip(a0.code)}, errors="coerce")', True, "num")
        if n in ("clng", "cint"):
            return V(f'pd.to_numeric({_strip(a0.code)}, errors="coerce").round().astype("Int64")', True, "num")
        if n == "cstr":
            return V(f'{ser(a0)}.astype("string")', True, "str")
        if n in ("cdate", "datevalue"):
            return V(f"pd.to_datetime({_strip(a0.code)}, dayfirst=True)", True, "date")
        if n == "format" and len(args) == 2 and not args[1].series and a0.kind == "date":
            fmt = _strftime(e.args[1].value if isinstance(e.args[1], A.Lit) else None)
            if fmt is None:
                raise NaoModernizavel(f"Format de data com padrão {expr_sql(e.args[1])}")
            return V(f"{ser(a0)}.dt.strftime({q(fmt)})", True, "str")
        if n == "dateadd" and isinstance(e.args[0], A.Lit) and not args[1].series:
            unit = e.args[0].value.lower()
            m = {"d": "days", "y": "days", "w": "days", "ww": "weeks", "m": "months", "yyyy": "years",
                 "q": None, "h": "hours", "n": "minutes", "s": "seconds"}.get(unit)
            if m is None:
                raise NaoModernizavel(f"DateAdd com intervalo '{unit}'")
            if m in ("months", "years"):
                return V(f"{ser(args[2])} + pd.DateOffset({m}=int({args[1].code}))", True, "date")
            return V(f"{ser(args[2])} + pd.Timedelta({m}={args[1].code})", True, "date")
        if n == "datediff" and isinstance(e.args[0], A.Lit) and len(args) == 3:
            unit = e.args[0].value.lower()
            a, b = args[1], args[2]
            if unit == "d":
                return V(f"({_wrap(b.code)} - {_wrap(a.code)}).dt.days", True, "num")
            if unit == "m":
                return V(f"({_wrap(b.code)}.dt.year - {_wrap(a.code)}.dt.year) * 12 + "
                         f"({_wrap(b.code)}.dt.month - {_wrap(a.code)}.dt.month)", True, "num")
            if unit == "yyyy":
                return V(f"{_wrap(b.code)}.dt.year - {_wrap(a.code)}.dt.year", True, "num")
            raise NaoModernizavel(f"DateDiff com intervalo '{unit}'")
        if n in self.udfs and len(args) >= 1:
            serie = [x for x in args if x.series]
            if len(serie) == 1:
                mod, fn = self.udfs[n].split(".")
                self.modulos_usados.add(mod)
                argl = ", ".join("v" if x.series else x.code for x in args)
                return V(f"{_wrap(serie[0].code)}.map(lambda v: {mod}.{fn}({argl}))", True)
        raise NaoModernizavel(f"função {e.name}")


# ======================================================================= utilidades
def group_cols_set(ctx) -> set:
    return ctx.group_cols


def _encadear(L: list) -> list:
    """'df = x' seguido de 'df = df.merge(...)' vira 'df = x.merge(...)'."""
    out = []
    for ln in L:
        if out and ln.startswith("df = df.merge(") and re.fullmatch(r"df = \w+", out[-1]):
            out[-1] = out[-1] + ln[len("df = df"):]
            continue
        out.append(ln)
    return out


def _split_and(e):
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


def _equi(c, left: set, right: set):
    if not (isinstance(c, A.Bin) and c.op == "=" and isinstance(c.left, A.Col) and isinstance(c.right, A.Col)):
        return None
    if len(c.left.parts) < 2 or len(c.right.parts) < 2:
        return None
    la, lc = c.left.parts[-2], c.left.parts[-1]
    ra, rc = c.right.parts[-2], c.right.parts[-1]
    if la.lower() in left and ra.lower() in right:
        return (la, lc), (ra, rc)
    if ra.lower() in left and la.lower() in right:
        return (ra, rc), (la, lc)
    return None


def _wrap(code: str) -> str:
    if re.fullmatch(r"[\w.]+(\[[^\[\]]*\])*(\.\w+\(\))?", code) or (code.startswith("(") and _balanced(code)):
        return code
    return f"({code})"


def _balanced(code: str) -> bool:
    depth, q = 0, None
    for i, ch in enumerate(code):
        if q:
            if ch == q and code[i - 1] != "\\":
                q = None
            continue
        if ch in "'\"":
            q = ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0 and i != len(code) - 1:
                return False
    return depth == 0


def _strip(code: str) -> str:
    if code.startswith("(") and code.endswith(")") and _balanced(code):
        return code[1:-1]
    return code


def _strftime(fmt: Optional[str]) -> Optional[str]:
    if fmt is None:
        return None
    low = fmt.lower()
    named = {"short date": "%d/%m/%Y", "general date": "%d/%m/%Y %H:%M:%S", "long time": "%H:%M:%S",
             "short time": "%H:%M"}
    if low in named:
        return named[low]
    toks = re.findall(r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|.", fmt, flags=re.I)
    out = []
    has_h = any(t.lower() in ("h", "hh") for t in toks)
    for i, t in enumerate(toks):
        tl = t.lower()
        if tl in ("mmmm", "mmm", "dddd", "ddd"):
            return None  # nomes de mês/dia dependem do idioma: fica na versão fiel
        if tl in ("mm", "m"):
            prev = next((x.lower() for x in reversed(toks[:i]) if x.strip(" :/-.")), "")
            if has_h and prev in ("h", "hh"):
                out.append("%M")
                continue
            out.append("%m")
            continue
        out.append({"yyyy": "%Y", "yy": "%y", "dd": "%d", "d": "%-d", "hh": "%H", "h": "%-H", "nn": "%M",
                    "n": "%M", "ss": "%S", "s": "%S"}.get(tl, t.replace("%", "%%")))
    res = "".join(out)
    return None if "%-" in res else res  # %-d não existe no Windows
