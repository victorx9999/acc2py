"""Validação da migração: roda a mesma sequência no Access e no Python e compara.

Precisa ser feita ANTES do desligamento do Access. O fluxo:

1. copia o .accdb para uma pasta temporária (o original nunca é alterado);
2. na cópia, lê as tabelas de entrada da sequência (antes de qualquer etapa);
3. executa as etapas no Access, na ordem, e guarda o resultado de cada consulta de seleção;
4. executa o pacote Python com AS MESMAS tabelas de entrada e os mesmos parâmetros;
5. compara cada resultado (linhas, colunas e valores) e grava um relatório em Excel.

Por segurança, a validação se recusa a rodar se alguma etapa de ação
(atualizar/excluir/acrescentar/criar tabela) escrever numa tabela VINCULADA,
porque a cópia do .accdb continua apontando para a fonte real.
"""
from __future__ import annotations

import importlib
import math
import shutil
import sys
import tempfile
from collections import Counter
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from .runtime.core import CIDict, Database, MissingParameter


# ======================================================================= comparação
NULO = "∅"


def _nv(v) -> str:
    if v is None or v is pd.NaT or v is pd.NA:
        return NULO
    if isinstance(v, (bool, np.bool_)):
        return "V" if v else "F"
    if isinstance(v, (int, float, np.integer, np.floating, Decimal)):
        f = float(v)
        if math.isnan(f):
            return NULO
        r = round(f, 6)
        return f"{r:.6f}".rstrip("0").rstrip(".") if r != 0 else "0"
    if hasattr(v, "strftime"):
        return pd.Timestamp(v).strftime("%Y-%m-%d %H:%M:%S")
    s = str(v)
    return s.rstrip() if s else s


def comparar(access: pd.DataFrame, python: pd.DataFrame, limite: int = 50) -> dict:
    """Compara dois resultados ignorando a ordem das linhas (e informa se a ordem também bate)."""
    ca = {str(c).lower(): c for c in access.columns}
    cp = {str(c).lower(): c for c in python.columns}
    so_a = [ca[k] for k in ca if k not in cp]
    so_p = [cp[k] for k in cp if k not in ca]
    comuns = [ca[k] for k in ca if k in cp]
    A = access[comuns].copy()
    P = python[[cp[str(c).lower()] for c in comuns]].copy()
    P.columns = comuns
    ra = [tuple(_nv(v) for v in row) for row in A.itertuples(index=False, name=None)]
    rp = [tuple(_nv(v) for v in row) for row in P.itertuples(index=False, name=None)]
    ca_, cp_ = Counter(ra), Counter(rp)
    falta = list((ca_ - cp_).elements())   # no Access e não no Python
    sobra = list((cp_ - ca_).elements())   # no Python e não no Access
    linhas = [("só no Access", *r) for r in falta[:limite]] + [("só no Python", *r) for r in sobra[:limite]]
    dif = pd.DataFrame(linhas, columns=["Onde", *[str(c) for c in comuns]])
    iguais = not so_a and not so_p and not falta and not sobra
    return {"iguais": iguais, "linhas_access": len(access), "linhas_python": len(python),
            "colunas_so_access": so_a, "colunas_so_python": so_p,
            "linhas_so_access": len(falta), "linhas_so_python": len(sobra),
            "mesma_ordem": iguais and ra == rp, "diferencas": dif}


# ======================================================================= Access (COM)
class AccessExecutor:
    """Executa consultas salvas numa CÓPIA do .accdb pelo Microsoft Access (Windows)."""

    def __init__(self, origem: str):
        from .extract.com import ComExtractor
        self.tmp = Path(tempfile.mkdtemp(prefix="acc2py_validacao_"))
        self.copia = self.tmp / Path(origem).name
        shutil.copy2(origem, self.copia)
        self.ex = ComExtractor(str(self.copia))
        self.db = self.ex.db

    def vinculadas(self) -> set:
        return {k.lower() for k in self.ex.linked_tables()}

    def ler_tabela(self, nome: str) -> pd.DataFrame:
        return self.ex.read_table(nome)

    def existe_tabela(self, nome: str) -> bool:
        return any(td.Name.lower() == nome.lower() for td in self.db.TableDefs)

    def executar(self, consulta: str, parametros: dict, cria_tabela: Optional[str] = None):
        """Executa a consulta salva. Seleção -> DataFrame; ação -> nº de linhas afetadas."""
        if cria_tabela and self.existe_tabela(cria_tabela):
            self.db.TableDefs.Delete(cria_tabela)  # como o Access faz ao confirmar "excluir tabela existente"
            self.db.TableDefs.Refresh()
        qd = self.db.QueryDefs(consulta)
        p = CIDict(parametros)
        for i in range(qd.Parameters.Count):
            prm = qd.Parameters(i)
            nome = prm.Name.strip("[]")
            if nome not in p:
                raise MissingParameter(nome)
            prm.Value = p[nome]
        if qd.Type in (0, 16, 128):  # seleção, referência cruzada, união
            rs = qd.OpenRecordset(4)  # dbOpenSnapshot
            cols = [rs.Fields(i).Name for i in range(rs.Fields.Count)]
            if rs.EOF:
                rs.Close()
                return pd.DataFrame(columns=cols)
            rs.MoveLast()
            n = rs.RecordCount
            rs.MoveFirst()
            data = rs.GetRows(n)
            rs.Close()
            from .extract.com import _convert
            return pd.DataFrame({c: [_convert(v) for v in vals] for c, vals in zip(cols, data)})
        qd.Execute(128)  # dbFailOnError
        return int(qd.RecordsAffected)

    def fechar(self):
        try:
            self.ex.close()
        finally:
            shutil.rmtree(self.tmp, ignore_errors=True)


# ======================================================================= validação
def _importar_pacote(pkg_dir) -> object:
    p = Path(pkg_dir).resolve()
    if str(p.parent) not in sys.path:
        sys.path.insert(0, str(p.parent))
    for k in [k for k in sys.modules if k == p.name or k.startswith(p.name + ".")]:
        del sys.modules[k]
    return importlib.import_module(p.name)


def _alvos_de_acao(sql: str) -> tuple[list[str], Optional[str]]:
    """Tabelas que a consulta altera e, se for criar-tabela, o nome criado."""
    from .sql import ast as SA
    from .sql.parser import parse
    st = parse(sql)
    if isinstance(st, SA.Select) and st.into:
        return [st.into], st.into
    if isinstance(st, SA.Insert):
        return [st.table], None
    if isinstance(st, (SA.Update, SA.Delete)):
        src = st.source

        def tabs(s):
            if isinstance(s, SA.TableRef):
                return [s.name]
            if isinstance(s, SA.Join):
                return tabs(s.left) + tabs(s.right)
            return []
        return tabs(src), None
    return [], None


def validar(origem: str, pacote_dir, parametros: Optional[dict] = None, destino=None,
            etapas: Optional[list[str]] = None, executor=None, log=print) -> pd.DataFrame:
    """Valida o pacote contra o Access. Devolve o resumo e grava ``destino`` (.xlsx)."""
    mod = _importar_pacote(pacote_dir)
    pk = mod.pacote()
    etapas = list(etapas or pk.sequencia)
    if not etapas:
        raise ValueError("o pacote não tem SEQUENCIA; informe as etapas a validar")
    parametros = dict(parametros or {})
    ex = executor or AccessExecutor(origem)
    resumo, difs = [], {}
    try:
        vinc = ex.vinculadas()
        criadas = {}
        for e in etapas:
            alvos, cria = _alvos_de_acao(pk.sql_original[e])
            perigo = [a for a in alvos if a.lower() in vinc]
            if perigo:
                raise RuntimeError(f"a etapa '{e}' altera a tabela vinculada {perigo[0]!r}; validar "
                                   f"alteraria a fonte real. Importe a tabela no Access (cópia local) "
                                   f"antes de validar.")
            if cria:
                criadas[e] = cria
        log("Lendo as tabelas de entrada no Access (antes das etapas)…")
        entradas = {}
        for t in pk.entradas:
            entradas[t] = ex.ler_tabela(t)
            log(f"  {t}: {len(entradas[t])} linha(s)")
        db = Database({k: v.copy() for k, v in entradas.items()}, pk.consultas, pk.funcoes,
                      autonumber=pk.autonumeracao)
        from .runtime import vba
        vba.set_database(db)
        for i, e in enumerate(etapas, 1):
            log(f"[{i}/{len(etapas)}] {e}")
            linha = {"Etapa": i, "Consulta": e}
            try:
                ra = ex.executar(e, parametros, criadas.get(e))
            except Exception as err:  # noqa: BLE001
                linha.update({"Resultado": "erro no Access", "Detalhe": str(err)})
                resumo.append(linha)
                log(f"   Access: ERRO {err}")
                break
            try:
                rp = db.run(e, parametros)
            except Exception as err:  # noqa: BLE001
                linha.update({"Resultado": "erro no Python", "Detalhe": str(err)})
                resumo.append(linha)
                log(f"   Python: ERRO {err}")
                break
            if isinstance(ra, pd.DataFrame):
                c = comparar(ra, rp if isinstance(rp, pd.DataFrame) else pd.DataFrame())
                linha.update({"Tipo": "seleção", "Linhas Access": c["linhas_access"],
                              "Linhas Python": c["linhas_python"],
                              "Resultado": "IGUAL" if c["iguais"] else "DIFERENTE",
                              "Mesma ordem": "sim" if c["mesma_ordem"] else "não",
                              "Detalhe": "; ".join(filter(None, [
                                  f"colunas só no Access: {', '.join(map(str, c['colunas_so_access']))}" if c["colunas_so_access"] else "",
                                  f"colunas só no Python: {', '.join(map(str, c['colunas_so_python']))}" if c["colunas_so_python"] else "",
                                  f"{c['linhas_so_access']} linha(s) só no Access" if c["linhas_so_access"] else "",
                                  f"{c['linhas_so_python']} linha(s) só no Python" if c["linhas_so_python"] else ""]))})
                if not c["iguais"]:
                    difs[e] = c["diferencas"]
                log(f"   {linha['Resultado']} ({c['linhas_access']} × {c['linhas_python']} linhas)")
            else:
                n_py = rp if isinstance(rp, int) else None
                ok = n_py == ra
                linha.update({"Tipo": "ação", "Linhas Access": ra, "Linhas Python": n_py,
                              "Resultado": "IGUAL" if ok else "DIFERENTE",
                              "Detalhe": "" if ok else "nº de linhas afetadas diferente"})
                log(f"   {linha['Resultado']} ({ra} × {n_py} linhas afetadas)")
            resumo.append(linha)
        # estado final das tabelas alteradas pelas etapas de ação
        alteradas = set()
        for e in etapas[:len(resumo)]:
            alteradas |= set(_alvos_de_acao(pk.sql_original[e])[0])
        for t in sorted(alteradas):
            try:
                ta = ex.ler_tabela(t)
            except Exception:  # noqa: BLE001
                continue
            tp = db.tables.get(t)
            if tp is None:
                continue
            c = comparar(ta, tp)
            resumo.append({"Etapa": "", "Consulta": f"(tabela final) {t}", "Tipo": "tabela",
                           "Linhas Access": c["linhas_access"], "Linhas Python": c["linhas_python"],
                           "Resultado": "IGUAL" if c["iguais"] else "DIFERENTE",
                           "Detalhe": "" if c["iguais"] else
                           f"{c['linhas_so_access']} só no Access, {c['linhas_so_python']} só no Python"})
            if not c["iguais"]:
                difs[f"tabela {t}"] = c["diferencas"]
    finally:
        if executor is None:
            ex.fechar()
    df = pd.DataFrame(resumo)
    if destino:
        from .runtime.core import _sheet_name
        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        used: set = set()
        with pd.ExcelWriter(destino, engine="openpyxl") as w:
            info = pd.DataFrame({"Item": ["Banco", "Pacote", "Data", "Parâmetros"],
                                 "Valor": [str(origem), str(pacote_dir), datetime.now().strftime("%d/%m/%Y %H:%M"),
                                           ", ".join(f"{k}={v}" for k, v in parametros.items())]})
            df.to_excel(w, sheet_name=_sheet_name("Resumo", used), index=False)
            info.to_excel(w, sheet_name=_sheet_name("Informações", used), index=False)
            for nome, d in difs.items():
                d.to_excel(w, sheet_name=_sheet_name("Dif " + nome, used), index=False)
    return df
