"""Validação Access × Python, com um executor que simula o Access."""
import sys
from pathlib import Path

import pandas as pd

from acc2py.project import transpile_database
from acc2py.validacao import comparar, validar


def test_comparar_ignora_ordem_e_tipos_numericos():
    a = pd.DataFrame({"Nome": ["b", "a"], "Total": [2, 1.0000001], "D": pd.to_datetime(["2026-01-02", "2026-01-01"])})
    p = pd.DataFrame({"nome": ["a", "b"], "TOTAL": [1.0, 2.0], "d": pd.to_datetime(["2026-01-01", "2026-01-02"])})
    c = comparar(a, p)
    assert c["iguais"] and not c["mesma_ordem"]
    p.loc[0, "TOTAL"] = 1.5
    c = comparar(a, p)
    assert not c["iguais"] and c["linhas_so_access"] == 1 and len(c["diferencas"]) == 2


class FakeAccess:
    """Simula o Access executando o próprio pacote, com um 'bug' opcional numa etapa."""
    def __init__(self, src, bug=None):
        self.tables = {"T": pd.DataFrame({"ID": [1, 2, 3], "V": [10.0, 20.0, 30.0]})}
        self.bug = bug
        self.db = None

    def vinculadas(self):
        return set()

    def ler_tabela(self, nome):
        if self.db is not None and nome in self.db.tables:
            return self.db.tables[nome].copy()
        return self.tables[nome].copy()

    def executar(self, consulta, params, cria=None):
        if self.db is None:
            from acc2py.runtime import Database
            mod = sys.modules["val"]
            self.db = Database({k: v.copy() for k, v in self.tables.items()}, mod.CONSULTAS)
        r = self.db.run(consulta, params)
        if self.bug == consulta and isinstance(r, pd.DataFrame):
            r = r.copy()
            r.iloc[0, 0] = -999
        return r

    def fechar(self):
        pass


def _pkg(tmp_path):
    src = tmp_path / "banco"
    src.mkdir()
    (src / "T.csv").write_text("ID,V\n1,10\n2,20\n3,30\n", encoding="utf-8")
    (src / "1 Base.sql").write_text("SELECT ID, V INTO Tmp FROM T WHERE V > [Min]", encoding="utf-8")
    (src / "2 Dobra.sql").write_text("UPDATE Tmp SET V = V * 2", encoding="utf-8")
    (src / "3 Final.sql").write_text("SELECT Sum(V) AS S, Count(*) AS N FROM Tmp", encoding="utf-8")
    transpile_database(str(src), str(tmp_path / "o"), package="val", sequence=["1*", "2*", "3*"])
    sys.path.insert(0, str(tmp_path / "o"))
    for k in [k for k in sys.modules if k == "val" or k.startswith("val.")]:
        del sys.modules[k]
    return tmp_path / "o" / "val"


def test_validar_igual(tmp_path):
    pkg = _pkg(tmp_path)
    r = validar("x.accdb", pkg, {"Min": 15}, tmp_path / "v.xlsx", executor=FakeAccess(None))
    assert list(r["Resultado"]) == ["IGUAL"] * 4  # 3 etapas + tabela final Tmp
    assert "Resumo" in pd.read_excel(tmp_path / "v.xlsx", sheet_name=None)


def test_validar_aponta_diferenca(tmp_path):
    pkg = _pkg(tmp_path)
    r = validar("x.accdb", pkg, {"Min": 15}, tmp_path / "v.xlsx", executor=FakeAccess(None, bug="3 Final"))
    assert r.loc[r["Consulta"] == "3 Final", "Resultado"].item() == "DIFERENTE"
    abas = pd.read_excel(tmp_path / "v.xlsx", sheet_name=None)
    assert "Dif 3 Final" in abas


def test_entradas_excel_csv_tipos(tmp_path):
    from acc2py.runtime.inputs import read_input
    f = tmp_path / "m.csv"
    f.write_bytes("Agência;Valor;Data\n0012;1.234,50;31/01/2026\n".encode("cp1252"))
    df, chk = read_input("M", {"arquivo": str(f), "colunas": {"Agencia": "texto", "Valor": "numero",
                                                              "Data": "data", "Conta": "texto"}})
    assert df["Agencia"][0] == "0012" and df["Valor"][0] == 1234.5
    assert df["Data"][0] == pd.Timestamp("2026-01-31") and chk.faltando == ["Conta"] and not chk.ok


def test_pacote_mensal_sem_access(tmp_path):
    import importlib
    import pytest
    from acc2py.runtime.pacote import EntradaError
    pkg = _pkg(tmp_path)
    ent = __import__("json").loads((pkg / "entradas.json").read_text(encoding="utf-8"))
    ent["T"]["modo"] = "mensal"
    (pkg / "entradas.json").write_text(__import__("json").dumps(ent), encoding="utf-8")
    mod = importlib.import_module("val")
    with pytest.raises(EntradaError):
        mod.executar({"Min": "15"})
    arq = tmp_path / "t_jan.xlsx"
    pd.DataFrame({"id": [1, 2], "v": ["100", "200,5"]}).to_excel(arq, index=False)
    chk = mod.definir_entrada("T", str(arq))
    assert chk.ok
    etapas, saida = mod.executar({"Min": "15"})
    assert list(etapas["Situação"]) == ["ok"] * 3
    assert pd.read_excel(saida, sheet_name=0)["S"][0] == (100 + 200.5) * 2


def test_tela_exige_token():
    import json
    import urllib.error
    import urllib.request
    from acc2py.tela import servir
    srv, url, _ = servir(abrir_navegador=False, bloquear=False)
    base, token = url.split("/?t=")
    try:
        with urllib.request.urlopen(f"{base}/api/estado?t={token}") as r:
            assert json.load(r)["modo"] == "migracao"
        try:
            urllib.request.urlopen(f"{base}/api/estado")
            assert False
        except urllib.error.HTTPError as e:
            assert e.code == 403
    finally:
        srv.shutdown()
