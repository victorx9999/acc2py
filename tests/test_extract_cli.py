"""Testes de extração (MS-OVBA, MSysQueries, SaveAsText) e da linha de comando."""
import subprocess
import sys
from pathlib import Path

from acc2py.extract.folder import sql_from_saveastext
from acc2py.extract.parser_backend import build_sql, ovba_decompress
from acc2py.sql.parser import parse

ROOT = Path(__file__).resolve().parents[1]


def test_ovba_spec_example():
    comp = bytes.fromhex("012FB00023616161626364658266007061676869" "6A013808616B6C00306D6E6F700671027004107273747576107778797A003C")
    assert ovba_decompress(comp) == b"#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa"


def test_msysqueries_rebuild():
    rows = [{"Attribute": 1, "Flag": 1}, {"Attribute": 3, "Flag": 0x10, "Name1": "3"},
            {"Attribute": 5, "Name1": "Clientes", "Name2": "c"}, {"Attribute": 5, "Name1": "Pedidos"},
            {"Attribute": 6, "Expression": "c.Nome"},
            {"Attribute": 7, "Name1": "c", "Name2": "Pedidos", "Expression": "c.ID = Pedidos.ClienteID", "Flag": 2},
            {"Attribute": 11, "Expression": "c.Nome", "Name1": "D"}]
    sql = build_sql(rows)
    assert "TOP 3" in sql and "LEFT JOIN" in sql and sql.rstrip(";").endswith("DESC")
    parse(sql)


def test_saveastext_query():
    txt = 'dbMemo "SQL" ="SELECT \\"x\\" AS a\\015\\012FROM T;\\015\\012"\n'
    assert sql_from_saveastext(txt) == 'SELECT "x" AS a\nFROM T;'


def _make_folder(tmp: Path):
    (tmp / "Clientes.csv").write_text("ID,Nome\n1,Ana\n2,Bia\n", encoding="utf-8")
    (tmp / "Lista.sql").write_text("SELECT Nome, Dobro(ID) AS D FROM Clientes ORDER BY Nome DESC", encoding="utf-8")
    (tmp / "Base.sql").write_text("SELECT * FROM Clientes WHERE ID > 1", encoding="utf-8")
    (tmp / "Usa Base.sql").write_text("SELECT Count(*) AS n FROM Base", encoding="utf-8")
    (tmp / "Util.bas").write_text('Attribute VB_Name = "Util"\nPublic Function Dobro(x)\n  Dobro = x * 2\nEnd Function\n',
                                  encoding="utf-8")
    (tmp / "Outro.bas").write_text('Attribute VB_Name = "Outro"\nSub Nada()\nEnd Sub\n', encoding="utf-8")


def test_cli_selected_queries(tmp_path):
    src = tmp_path / "banco"
    src.mkdir()
    _make_folder(src)
    out = tmp_path / "saida"
    r = subprocess.run([sys.executable, "-m", "acc2py", str(src), "-o", str(out), "--pacote", "meu",
                        "--consultas", "usa base", "lista"], cwd=ROOT, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    pkg = out / "meu"
    rel = (pkg / "RELATORIO.md").read_text(encoding="utf-8")
    assert "Base" in rel and "dependência" in rel
    assert (pkg / "modulos" / "Util.py").exists()          # função usada pela consulta
    assert not (pkg / "modulos" / "Outro.py").exists()     # módulo não pedido
    code = ("from meu import abrir\ndb = abrir()\n"
            "print(db.run('Usa Base')['n'][0], list(db.run('Lista')['D']))")
    r = subprocess.run([sys.executable, "-c", code], cwd=out, capture_output=True, text=True,
                       env={"PYTHONPATH": str(ROOT)})
    assert r.returncode == 0, r.stderr
    assert r.stdout.strip() == "1 [4, 2]"


def test_cli_unknown_query(tmp_path):
    src = tmp_path / "banco"
    src.mkdir()
    _make_folder(src)
    r = subprocess.run([sys.executable, "-m", "acc2py", str(src), "-o", str(tmp_path / "o"),
                        "--consultas", "Listaa"], cwd=ROOT, capture_output=True, text=True)
    assert r.returncode != 0 and "Lista" in r.stderr


def test_cli_list_file_wildcard_and_single_output(tmp_path):
    import pandas as pd
    src = tmp_path / "banco"
    src.mkdir()
    _make_folder(src)
    (src / "Apaga.sql").write_text("DELETE * FROM Clientes", encoding="utf-8")
    lista = tmp_path / "lista.txt"
    lista.write_text("# relatório\nLista\nUsa*\nApaga\n", encoding="utf-8")
    out = tmp_path / "saida"
    res = out / "tudo.xlsx"
    r = subprocess.run([sys.executable, "-m", "acc2py", str(src), "-o", str(out), "--pacote", "rel",
                        "-f", str(lista), "-r", str(res)], cwd=ROOT, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    sheets = pd.read_excel(res, sheet_name=None)
    assert list(sheets) == ["Resumo", "Lista", "Usa Base"]        # ação não é executada
    assert sheets["Usa Base"]["n"][0] == 1
    assert len(sheets["Lista"]) == 2


def test_sequence_runs_in_order_and_stops_on_error(tmp_path):
    import pandas as pd
    sys.path.insert(0, str(ROOT))
    from acc2py.project import transpile_database
    src = tmp_path / "banco"
    src.mkdir()
    (src / "T.csv").write_text("ID,V\n1,10\n2,20\n3,30\n", encoding="utf-8")
    (src / "Passo 1.sql").write_text("SELECT ID, V INTO Tmp FROM T WHERE V > [Min]", encoding="utf-8")
    (src / "Passo 2.sql").write_text("UPDATE Tmp SET V = V * 2", encoding="utf-8")
    (src / "Passo 10.sql").write_text("SELECT Sum(V) AS S FROM Tmp", encoding="utf-8")
    (src / "Quebra.sql").write_text("SELECT X FROM NaoExiste", encoding="utf-8")
    res = transpile_database(str(src), str(tmp_path / "o"), package="seq", sequence=["Passo*"])
    assert res.sequence == ["Passo 1", "Passo 2", "Passo 10"]  # ordem natural
    pkg = __import__("importlib").import_module
    sys.path.insert(0, str(tmp_path / "o"))
    seq = pkg("seq")
    log = seq.executar_sequencia(str(tmp_path / "f.xlsx"), {"Min": 15})
    assert list(log["Situação"]) == ["ok"] * 3
    final = pd.read_excel(tmp_path / "f.xlsx", sheet_name=None)
    assert list(final) == ["Final - Passo 10", "Etapas"]
    assert final["Final - Passo 10"]["S"][0] == 100            # (20+30)*2
    log = seq.executar_sequencia(None, {"Min": 15}, etapas=["Passo 1", "Quebra", "Passo 10"],
                                 db=seq.abrir(str(src)))
    assert list(log["Situação"]) == ["ok", "ERRO", "não executada"]
