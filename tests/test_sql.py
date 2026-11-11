"""Testes do transpilador SQL Access -> pandas (executam o código gerado)."""
import numpy as np
import pandas as pd
import pytest

from acc2py.runtime import Database, MissingParameter
from acc2py.runtime import core as rt, sqlfuncs as F
from acc2py.sql.parser import parse
from acc2py.sql.to_pandas import SQLCompiler


def make_db():
    clientes = pd.DataFrame({"ID": [1, 2, 3, 4], "Nome": ["Ana", "bruno", "Carla", "Davi"],
                             "UF": ["PB", "PE", "pb", None], "Ativo": [True, True, False, True]})
    pedidos = pd.DataFrame({"PedidoID": [10, 11, 12, 13, 14], "ClienteID": [1, 1, 2, 3, None],
                            "Data": pd.to_datetime(["2026-01-05", "2026-02-10", "2026-02-11",
                                                    "2026-03-01", "2026-03-02"]),
                            "Valor": [100.0, 50.5, 200.0, None, 10.0]})
    return Database({"Clientes": clientes, "Pedidos": pedidos})


def run(sql, db=None, params=None, name="t"):
    db = db or make_db()
    code = SQLCompiler().compile_function(name, sql, "_q")
    ns = {"rt": rt, "F": F, "pd": pd, "np": np}
    exec(code, ns)
    db.queries[name] = ns["_q"]
    return ns["_q"](db, params), db


# ------------------------------------------------------------------ parser
@pytest.mark.parametrize("sql", [
    "SELECT DISTINCTROW c.* FROM Clientes AS c;",
    "SELECT [Nome] AS [Nome do Cliente], IIf(IsNull([UF]),'?',[UF]) FROM Clientes",
    "SELECT * FROM ((A INNER JOIN B ON A.id = B.a) INNER JOIN C ON B.id = C.b) LEFT JOIN D ON C.x = D.x",
    "SELECT * FROM T WHERE d >= #1/31/2026# AND Not x In (1,2) AND y Not Like 'a*'",
    "SELECT * FROM T WHERE Forms!Pedidos!txtCliente = ID",
    "SELECT x INTO Nova FROM T",
    "INSERT INTO T (a, b) SELECT x, y FROM U WHERE z > 1",
    "DELETE T.* FROM T WHERE a Is Null",
    "SELECT q.n FROM (SELECT Count(*) AS n FROM T) AS q",
    "SELECT * FROM T AS a WHERE EXISTS (SELECT * FROM U WHERE U.k = a.k)",
    "TRANSFORM Count(x) AS c SELECT a FROM T GROUP BY a PIVOT b IN ('x','y')",
    "SELECT 2^-1, 7 \\ 2, 7 Mod 3, -2^2 FROM T",
    "PARAMETERS [Data inicial] DateTime, Limite Long; SELECT * FROM T WHERE d > [Data inicial]",
])
def test_parse_variety(sql):
    parse(sql)


# ------------------------------------------------------------------ semântica
def test_where_case_insensitive_and_null():
    out, _ = run("SELECT Nome FROM Clientes WHERE UF = 'pb' ORDER BY Nome")
    assert list(out["Nome"]) == ["Ana", "Carla"]
    out, _ = run("SELECT Nome FROM Clientes WHERE UF <> 'PB'")
    assert list(out["Nome"]) == ["bruno"]  # Null <> 'PB' é Null, não verdadeiro


def test_left_join_group_count():
    out, _ = run("""SELECT Clientes.Nome, Count(Pedidos.PedidoID) AS N, Sum(Pedidos.Valor) AS Total
                    FROM Clientes LEFT JOIN Pedidos ON Clientes.ID = Pedidos.ClienteID
                    GROUP BY Clientes.Nome ORDER BY Clientes.Nome""")
    assert list(out["Nome"]) == ["Ana", "bruno", "Carla", "Davi"]
    assert list(out["N"]) == [2, 1, 1, 0]
    assert out["Total"].isna().tolist() == [False, False, True, True]


def test_join_null_keys_never_match():
    db = make_db()
    db.tables["X"] = pd.DataFrame({"k": [None, 1.0]})
    out, _ = run("SELECT Pedidos.PedidoID FROM Pedidos INNER JOIN X ON Pedidos.ClienteID = X.k", db)
    assert sorted(out["PedidoID"]) == [10, 11]


def test_top_includes_ties():
    out, _ = run("SELECT TOP 1 ClienteID, Count(*) AS N FROM Pedidos WHERE ClienteID Is Not Null "
                 "GROUP BY ClienteID ORDER BY Count(*) ASC")
    assert len(out) == 2  # ClienteID 2 e 3 empatam com 1 pedido


def test_order_nulls_first_asc():
    out, _ = run("SELECT Valor FROM Pedidos ORDER BY Valor")
    assert pd.isna(out["Valor"].iloc[0])


def test_parameters_and_missing():
    sql = "SELECT PedidoID FROM Pedidos WHERE Data Between [Inicio] And [Fim]"
    out, _ = run(sql, params={"inicio": pd.Timestamp("2026-02-01"), "[Fim]": pd.Timestamp("2026-02-28")})
    assert list(out["PedidoID"]) == [11, 12]
    with pytest.raises(MissingParameter):
        run(sql)


def test_like_wildcards():
    out, _ = run("SELECT Nome FROM Clientes WHERE Nome Like '[a-c]*' AND Nome Not Like '*a'")
    assert list(out["Nome"]) == ["bruno"]


def test_concat_and_functions():
    out, _ = run("SELECT Nome & ' - ' & UF AS R, Nz(UF,'?') AS U, Len(Nome) AS L, "
                 "Format(1234.5,'#,##0.00') AS Fm FROM Clientes WHERE ID = 4")
    row = out.iloc[0]
    assert row["R"] == "Davi - " and row["U"] == "?" and row["L"] == 4 and row["Fm"] == "1.234,50"


def test_union_and_distinct():
    out, _ = run("SELECT UF FROM Clientes WHERE UF Is Not Null UNION SELECT 'PB' FROM Clientes ORDER BY UF")
    assert list(out["UF"]) == ["PB", "PE"]  # 'PB' e 'pb' são iguais para o Access


def test_correlated_subquery_and_in():
    out, _ = run("SELECT c.Nome, (SELECT Count(*) FROM Pedidos AS p WHERE p.ClienteID = c.ID) AS N "
                 "FROM Clientes AS c WHERE c.ID NOT IN (SELECT ClienteID FROM Pedidos WHERE ClienteID Is Not Null)")
    assert list(out["Nome"]) == ["Davi"] and list(out["N"]) == [0]


def test_crosstab():
    out, _ = run("""TRANSFORM Sum(Valor) SELECT ClienteID FROM Pedidos WHERE ClienteID Is Not Null
                    GROUP BY ClienteID PIVOT Year([Data])""")
    assert list(out.columns) == ["ClienteID", "2026"]
    assert list(out["2026"].fillna(-1)) == [150.5, 200.0, -1]


def test_saved_query_as_source_and_dlookup():
    db = make_db()
    run("SELECT ClienteID, Sum(Valor) AS Total FROM Pedidos GROUP BY ClienteID", db, name="Totais")
    out, _ = run("SELECT c.Nome, t.Total, DLookup('Total','Totais','ClienteID=' & c.ID) AS T2 "
                 "FROM Clientes AS c INNER JOIN Totais AS t ON c.ID = t.ClienteID ORDER BY c.ID", db)
    assert out["Total"].fillna(-1).tolist() == out["T2"].fillna(-1).tolist() == [150.5, 200.0, -1]


def test_action_queries():
    db = make_db()
    n, db = run("UPDATE Clientes SET Ativo = False WHERE UF = 'PE'", db)
    assert n == 1 and not db.tables["Clientes"].loc[1, "Ativo"]
    n, db = run("DELETE * FROM Pedidos WHERE Valor Is Null", db)
    assert n == 1 and len(db.tables["Pedidos"]) == 4
    n, db = run("INSERT INTO Clientes (ID, Nome) VALUES (9, 'Zé')", db)
    assert n == 1 and db.tables["Clientes"]["Nome"].iloc[-1] == "Zé"
    n, db = run("SELECT Nome INTO Copia FROM Clientes WHERE Ativo = True", db)
    assert list(db.tables["Copia"]["Nome"]) == ["Ana", "Davi"]


def test_implicit_join_uses_keys():
    out, _ = run("SELECT c.Nome, p.Valor FROM Clientes c, Pedidos p WHERE c.ID = p.ClienteID AND p.Valor > 60 "
                 "ORDER BY p.Valor")
    assert list(out["Nome"]) == ["Ana", "bruno"]


def test_runtime_sql():
    db = make_db()
    out = db.sql("SELECT Count(*) AS n FROM Clientes WHERE Ativo")
    assert out["n"].iloc[0] == 3
