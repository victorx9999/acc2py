"""Versão moderna (pandas idiomático): deve dar o mesmo resultado que a fiel."""
import importlib
import shutil
import sys
from pathlib import Path

import pandas as pd

from acc2py.project import transpile_database

ROOT = Path(__file__).resolve().parents[1]
BANCO = ROOT / "exemplos" / "banco_exportado"

CONSULTAS = {
    "q01 Filtro": "SELECT c.Nome, c.UF, c.Limite * 1.1 AS NovoLimite FROM Clientes AS c "
                  "WHERE c.UF = 'PB' AND c.Segmento <> 'Atacado' ORDER BY c.Nome",
    "q02 Join": "SELECT Clientes.Nome, Pedidos.PedidoID, Pedidos.Valor FROM Clientes INNER JOIN Pedidos "
                "ON Clientes.ID = Pedidos.ClienteID WHERE Pedidos.Status IN ('Pago','Aberto') ORDER BY Pedidos.Valor DESC",
    "q03 Agrupa": "SELECT c.Segmento, Count(*) AS Qtde, Sum(p.Valor) AS Total, Avg(p.Valor) AS Media, "
                  "Max(p.Data) AS Ultimo FROM Clientes c LEFT JOIN Pedidos p ON c.ID = p.ClienteID "
                  "GROUP BY c.Segmento HAVING Count(*) > 1 ORDER BY c.Segmento",
    "q04 Funcoes": "SELECT PedidoID, Year([Data]) AS Ano, Format([Data],'dd/mm/yyyy') AS Dia, "
                   "IIf([Valor]>100,'Alto','Baixo') AS Faixa, Nz([Valor],0) AS V, DateAdd('m',1,[Data]) AS Venc, "
                   "UCase(Left([Status],3)) AS St FROM Pedidos",
    "q05 Like": "SELECT Nome FROM Clientes WHERE Nome Like '*a' OR Nome Like 'D*' ORDER BY Nome",
    "q06 Uniao": "SELECT UF AS Local FROM Clientes WHERE UF Is Not Null UNION SELECT Status FROM Pedidos ORDER BY Local",
    "q07 Implicita": "SELECT c.Nome, Sum(p.Valor) AS Total FROM Clientes c, Pedidos p WHERE c.ID = p.ClienteID "
                     "AND p.Data Between #1/1/2026# And #2/28/2026# GROUP BY c.Nome ORDER BY Sum(p.Valor) DESC",
    "q08 Param": "PARAMETERS [Valor minimo] Currency; SELECT PedidoID, Valor FROM Pedidos WHERE Valor >= [Valor minimo]",
    "q09 Sobre consulta": "SELECT q.Nome, q.Valor * 2 AS Dobro FROM [q02 Join] AS q WHERE q.Valor > 60",
    "q10 Subconsulta": "SELECT Nome FROM Clientes WHERE ID IN (SELECT ClienteID FROM Pedidos WHERE Valor > 150)",
    "q11 Concat": "SELECT Nome & ' (' & UF & ')' AS Rotulo, Len(Nome) AS L FROM Clientes ORDER BY ID",
    "q12 Total": "SELECT Count(*) AS N, Sum(Valor) AS Soma FROM Pedidos WHERE Status = 'Pago'",
}


def _importar(out, nome):
    sys.path.insert(0, str(out))
    for k in [k for k in sys.modules if k == nome or k.startswith(nome + ".")]:
        del sys.modules[k]
    return importlib.import_module(nome), importlib.import_module(nome + ".moderno")


def test_consultas_independentes(tmp_path):
    src = tmp_path / "banco"
    src.mkdir()
    (src / "Clientes.csv").write_text(
        "ID,Nome,UF,Segmento,Limite\n1,Ana Maria,PB,Varejo,1000\n2,Bruno Lima,PE,Atacado,5000\n"
        "3,Carla Souza,PB,Varejo,\n4,Davi Rocha,SP,Atacado,2500\n5,Eva Nunes,,Varejo,800\n", encoding="utf-8")
    (src / "Pedidos.csv").write_text(
        "PedidoID,ClienteID,Data,Valor,Status\n10,1,2026-01-05,100.0,Pago\n11,1,2026-02-10,50.5,Aberto\n"
        "12,2,2026-02-11,200.0,Pago\n13,3,2026-03-01,,Cancelado\n14,4,2026-03-02,10.0,Pago\n"
        "15,2,2026-03-15,999.9,Pago\n16,5,2026-01-20,75.25,Aberto\n", encoding="utf-8")
    for n, sql in CONSULTAS.items():
        (src / f"{n}.sql").write_text(sql, encoding="utf-8")
    res = transpile_database(str(src), str(tmp_path / "o"), package="indep", modern=True)
    assert res.modern["q10 Subconsulta"].startswith("fiel")
    assert sum(v == "moderna" for v in res.modern.values()) == len(CONSULTAS) - 1
    etapas = (tmp_path / "o" / "indep" / "moderno" / "etapas.py").read_text(encoding="utf-8")
    assert "import acc2py" not in etapas and "from acc2py" not in etapas and "rt." not in etapas  # pandas puro
    _, mod = _importar(tmp_path / "o", "indep")
    entradas = mod.ler_entradas({})
    r = mod.comparar_com_fiel(entradas, mod.Parametros(valor_minimo=60))
    assert (r["Resultado"] == "IGUAL").all(), r.to_string()


def test_rotina_em_sequencia(tmp_path):
    res = transpile_database(str(BANCO), str(tmp_path / "o"), package="rot", modern=True,
                             sequence=[l.strip() for l in (ROOT / "exemplos/rotina/sequencia.txt")
                                       .read_text(encoding="utf-8").splitlines()
                                       if l.strip() and not l.startswith("#")])
    assert all(v == "moderna" for v in res.modern.values())
    _, mod = _importar(tmp_path / "o", "rot")
    entradas = mod.ler_entradas({})
    p = mod.Parametros(desconto=0.1)
    final = mod.executar(entradas, p)
    assert list(final["Nome"]) == ["bruno lima", "Ana Maria"] and final["Total"].tolist() == [180.0, 150.5]
    assert (mod.comparar_com_fiel(entradas, p)["Resultado"] == "IGUAL").all()
    try:
        mod.executar(entradas, mod.Parametros())
        assert False
    except ValueError as e:
        assert "Desconto" in str(e)
