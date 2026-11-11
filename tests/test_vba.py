"""Testes do transpilador VBA -> Python (executam o código gerado)."""
import sys
import types

import pandas as pd
import pytest

from acc2py.diagnostics import Diagnostics
from acc2py.runtime import Database
from acc2py.runtime import vba as vb
from acc2py.vba.parser import parse_module
from acc2py.vba.to_python import transpile_modules


def load(sources: dict, kinds=None):
    """Transpila e importa os módulos como pacote em memória."""
    mods = [parse_module(code, name, (kinds or {}).get(name)) for name, code in sources.items()]
    d = Diagnostics()
    out = transpile_modules(mods, d)
    pkg = types.ModuleType("vbapkg")
    pkg.__path__ = []
    sys.modules["vbapkg"] = pkg
    loaded = {}
    for name, code in out.items():
        mod = types.ModuleType(f"vbapkg.{name}")
        mod.__package__ = "vbapkg"
        sys.modules[f"vbapkg.{name}"] = mod
        setattr(pkg, name, mod)
        loaded[name] = (mod, code)
    for name, (mod, code) in loaded.items():
        exec(compile(code, f"<{name}>", "exec"), mod.__dict__)
    return {n: m for n, (m, _) in loaded.items()}, out, d


def test_select_case_and_text_compare():
    m, code, _ = load({"M": '''Option Compare Database
Function F(x As String) As String
    Select Case x
        Case "a" To "f": F = "1"
        Case "x", "Y": F = "2"
        Case Else: F = "3"
    End Select
End Function
Function G(n As Long) As String
    Select Case n
        Case Is < 0
            G = "neg"
        Case 0
            G = "zero"
        Case 1 To 9, 100
            G = "ok"
    End Select
End Function'''})
    M = m["M"]
    assert [M.F("B"), M.F("y"), M.F("z")] == ["1", "2", "3"]
    assert [M.G(-1), M.G(0), M.G(5), M.G(100), M.G(50)] == ["neg", "zero", "ok", "ok", ""]


def test_loops_exit_and_goto_continue():
    m, _, _ = load({"M": '''Function Soma(n As Integer) As Long
    Dim i As Integer, t As Long
    For i = n To 1 Step -1
        If i = 3 Then GoTo Prox
        t = t + i
        If t > 100 Then Exit For
Prox:
    Next i
    Do
        t = t + 1
    Loop Until t Mod 5 = 0
    Soma = t
End Function'''})
    assert m["M"].Soma(5) == 15  # 5+4+2+1 = 12 -> 15


def test_on_error_pattern():
    m, code, _ = load({"M": '''Function Divide(a As Double, b As Double) As Variant
On Error GoTo Erro
    Divide = a / b
Sair:
    Exit Function
Erro:
    Divide = "Erro " & Err.Number
    Resume Sair
End Function
Function Seguro() As Long
    On Error Resume Next
    Dim x As Long
    x = 1 / 0
    Seguro = Err.Number
End Function'''})
    assert m["M"].Divide(6, 3) == 2
    assert m["M"].Divide(1, 0) == "Erro 11"
    assert m["M"].Seguro() == 11
    assert "except Exception" in code["M"]


def test_byref_warning_and_arrays():
    m, code, d = load({"M": '''Option Base 1
Sub Dobra(x As Long)
    x = x * 2
End Sub
Function Arr() As Long
    Dim a(3) As Long, i As Integer, p
    For i = 1 To 3
        a(i) = i * 10
    Next
    ReDim Preserve a(4)
    a(4) = 1
    p = Split("x,y,z", ",")
    Arr = a(1) + a(3) + a(4) + UBound(p)
End Function'''})
    assert m["M"].Arr() == 10 + 30 + 1 + 2
    assert "ByRef" in code["M"]
    assert any("ByRef" in i.message for i in d.items)


def test_class_with_properties_and_collection():
    m, _, _ = load({
        "Conta": '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Conta"
Private mSaldo As Double
Public Property Get Saldo() As Double
    Saldo = mSaldo
End Property
Public Property Let Saldo(v As Double)
    If v < 0 Then Err.Raise 5, , "negativo"
    mSaldo = v
End Property''',
        "Uso": '''Function Total() As Double
    Dim c As New Collection, k As Conta, i As Integer
    For i = 1 To 3
        Set k = New Conta
        k.Saldo = i * 1.5
        c.Add k
    Next
    For Each k In c
        Total = Total + k.Saldo
    Next
End Function'''}, kinds={"Conta": "class"})
    assert m["Uso"].Total() == 9.0
    with pytest.raises(vb.VBAError):
        m["Conta"].Conta().Saldo = -1


def test_recordset_and_dao():
    db = Database({"T": pd.DataFrame({"ID": [1, 2, 3], "Nome": ["a", "b", "c"], "V": [1.0, 2.0, 3.0]})})
    vb.set_database(db)
    m, _, _ = load({"M": '''Option Compare Database
Function Processa() As Double
    Dim rs As DAO.Recordset, t As Double
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM T WHERE V > 1", dbOpenDynaset)
    With rs
        Do While Not .EOF
            t = t + !V
            .Edit
            !Nome = UCase(!Nome)
            .Update
            .MoveNext
        Loop
        .Close
    End With
    Set rs = CurrentDb.OpenRecordset("T")
    rs.AddNew
    rs!ID = DMax("ID", "T") + 1
    rs("Nome") = "novo"
    rs.Update
    rs.FindFirst "Nome = 'B'"
    If Not rs.NoMatch Then t = t + rs.Fields("ID").Value * 100
    CurrentDb.Execute "UPDATE T SET V = 0 WHERE ID = 1"
    Processa = t
End Function'''})
    assert m["M"].Processa() == 5.0 + 200
    t = db.tables["T"]
    assert list(t["Nome"]) == ["a", "B", "C", "novo"] and t.loc[0, "V"] == 0


def test_form_module_controls():
    vb.set_database(Database({"P": pd.DataFrame({"C": [1, 1], "V": [2.0, 3.0]})}))
    m, _, _ = load({"Form_F": '''Option Compare Database
Option Explicit
Private Sub btn_Click()
    If IsNull(Me!txtC) Then Exit Sub
    Me.lblTotal.Caption = "Total: " & DSum("V", "P", "C=" & txtC)
    txtStatus = "ok"
End Sub'''})
    f = m["Form_F"].Form_F(txtC=1)
    f.btn_Click()
    assert f.item("lblTotal").Caption == "Total: 5" and f["txtStatus"] == "ok"


def test_misc_functions():
    from acc2py.runtime import vbafuncs as V
    from datetime import datetime
    assert V.DateDiff("m", datetime(2026, 1, 31), datetime(2026, 2, 1)) == 1
    assert V.DateAdd("m", 1, datetime(2026, 1, 31)) == datetime(2026, 2, 28)
    assert V.Weekday(datetime(2026, 11, 11)) == 4  # quarta-feira
    assert V.Format(datetime(2026, 3, 5, 14, 7), "dd/mm/yyyy hh:nn") == "05/03/2026 14:07"
    assert V.Round(2.5) == 2 and V.Round(3.5) == 4 and V.CInt("2,5") == 2
    assert V.InStr("Banana", "na") == 3 and V.Mid("abcdef", 2, 3) == "bcd"
    assert V.Val("12abc") == 12 and V.Format(0.256, "0.0%") == "25,6%"
