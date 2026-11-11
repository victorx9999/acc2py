VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConta"
Attribute VB_PredeclaredId = False
Option Compare Database
Option Explicit

Private mSaldo As Double
Private mTitular As String

Private Sub Class_Initialize()
    mSaldo = 0
End Sub

Public Property Get Saldo() As Double
    Saldo = mSaldo
End Property

Public Property Let Titular(ByVal v As String)
    mTitular = v
End Property

Public Property Get Titular() As String
    Titular = mTitular
End Property

Public Sub Deposita(valor As Double)
    If valor <= 0 Then Err.Raise vbObjectError + 1, "clsConta", "Valor inválido"
    mSaldo = mSaldo + valor
End Sub

Public Function Resumo() As String
    Resumo = Titular & ": " & Format(Saldo, "#,##0.00")
End Function
