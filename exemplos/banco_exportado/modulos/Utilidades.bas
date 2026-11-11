Attribute VB_Name = "Utilidades"
Option Compare Database
Option Explicit

' Funções gerais do sistema
Public Const TAXA_PADRAO As Double = 0.1
Private contador As Long

Public Enum Situacao
    Aberto = 1
    Pago
    Cancelado = 9
End Enum

Public Function CalculaDesconto(ByVal valor As Double, Optional ByVal pct As Variant) As Double
    If IsMissing(pct) Then pct = TAXA_PADRAO
    If valor <= 0 Then
        CalculaDesconto = 0
        Exit Function
    ElseIf valor > 1000 Then
        CalculaDesconto = Round(valor * (pct + 0.05), 2)
    Else
        CalculaDesconto = Round(valor * pct, 2)
    End If
    contador = contador + 1
End Function

Public Function Classifica(nome As String) As String
    Select Case UCase(Left$(nome, 1))
        Case "A" To "F"
            Classifica = "Grupo 1"
        Case "G", "H", "I"
            Classifica = "Grupo 2"
        Case Else
            Classifica = "Outros"
    End Select
End Function

Function SomaAte(n As Integer) As Long
    Dim i As Integer, total As Long
    For i = 1 To n
        If i Mod 2 = 0 Then GoTo Proximo
        total = total + i
Proximo:
    Next i
    SomaAte = total
End Function

Function Iniciais(ByVal nome As String) As String
    Dim partes() As String, i As Integer, r As String
    partes = Split(Trim(nome), " ")
    For i = LBound(partes) To UBound(partes)
        If Len(partes(i)) > 0 Then r = r & UCase(Left(partes(i), 1))
    Next
    Iniciais = r
End Function

Public Sub DesativaGrandes(limite As Currency)
On Error GoTo Trata_Erro
    Dim rs As DAO.Recordset
    Dim n As Long
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Clientes WHERE Ativo = True", dbOpenDynaset)
    Do While Not rs.EOF
        If Nz(DSum("Valor", "Pedidos", "ClienteID=" & rs!ID), 0) > limite Then
            rs.Edit
            rs!Ativo = False
            rs.Update
            n = n + 1
        End If
        rs.MoveNext
    Loop
    MsgBox "Clientes desativados: " & n, vbInformation, "Pronto"

Sair:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

Trata_Erro:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    Resume Sair
End Sub

Sub ExportaRelatorio(arquivo As String)
    DoCmd.SetWarnings False
    CurrentDb.Execute "UPDATE Pedidos SET Valor = Valor * 1.1 WHERE Valor < 60", dbFailOnError
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "Vendas por Cliente", arquivo, True
    DoCmd.SetWarnings True
    Debug.Print "Exportado:"; arquivo
End Sub

Function NovoCliente(nome As String) As Long
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("Clientes")
    With rs
        .AddNew
        !ID = DMax("ID", "Clientes") + 1
        !Nome = nome
        !Ativo = True
        .Update
    End With
    NovoCliente = DMax("ID", "Clientes")
    rs.Close
End Function

Function TotalCaixa() As Currency
    Dim c As New Collection, v As Variant, t As Currency
    c.Add 10.5
    c.Add 20
    For Each v In c
        t = t + v
    Next v
    TotalCaixa = t
End Function
