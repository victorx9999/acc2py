Option Compare Database

Public Sub Teste()
    Dim conta As clsConta
    Set conta = New clsConta
    conta.Titular = "Ana"
    conta.Deposita 1234.5
    Debug.Print conta.Resumo()
    Debug.Print Utilidades.CalculaDesconto(2000), CalculaDesconto(100, 0.2)
    Debug.Print Classifica("bruno"), Classifica("Helena"), Classifica("zeca")
    Debug.Print SomaAte(10), Iniciais("  ana maria  braga ")
    Debug.Print Pago, Situacao.Cancelado
End Sub
