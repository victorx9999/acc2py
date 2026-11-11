SELECT Clientes.Nome, Count(Pedidos.PedidoID) AS Qtde, Sum(Pedidos.Valor) AS Total
FROM Clientes LEFT JOIN Pedidos ON Clientes.ID = Pedidos.ClienteID
GROUP BY Clientes.Nome
ORDER BY Sum(Pedidos.Valor) DESC;
