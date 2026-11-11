SELECT Clientes.Nome, Pedidos.Data, Pedidos.Valor INTO tmpVendas FROM Clientes INNER JOIN Pedidos ON Clientes.ID = Pedidos.ClienteID WHERE Pedidos.Valor Is Not Null;
