UPDATE Pedidos SET Pedidos.Valor = Round([Valor]*1.1,2) WHERE Pedidos.Valor < 60;
