SELECT Nome, Count(*) AS Pedidos, Sum(Valor) AS Total
FROM tmpVendas
GROUP BY Nome
ORDER BY Sum(Valor) DESC;
