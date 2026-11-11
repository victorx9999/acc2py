SELECT TOP 2 v.Nome, v.Total, Iniciais(v.Nome) AS Sigla, Classifica(v.Nome) AS Grupo
FROM [Vendas por Cliente] AS v
WHERE v.Total Is Not Null
ORDER BY v.Total DESC;
