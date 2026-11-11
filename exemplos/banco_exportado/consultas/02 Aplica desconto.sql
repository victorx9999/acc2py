UPDATE tmpVendas SET tmpVendas.Valor = [Valor]*(1-[Desconto]) WHERE tmpVendas.Valor > 150;
