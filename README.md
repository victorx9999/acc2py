# acc2py — transpilador de Access para Python

Converte **consultas SQL** e **módulos VBA** de um banco Microsoft Access (.accdb/.mdb)
em um pacote Python que roda sobre **pandas**. O código gerado é legível, comentado
com o SQL/VBA original e acompanhado de um `RELATORIO.md` com tudo que precisa de revisão.

## Instalação

```bash
pip install -e .                    # dentro desta pasta
pip install -e ".[windows]"         # + leitura direta do .accdb pelo Access (recomendado)
pip install -e ".[multiplataforma]" # + leitura do .accdb sem Access (Linux/macOS)
```

## Uso

```bash
# tudo: todas as consultas e todos os módulos VBA
acc2py vendas.accdb -o saida

# só as consultas que você escolher (as consultas que elas usam entram junto)
acc2py vendas.accdb -o saida --consultas "Vendas por Mês" "Clientes Ativos"

# consultas escolhidas + todo o VBA
acc2py vendas.accdb -o saida --consultas "Vendas por Mês" --modulos todos

# ver o que existe no banco antes de escolher
acc2py vendas.accdb --listar
```

Opções úteis: `--sem-vba`, `--exportar-dados` (copia as tabelas para o pacote),
`--locale en-US` (se o Access de origem estava em inglês), `--backend com|parser|pasta`,
`--pacote nome`.

Os nomes passados em `--consultas` não diferenciam maiúsculas; se um nome não existir,
o acc2py sugere os mais parecidos. Com `--consultas` e sem `--modulos`, só entram os
módulos VBA cujas funções são chamadas dentro das consultas escolhidas.

### Muitas consultas, uma saída só

Para bancos com dezenas de consultas, coloque os nomes num arquivo (um por linha;
`*` e `?` funcionam como curinga, `#` comenta):

```text
# lista.txt — consultas do fechamento mensal
Vendas por Mês
Clientes Ativos
Rel_*            # todas que começam com Rel_
```

```bash
# transpila tudo num pacote só E já grava os resultados num único Excel
acc2py vendas.accdb -o saida -f lista.txt -r fechamento.xlsx -p "Data inicial=01/01/2026" -p "Data final=31/01/2026"
```

- Tudo vai para **um pacote** (`saida/vendas/`): todas as consultas em `consultas.py`, os
  módulos em `modulos/`, um único `RELATORIO.md`. Consultas que dependem de outras são
  resolvidas automaticamente.
- `-r arquivo.xlsx` grava **uma aba por consulta + aba "Resumo"** (linhas e erro de cada uma);
  `-r pasta/` grava um `.csv` por consulta. Uma consulta com erro não derruba as outras.
- Consultas de **ação** (UPDATE/DELETE/INSERT) são transpiladas mas **nunca executadas** pelo `-r`.
- Parâmetros (`-p`) valem para todas as consultas que os usam; datas em `dd/mm/aaaa` ou `aaaa-mm-dd`.

Depois, para rodar de novo sem transpilar:

```python
from vendas import exportar
import pandas as pd
exportar("fevereiro.xlsx", parametros={
    "Data inicial": pd.Timestamp("2026-02-01"),        # comum a todas
    "Vendas por Mês": {"Ano": 2026},                    # só para esta consulta
})
```

### Rotinas em etapas (uma consulta depois da outra)

Há dois casos, e o acc2py trata os dois:

**1. Consulta que lê outra consulta** (a final usa a anterior no `FROM`): não precisa
fazer nada. Peça só a final; as anteriores entram como dependência e rodam sozinhas,
na ordem certa, quando a final é executada.

**2. Etapas que alteram dados** — criar tabela temporária, atualizar, acrescentar,
excluir e, no fim, a consulta de saída. Aqui a ordem importa. Liste as consultas
**na ordem em que rodam**:

```text
# sequencia.txt
01 Cria base            # SELECT … INTO tmpVendas
02 Aplica desconto      # UPDATE tmpVendas …
03 Acrescenta avulsos   # INSERT INTO tmpVendas …
04 Remove pequenos      # DELETE FROM tmpVendas …
05 Resultado final      # SELECT … FROM tmpVendas
```

```bash
acc2py vendas.accdb -o saida -s sequencia.txt -r saida_final.xlsx -p "Desconto=0,1"
```

```text
  [1/5] 01 Cria base: 3 linha(s) afetada(s)
  [2/5] 02 Aplica desconto: 1 linha(s) afetada(s)
  …
  [5/5] 05 Resultado final: 2 linha(s)
Saída final gravada em: saida_final.xlsx
```

- Cada etapa enxerga o que a anterior fez (as tabelas são alteradas **em memória**; o
  .accdb original **não é modificado**).
- O Excel tem a **saída final** (última consulta de seleção) e a aba **Etapas** (linhas
  afetadas, tempo, erro). `--etapas-intermediarias` grava também cada etapa de seleção.
- Se uma etapa falha, a rotina **para** e o log mostra qual falhou e quais não rodaram.
- Curingas seguem ordem natural: `Passo*` → Passo 1, Passo 2, Passo 10.
- A ordem fica salva em `SEQUENCIA` (em `consultas.py`); para rodar de novo:

```python
from vendas import executar_sequencia
executar_sequencia("janeiro.xlsx", parametros={"Desconto": 0.1})
```

Se a rotina no Access é disparada por **VBA** (`DoCmd.OpenQuery "…"` em sequência),
transpile o módulo com `--modulos NomeDoModulo` e chame a Sub convertida: ela executa as
consultas na mesma ordem. Se é uma **macro**, converta-a antes no Access
(*Ferramentas de Banco de Dados → Converter Macros em Visual Basic*).

### Usando o pacote gerado

```python
from vendas import abrir
import pandas as pd

db = abrir()   # lê as tabelas do .accdb (ou da pasta 'dados', se exportada)
df = db.run("Vendas por Mês", {"Data inicial": pd.Timestamp("2026-01-01")})
db.run("Reajuste de Preços")           # consulta de ação: devolve nº de linhas afetadas
db.sql("SELECT Count(*) FROM Clientes") # SQL do Access em tempo de execução

from vendas.modulos import Utilidades  # VBA transpilado
Utilidades.CalculaDesconto(2000)
```

## Como o banco é lido

| backend | Onde funciona | O que lê | Observações |
|---|---|---|---|
| `com` | Windows com **Access instalado** + pywin32 | SQL exato de cada consulta, todo o VBA, tabelas | Mais fiel. Macros/AutoExec do banco ficam desativados durante a leitura. |
| `parser` | qualquer sistema + `access-parser` | tabelas; SQL remontado de `MSysQueries`; VBA de `MSysAccessStorage` (experimental) | O SQL é equivalente, mas pode não ser idêntico ao texto original. |
| `pasta` | qualquer sistema | `*.sql`, saídas do `SaveAsText`, `*.bas`/`*.cls`, tabelas `*.csv`/`*.xlsx` | Útil quando o .accdb não pode sair da máquina com Access. |

`auto` (padrão) usa `pasta` para diretórios, `com` no Windows se possível e `parser` no resto.

## O que é convertido

**SQL do Access** → funções pandas: SELECT (DISTINCT, TOP n [PERCENT] com empates), INNER/LEFT/RIGHT JOIN
(inclusive aninhados e junções implícitas com vírgula), WHERE, GROUP BY, HAVING, ORDER BY, UNION [ALL],
TRANSFORM…PIVOT, subconsultas (IN, EXISTS, escalares, no FROM — correlacionadas também), PARAMETERS,
`[Forms]![F]![ctl]`, consultas usando outras consultas, UPDATE/DELETE/INSERT/SELECT INTO, ~90 funções
(IIf, Nz, Format, DateAdd, DLookup/DSum…) e funções VBA do próprio banco usadas dentro das consultas.

Semântica do Access preservada no runtime: Null em comparações (lógica de três valores), texto sem
diferenciar maiúsculas em `=`, GROUP BY, DISTINCT, JOIN e ORDER BY; chaves Null não casam em JOIN;
Null primeiro em ORDER BY ascendente; curingas `* ? # [a-z] [!x]` no LIKE; `&` tratando Null como vazio;
`\` e `Mod` com arredondamento do VBA; nomes de colunas `SomaDeX`/`Expr1000` como no Access em português.

**VBA** → módulos Python: Sub/Function/Property, módulos de classe (com `@property`), formulários
(`Me!txt`, `Me.lbl.Caption`, controles por nome), Dim/ReDim Preserve/arrays com limites, Const, Type,
Enum, If/ElseIf, Select Case (To, Is, listas), For/Step, For Each, Do/Loop, While/Wend, With,
Exit, GoTo para o fim do laço (vira `continue`), `On Error GoTo` no padrão corpo/saída/tratador
(vira `try/except`), `On Error Resume Next`, recordsets DAO (Edit/AddNew/Update/FindFirst/Fields),
`CurrentDb.Execute`, `DoCmd.RunSQL/OpenQuery/TransferSpreadsheet/TransferText/OutputTo`, Collection,
Scripting.Dictionary, MsgBox/InputBox (configuráveis) e ~100 funções nativas.

## Limitações conhecidas (sinalizadas no RELATORIO.md)

- **ByRef**: alterações em parâmetros não voltam para quem chamou (o Python não tem ByRef). Cada caso é marcado com `ATENÇÃO`.
- **GoTo arbitrário**, `Resume` fora do padrão, E/S de arquivo do VBA (`Open … For Input`), `Declare` (APIs do Windows) e eventos de classe viram `NotImplementedError` explícito.
- **Interface** (abrir formulários/relatórios, `Requery`, `SetFocus`) não tem equivalente: `DoCmd.OpenForm` etc. só registram em `vb.LOG`.
- **Currency** vira `float` (4 casas no `CCur`), não decimal exato.
- **Format/CStr/CDate** seguem a configuração regional escolhida (`pt-BR` por padrão).
- Subconsultas correlacionadas e `DLookup` linha a linha funcionam, mas são lentos em tabelas grandes: prefira reescrever como JOIN.
- Consultas pass-through/DDL e tabelas vinculadas a outros bancos não são convertidas.

## Validação recomendada

Para cada consulta convertida, rode no Access e no Python com os mesmos dados e compare com
`pandas.testing.assert_frame_equal(..., check_dtype=False)`. O `RELATORIO.md` lista tudo o que
foi aproximado ou não convertido.

## Exemplo incluído

```bash
acc2py exemplos/banco_exportado -o saida --consultas Ranking
cd saida && python exemplo_banco_exportado.py
```

## Testes

```bash
pip install -e ".[dev]" && pytest
```
