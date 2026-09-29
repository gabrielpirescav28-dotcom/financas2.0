# Finanças Pessoais — Planilha com Dashboard

![Dashboard](preview_dashboard.png)

- **`Financas_Pessoais_Dashboard.xlsx`**: a planilha pronta para usar (tema escuro).
- **`build_planilha.py`**: script que gera a planilha do zero (`pip install openpyxl`, depois `python build_planilha.py`). Para o tema claro, use `python build_planilha.py claro`. A paleta fica no topo do script.

Todos os números (Dashboard, Orçamento, Anual) são fórmulas do Excel. Você só lança as transações.

## Abas

| Aba | Para que serve |
|---|---|
| **Dashboard** | Indicadores do mês, 4 gráficos, Top 5 maiores gastos e orçado × realizado. |
| **Lancamentos** | Onde você lança receitas e despesas (tabela `tbLancamentos`). |
| **Orcamento** | Meta mensal por categoria de despesa e quanto já foi usado no mês escolhido. |
| **Anual** | Categorias × 12 meses do ano escolhido, com total, média e minigráficos. |
| **Metas** | Metas de poupança e dívidas/parcelamentos. |
| **Config** | Listas editáveis: categorias, contas, cartões, formas de pagamento, anos. |
| Calc (oculta) | Cálculos auxiliares dos gráficos. Não precisa mexer. |

## Como lançar uma transação

1. Vá à aba **Lancamentos** e clique na primeira linha vazia logo abaixo da tabela.
2. Preencha **Data**, **Descrição**, **Tipo** (Receita/Despesa), **Categoria**, **Conta** e **Forma de pagamento** (as quatro últimas têm lista suspensa) e o **Valor**.
3. O valor é **sempre positivo**: é o Tipo que diz se o dinheiro entrou ou saiu.
4. **Mês** e **Ano** são preenchidos sozinhos. A tabela cresce e tudo se atualiza.

## Como trocar o mês do Dashboard

No topo do **Dashboard**, escolha o mês na lista **MÊS** e o ano na lista **ANO**. O Dashboard, a aba Orçamento (realizado) e a aba Anual (ano) seguem essa seleção.

## Como adicionar uma categoria (ou conta, cartão, forma de pagamento)

1. Na aba **Config**, digite o novo item na primeira linha vazia logo abaixo da tabela certa (por exemplo, "Pets" embaixo de *Categoria de despesa*).
2. A tabela cresce sozinha e o item já aparece nas listas suspensas de Lançamentos.
3. Para uma categoria de **despesa** ter meta, adicione também uma linha na tabela da aba **Orcamento** com a categoria e o valor da meta.

Observações:
- A aba Anual tem 2 linhas livres para receitas novas e 3 para despesas novas. Elas são preenchidas sozinhas.
- O gráfico de rosca mostra as 7 maiores categorias do mês, e as demais aparecem somadas como "Outras".
- O Dashboard mostra as 12 categorias do orçamento que vêm na planilha. Se você adicionar mais categorias ao orçamento, elas entram na aba Orcamento e no KPI "% do orçamento", mas não na tabela nem no gráfico de barras do Dashboard. Para incluí-las, aumente as categorias em `ORCAMENTO` no script e gere a planilha de novo.
- Anos novos (depois de 2032) entram na tabela *Ano* da aba Config.

## Como apagar os dados de exemplo

Os cerca de 70 lançamentos fictícios (abril a setembro de 2026) têm **"EXEMPLO"** na coluna **Observação**.

1. Na aba **Lancamentos**, clique na seta de filtro de **Observação** e deixe marcado só `EXEMPLO`.
2. Selecione as linhas visíveis da tabela (clique nos números das linhas), clique com o botão direito e escolha **Excluir linha**.
3. Limpe o filtro. As fórmulas e os gráficos continuam funcionando com a tabela vazia.

As metas, dívidas e metas de orçamento também são exemplos: é só sobrescrever com os seus valores.

## Detalhes técnicos

- `.xlsx` sem macros. As fórmulas estão gravadas em inglês e o Excel em português as traduz ao abrir.
- Usa só funções compatíveis (SUMIFS, INDEX/MATCH, LARGE, IFERROR, DATE, OFFSET), sem funções de matriz dinâmica.
- Intervalos nomeados: `Categorias`, `Contas`, `FormasPagamento`, `MesSelecionado`, `AnoSelecionado`, `Meses`, `Tipos`, `Anos`, `CatReceita`, `CatDespesa`.
- Verificação feita: recálculo no LibreOffice sem nenhum erro (#REF!, #NAME?, #VALUE!, #DIV/0!, #N/A) em 1.171 fórmulas. Os KPIs, o Top 5 e o orçado × realizado conferem com somas feitas em Python para setembro, agosto, abril e março/2026 (mês sem dados).
