# Finanças Pessoais — Central da Família

![Início](preview_inicio.png)

![Caderno do mês](preview_caderno.png)

![Dashboard](preview_dashboard.png)

*(imagens feitas com os dados de exemplo)*

## No jeito do Plínio

A planilha segue o jeito de visualizar e controlar da **Finanças Família 3.0** e da **Tela do Dinheiro do Plínio**:

- **Tema do Plínio.** Fundo escuro, **latão** como destaque, **sereno** (verde) e **brasa** (vermelho), fonte Segoe UI. Células em latão são as que você edita, como o "amarelo = você edita" dos Planos.
- **Caderno do mês.** É o antigo Painel Mensal.
  - **Mês de trabalho** fica no topo e todo o resto segue esse mês.
  - Três blocos: **Receitas · Gastos fixos · Gastos extras**, com a grade **Data · Descrição · Parcela · Valor · Vencimento · Situação · Categoria**.
  - **Resumo do mês** com os seus números: Saldo que veio do mês anterior, Recebido, Gastos fixos, Gastos extras, Resultado, % comprometido da renda, Aporte, Resgate, Saldo livre, Saldo acumulado, Total investido, IR, Líquido e o previsto a pagar/receber.
  - **Semáforo:** verde = pago/recebido, amarelo = pendente que vence em até 3 dias, vermelho = atrasado, roxo = Descontar Depois. Pendente aparece, mas **não entra no total**.
  - **Conferência do mês:** pendências vencidas, sem categoria, possíveis lançamentos em dobro e gastos vs. mês anterior. O fechamento é o ritual de declarar que revisou: você escreve a data em **Conferido em**, na aba Histórico.
- **Aporte e Saldo livre.**
  - **Aporte** = aplicações novas da aba Investimentos no mês, sem as marcadas como "Reaplicação" e sem linha de "correção". Confere com o aporte da Tela do Dinheiro.
  - **Resgate** = valores negativos na aba Investimentos, a sua convenção de resgate parcial.
  - **Saldo livre** = Resultado − Aporte + Resgate.
  - **Saldo que veio do mês anterior** = soma do saldo livre dos meses desde **Conta a partir de**, um campo editável no Caderno.
- **Potes 80 / 10 / 7 / 3.** Investimento, Lazer, Manutenção e Vestimenta, com **Guardado** e **% alvo** editáveis. Ao lado ficam o % real, o desvio e o **gasto no ano / no mês de cada pote**. A barra "guardou × gastou" mostra o dinheiro pelos dois lados. O **reequilíbrio** é só sugestão: some o valor gasto nos potes na proporção do alvo, e você aprova editando o Guardado.
- **Categoria ≠ Pote.** A categoria diz *com o quê*; o pote (fundo) diz *de qual dinheiro*. Cada categoria tem um **Fundo padrão** (aba Config), e um gasto sem pote recebe o fundo da categoria, a mesma regra da Base_Dados.
- **Mapa completo de categorias.** A aba Atalhos (Favoritos) tem **todas as descrições do histórico** → Tipo, Categoria, Pote. Digitou uma descrição conhecida, e o resto se preenche.
- **🤖 Correções IA.** É o seu canal de pedidos, como a Melhorias IA: você escreve o **Pedido**, a IA implementa, marca **FEITO ✔** e explica em **Como ficou**. Os pedidos são mantidos quando a planilha é gerada de novo.
- **Quem.** Sempre a pessoa definida em `config_pessoal.json` (hoje, uma só), preenchida sozinha (Correção IA #1).

## Preenchimento rápido (o dia a dia)

| Ferramenta | O que faz |
|---|---|
| **Início** | Primeira aba. Tem os botões **✚ Novo lançamento** (leva à primeira linha vazia), **📒 Caderno do mês**, **✔ Contas do mês**, **⟳ Parcelar / repetir**, **◉ Potes** e **Dashboard**. Mostra saldo livre, contas fixas lançadas, a pagar, % comprometido, a receber e investimentos, além dos **próximos vencimentos e atrasados** e dos **últimos lançamentos**. |
| **Lançamentos** | Você digita **Data, Descrição e Valor** (e Parcela/Vencimento quando houver). As colunas roxas (Tipo, Categoria, Situação, Pote, Quem, Conta, Forma e o Valor de contas fixas) **se preenchem sozinhas**. Dá para digitar por cima em qualquer linha. Conta, forma, competência etc. ficam recolhidas no **[+]**. |
| **Contas do mês** | Checklist das contas recorrentes do mês (✔ pago, ⏳ pendente, ✖ falta lançar). Ao lado vem um **bloco pronto para colar** com o que falta: copie, vá à primeira linha vazia e cole só os valores (Ctrl+Shift+V). |
| **Parcelar** | Informe descrição, valor, nº de parcelas e a 1ª data: as linhas 1/n, 2/n… saem prontas para colar. O modo "Repetir todo mês" serve para o que se repete. |

**Lançar em 10 segundos:** ✚ Novo lançamento ➜ `Ctrl+;` (data de hoje) ➜ Tab ➜ descrição (`Alt+↓` abre a lista) ➜ Tab ➜ valor ➜ Enter.

## Arquivos

- **`build_planilha.py`**: gera a planilha do zero (`pip install openpyxl`).
  - `python build_planilha.py --dados Financas_Plinio_AAAA-MM-DD.xlsx` importa a exportação do app.
  - `python build_planilha.py --dados Financas_Pessoais_Dashboard.xlsx` gera uma **versão nova a partir da própria planilha**. Mantém lançamentos, favoritos, potes, correções IA, meses conferidos, orçamento, metas, dívidas, investimentos e acertos.
  - `python build_planilha.py` sem dados gera **`Financas_Pessoais_Dashboard_EXEMPLO.xlsx`** (dados fictícios).
  - Opções: `--tema plinio` (padrão), `escuro` ou `claro`; `--saida Nome.xlsx`.
- **`Financas_Pessoais_Dashboard_EXEMPLO.xlsx`**: a mesma planilha com dados fictícios.

A planilha com dados reais, a exportação do app e as prévias com dados reais **não vão para o git** (`.gitignore`), porque o repositório é público.

## Abas

| Aba | Para que serve |
|---|---|
| **Inicio** | Central da família: botões, status do mês, vencimentos e últimos lançamentos. |
| **Caderno** | O mês de trabalho em três blocos, com resumo, conferência e semáforo. **É aqui que se troca o mês.** |
| **Dashboard** | Recebido, Gastos, Aporte, Saldo livre, % comprometido e % do orçamento, com 4 gráficos, Top 5 e orçado × realizado. |
| **Lancamentos** | Tabela `tbLancamentos`: onde tudo é lançado. |
| **ContasMes** | Checklist das contas recorrentes e o bloco do que falta lançar. |
| **Parcelar** | Gerador de parcelas ou de lançamentos repetidos. |
| **Orcamento** | Meta mensal por categoria e quanto já foi usado. |
| **Potes** | Guardado × gasto dos potes 80/10/7/3 e sugestão de reequilíbrio. |
| **Anual** | Categorias × 12 meses do ano, com minigráficos. |
| **Historico** | Todos os meses: receitas, gasto fixo, gasto extra, resultado, aporte, resgate, saldo livre, % comprometido e **Conferido em**. |
| **Metas** | Metas de poupança e parcelamentos em andamento. |
| **Investimentos** | Aplicações importadas do app e o total da carteira. |
| **Acertos** | Valores a cobrar, com o rateio passo a passo. |
| **Atalhos** | Favoritos / mapa Descrição → Tipo, Categoria e Pote. |
| **Correções IA** | Os seus pedidos para a IA. |
| **Config** | Listas editáveis: categorias (com Fundo padrão), potes, Quem, contas, cartões, formas, tipos, situações e anos. |
| Calc (oculta) | Cálculos auxiliares. |

## Regra que decide o que entra nos totais

A coluna **Conta no mês?** segue a regra do app: **sim** para receita *Recebida* e para gasto *Pago* ou *Descontar Depois* (Situação em branco também conta); **não** para *Investimento*, *Pendente* e parcelas projetadas. O mês vem da **Data**; se o lançamento é de outro mês, preencha a **Competência**.

## Sobre a importação dos dados reais

- **Categorias em branco:** recebem a categoria mais usada para a mesma descrição, com a Observação "categoria sugerida pelo histórico". Sem nenhum histórico, a linha fica **Sem categoria**.
- **Gastos sem pote:** recebem o Fundo padrão da categoria. "Sem categoria" continua sem pote.
- **Guardado dos potes, pessoas e "Conta a partir de":** vêm do arquivo local `config_pessoal.json` (fora do git, porque o repositório é público). Depois ficam na própria planilha (abas Potes, Config e Caderno).
- **Investimentos:** *Saldo hoje*, *IR* e *Líquido* são a posição do app no dia da exportação e são valores fixos. Os totais e o Aporte/Resgate do mês são fórmulas.

## Detalhes técnicos

- `.xlsx` sem macros. As fórmulas estão em inglês, e o Excel em português as traduz.
- Funções compatíveis: SUMIFS, SUMPRODUCT, INDEX/MATCH, SMALL/LARGE em fórmula matricial, IFERROR, OFFSET. Nada de matriz dinâmica.
- As colunas automáticas de Lançamentos são colunas calculadas da tabela, e o aviso "fórmula inconsistente" fica desligado nas linhas antigas.
- Verificação:
  - **Recálculo:** no LibreOffice, nenhuma das 16.456 fórmulas da planilha com dados reais dá erro.
  - **Histórico:** bate com o "Resumo por mês" do app nos 54 meses.
  - **Caderno:** mostra todos os lançamentos de set/2026, ago/2026 e mar/2023.
  - **Aporte:** confere com a Tela do Dinheiro do Plínio.
  - **Linhas novas:** o preenchimento automático (Quem automático, Pote pela categoria) e os blocos para colar foram testados simulando linhas novas.
  - **Excel de verdade:** os cálculos foram conferidos com a cópia salva por ele.
