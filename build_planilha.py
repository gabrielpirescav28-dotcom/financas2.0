# -*- coding: utf-8 -*-
"""
Gera a planilha de finanças pessoais com dashboard do zero.

    python build_planilha.py                                   # dados de exemplo
    python build_planilha.py --dados Financas_Plinio.xlsx      # importa os dados reais
    python build_planilha.py --tema claro                      # tema claro
    python build_planilha.py --dados X.xlsx --saida Minha.xlsx # escolhe o nome do arquivo

Todos os números do Caderno, Orçamento, Anual e Histórico são fórmulas vivas do Excel:
o script só escreve os lançamentos, as metas e as listas de configuração.
Fórmulas em inglês com vírgula como separador (o Excel pt-BR traduz ao abrir).
"""
import argparse
import collections
import json
import datetime as dt
import math
import os
import re
import sys
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import (CharacterProperties, Font as DFont, Paragraph,
                                   ParagraphProperties, RegularTextRun, RichTextProperties)
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

# =============================================================================
# CONFIGURAÇÃO GERAL
# =============================================================================
_ap = argparse.ArgumentParser(description="Gera a planilha de finanças com dashboard.")
_ap.add_argument("tema_pos", nargs="?", choices=["plinio", "plinio_escuro", "escuro", "claro"], help=argparse.SUPPRESS)
_ap.add_argument("--tema", choices=["plinio", "plinio_escuro", "escuro", "claro"], help="tema visual (padrão: plinio)")
_ap.add_argument("--dados", help="planilha exportada do app (Financas_Plinio_*.xlsx) para importar")
_ap.add_argument("--saida", help="nome do arquivo gerado")
ARGS = _ap.parse_args() if __name__ == "__main__" else _ap.parse_args([])

TEMA = ARGS.tema or ARGS.tema_pos or "plinio"   # "plinio" (padrão), "plinio_escuro", "escuro" ou "claro"
NOME_ARQUIVO = "Financas_Pessoais_Dashboard.xlsx"
# O tema "plinio" veste a identidade do app Plínio (plinio_ui.py): tinta escura, latão como
# acento, sereno = verde, brasa = vermelho. A fonte do app cai em "Segoe UI" no Windows.
# Correção IA: "letras padronizadas na cor preta, não dá pra ver nada" — o que se digita numa célula nova
# no Excel sai em preto; por isso o tema "plinio" padrão é claro (papel, letra preta, latão só como acento).
# O visual escuro original continua disponível como "plinio_escuro".
PLINIO = TEMA.startswith("plinio")
FONTE = "Segoe UI" if PLINIO else "Calibri"
LARG_CARD = 11 if FONTE == "Segoe UI" else 10      # Segoe UI é mais larga que a Calibri
KPI_PT = 19 if FONTE == "Segoe UI" else 22

PALETAS = {
    "plinio": dict(
        fundo="FFFFFF", card="FFFFFF", card2="F6F1E1", texto="000000", suave="5B6670",
        borda="DDD5BD", grade="E8E2CF", destaque="A8841A", positivo="3E7C52",
        alerta="B07D0A", negativo="B3413A", cabecalho_txt="E0B73A", roxo="6A52AE",
        heat_max="EFB3AE", cab_fundo="18212A", azul="2F6F9F", botao_txt="0F151B",
    ),
    "plinio_escuro": dict(
        fundo="0F151B", card="18212A", card2="1E2833", texto="E6EDF3", suave="8496A6",
        borda="26333F", grade="26333F", destaque="C9A227", positivo="7FA98A",
        alerta="C9A227", negativo="C0605A", cabecalho_txt="C9A227", roxo="A38BD1",
        heat_max="5A2E2E", cab_fundo="1E2833", azul="6FA3C7", botao_txt="0F151B",
    ),
    "escuro": dict(
        fundo="121417", card="1E2228", card2="262B33", texto="E8EAED", suave="9AA0A6",
        borda="2C323B", grade="2A2F37", destaque="4F8CFF", positivo="2ECC71",
        alerta="F5B942", negativo="FF5C5C", cabecalho_txt="FFFFFF", roxo="9B7BFF",
        heat_max="6B2F36", cab_fundo="4F8CFF", azul="4F8CFF", botao_txt="FFFFFF",
    ),
    "claro": dict(
        fundo="F7F8FA", card="FFFFFF", card2="EEF1F5", texto="1F2937", suave="6B7280",
        borda="E5E7EB", grade="E5E7EB", destaque="2563EB", positivo="16A34A",
        alerta="D97706", negativo="DC2626", cabecalho_txt="FFFFFF", roxo="7C3AED",
        heat_max="F8C4C4", cab_fundo="2563EB", azul="2563EB", botao_txt="FFFFFF",
    ),
}
P = PALETAS[TEMA]

def misturar(c1, c2, t):
    """Mistura duas cores hex (t=0 → c1, t=1 → c2)."""
    a = [int(c1[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


BARRA = misturar(P["destaque"], P["card"], 0.45)       # barra de dados discreta (texto continua legível)
BARRA_META = misturar(P["positivo"], P["card"], 0.45)

# Cores das categorias nos gráficos (derivadas da paleta, nada de cores padrão do Excel)
CORES_CATEGORIAS = ([P["destaque"], P["positivo"], P["azul"], P["negativo"], P["roxo"], "5FB3A8", "D08C5B", "8496A6"]
                    if PLINIO else
                    [P["destaque"], P["positivo"], P["alerta"], P["negativo"], P["roxo"], "2EC4B6", "FF9F43", "8A94A6"])

FMT_MOEDA = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
FMT_MOEDA_ZERO_TRACO = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"–"'
FMT_DATA = "dd/mm/yyyy"
FMT_PCT = "0.0%"
FMT_VAR = '"▲ "0.0%;"▼ "0.0%;0.0%'

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
         "Setembro", "Outubro", "Novembro", "Dezembro"]
MESES_ABREV = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


FMT_COMPETENCIA = "mm/yyyy"

# Colunas da tabela tbLancamentos (ordem na aba): o essencial primeiro, detalhes depois
# (as seis primeiras são a grade do caderno do Plínio: Data · Descrição · Parcela · Valor · Vencimento · Situação)
# "Descontar de" fica visível logo depois de Quem (Correção IA: com Descontar Depois, dizer de quem descontar)
COLS_LANC = ["Data", "Descrição", "Parcela", "Valor", "Vencimento", "Situação", "Tipo", "Categoria",
             "Pote", "Quem", "Descontar de", "Conta", "Forma de pagamento", "Competência", "Subcategoria",
             "Observação", "Mês", "Ano", "Conta no mês?"]
LARG_LANC = [13.5, 30, 10, 15, 14.5, 17, 12.5, 23, 13, 11, 14, 17, 18, 12, 16, 32, 6.5, 7.5, 10.5]
# Colunas preenchidas sozinhas a partir dos Favoritos (podem ser sobrescritas linha a linha)
COLS_AUTO = ["Valor", "Tipo", "Categoria", "Situação", "Pote", "Quem", "Conta", "Forma de pagamento"]
# Colunas dos blocos prontos para colar (Contas do mês / Parcelar): as 12 primeiras de Lançamentos
COLS_BLOCO = COLS_LANC[:13]
# Dados pessoais ficam FORA do código (o repositório é público): config_pessoal.json, ao lado do
# script e no .gitignore. Chaves: "pessoas" (lista fixa para a coluna Quem — Correção IA #1),
# "potes" ([pote, % alvo, guardado] da Alocação por Projeto) e "conta_a_partir" ([ano, mês]).
CONFIG_PESSOAL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_pessoal.json")
_CFG = json.load(open(CONFIG_PESSOAL, encoding="utf-8")) if os.path.exists(CONFIG_PESSOAL) else {}
PESSOAS = _CFG.get("pessoas") or None          # None = usa a lista que já estiver na planilha
LARG_BLOCO = {"Data": 13.5, "Descrição": 26, "Parcela": 9, "Valor": 14, "Vencimento": 13.5, "Situação": 12, "Tipo": 12,
              "Categoria": 20, "Pote": 12, "Quem": 10, "Descontar de": 12, "Conta": 14, "Forma de pagamento": 16}
# Potes (fundos): % alvo e "guardado" da Alocação por Projeto — só na importação do app; depois ficam na aba Potes.
POTES_PLINIO = [tuple(x) for x in _CFG.get("potes", [])]
# "Conta a partir de" do Painel Mensal: desde quando somar o saldo que veio dos meses anteriores
CONTA_A_PARTIR_PLINIO = tuple(_CFG["conta_a_partir"]) if _CFG.get("conta_a_partir") else None
# Respostas da aba Correções IA para pedidos já implementados (casados por palavra-chave)
_QUEM = (PESSOAS or ["a pessoa da lista"])[0]
CORRECOES_FEITAS = [
    (r"\bquem\b", "FEITO ✔",
     f"A coluna Quem agora se preenche sozinha com {_QUEM} em todo lançamento novo, e o histórico inteiro ficou "
     f"como {_QUEM}. A lista de Config (tabela Quem) tem só {_QUEM} — se um dia entrar outra pessoa, é só acrescentar lá."),
]
CORRECOES_FEITAS[:0] = [       # rodada 3 (antes das outras: os textos se parecem com pedidos antigos)
    (r"g[aá]s.*atrasad", "FEITO ✔",
     "A conta de gás que aparecia como atrasada estava como Pendente; foi marcada como Pago. Ela passa a contar no mês "
     "de competência dela e saiu de Atrasados no Início."),
    (r"descont[ae]r? depois|desconta depois", "FEITO ✔",
     "A coluna Descontar de agora fica visível logo depois de Quem, com lista de pessoas (Alt+↓). Situação = Descontar "
     "Depois sem nome: a célula fica vermelha e aparece na Conferência do Caderno. A aba Acertos virou automática: soma por "
     "pessoa tudo o que foi lançado como Descontar Depois no mês do Caderno, mostra lançamento a lançamento, e o rateio da "
     "casa usa esse total. Quando a pessoa pagar, registre em Recebidos (data, pessoa, valor): a coluna Falta desconta sozinha. "
     "Pessoa nova: escreva o nome na linha vazia da tabela de Acertos. O total a receber aparece no Início."),
    (r"selic", "FEITO ✔",
     "Conferido e corrigido: antes, uma taxa nova recalculava também o passado. Agora há o Histórico do CDI na aba Planos: "
     "cada linha vale a partir da sua data. Mudou a Selic? Acrescente uma linha (data + CDI novo): o que já rendeu fica com a "
     "taxa antiga e só o rendimento daqui para frente usa a nova. As projeções usam o CDI da linha mais recente."),
    (r"interligad|diminuir no saldo", "FEITO ✔",
     "Está ligada: toda aplicação lançada em Investimentos entra como Aporte do mês (pela Data), e o Saldo livre = Resultado − "
     "Aporte + Resgate diminui na hora no Caderno, no Início e no Histórico. Testado: uma aplicação de R$ 10.000 aumentou o "
     "aporte do mês em R$ 10.000 e baixou o saldo livre no mesmo valor. A carteira (Saldo hoje, Líquido) alimenta também o "
     "Caderno, o Início e os Planos. Reaplicação não conta como aporte. O Guardado da aba Potes continua sendo digitado por você."),
    (r"introdu|linha 16", "FEITO ✔",
     "A explicação do topo de Contas do mês ficou recolhida (clique no [+] à esquerda, na linha 8, para abrir) e a lista "
     "começa logo abaixo dos cartões: sobra mais espaço para as contas."),
    (r"unir plano|[uú]nica aba", "FEITO ✔",
     "Plano 1 e Plano 2 agora estão numa aba só, Planos: premissas e Plano 1 à esquerda, Plano 2 e a carreira à direita, o "
     "Histórico do CDI no meio e, embaixo, as duas projeções lado a lado (48 e 96 meses). O botão ◎ Plano 1 e 2 do Início leva até ela."),
    (r"ocult", "FEITO ✔",
     "Ficaram ocultas a Config (listas de categorias, potes, contas…) e a Calc (cálculos internos). Para ver: botão direito em "
     "qualquer aba → Reexibir. As demais são de preenchimento ou de consulta; se quiser esconder mais alguma (Anual, Histórico, "
     "Atalhos…), é só pedir."),
]
CORRECOES_FEITAS += [
    (r"letras.*pret|cor preta", "FEITO ✔",
     "Tema novo, claro: fundo branco e letra preta em todas as abas; o latão ficou só nos cabeçalhos, nos botões e no que você "
     "edita. O que você digita numa linha nova agora aparece em preto sobre branco. (O visual escuro antigo continua "
     "existindo: --tema plinio_escuro.)"),
    (r"data.*autom|dia de hoje", "FEITO ✔",
     "A coluna Data de Lançamentos (e a de Investimentos) agora tem uma lista: Alt+↓ e Enter põe a data de HOJE (tem também "
     "ontem e anteontem). Ctrl+; continua valendo. Uma data que se preenche sozinha e fica fixa só seria possível com macro "
     "(a fórmula HOJE() mudaria todo dia), por isso a lista. Linha sem data fica amarela e aparece na Conferência do Caderno "
     "(\"Lançamentos sem data\"); as linhas de teste que estavam sem data entraram com a data da atualização, marcadas na Observação."),
    (r"saldo acumulado", "FEITO ✔",
     "Saíram do Caderno o \"Saldo que veio do mês anterior\", o \"Saldo acumulado\" e o \"Conta a partir de\". O resumo mostra "
     "só o Resultado do mês e o Saldo livre do mês (Resultado − Aporte + Resgate): o negativo do mês que você acompanha."),
    (r"dashboard", "FEITO ✔",
     "Aba Dashboard removida: o Caderno é a visão do mês. No Início, o botão do Dashboard virou ◎ Plano 1 e 2. Recebido, gastos, "
     "aporte, saldo livre, % comprometido e % do orçamento continuam no Caderno e no Início."),
    (r"contas do m", "FEITO ✔",
     "Contas do mês = checklist das contas que se repetem todo mês (as marcadas Recorrente? = sim na aba Atalhos). Para o mês "
     "de trabalho do Caderno mostra ✔ Pago, ⏳ Pendente ou ✖ Falta lançar, com o previsto e o já lançado. À direita vem um bloco "
     "pronto só com o que falta: copie, vá à primeira linha vazia de Lançamentos e cole só os valores (Ctrl+Shift+V); entram "
     "como Pendente. Uso: no começo do mês cole o bloco; quando pagar, troque para Pago. Há uma caixa \"Como funciona\" na aba."),
    (r"parcelar", "FEITO ✔",
     "Parcelar = gerador de linhas. Você preenche descrição, valor (total ou da parcela), nº de parcelas e a data da 1ª; o bloco "
     "à direita monta 1/10, 2/10… já com tipo, categoria e pote. Copie e cole só os valores em Lançamentos: as futuras entram "
     "como Pendente e só contam no mês quando virarem Pago. O modo \"Repetir todo mês\" serve para aluguel, mesada e assinatura. "
     "Os parcelamentos em andamento (antes na aba Metas) agora ficam mais abaixo, nesta aba. Há uma caixa \"Como funciona\"."),
    (r"or[cç]amento", "FEITO ✔",
     "Orçamento = um teto mensal por categoria. A Meta mensal (latão) começou pela média dos seus últimos 12 meses; Realizado = "
     "o que foi pago naquela categoria no mês de trabalho do Caderno; Diferença e % usado (verde até 80%, amarelo até 100%, "
     "vermelho acima). Mude a meta digitando por cima. O Início mostra o % do orçamento usado. Há uma caixa \"Como funciona\"."),
    (r"sem categoria|\boutros\b", "FEITO ✔",
     "Nova aba Categorizar: as {n_rev} descrições que estavam em Outros ou Sem categoria ({n_lanc} lançamentos), agrupadas. "
     "Preenchi a Categoria certa onde deu para ter certeza ({n_certa}); as outras têm Sugestão e uma pergunta, e as dúvidas "
     "foram para o chat. Escolheu a categoria (Alt+↓): todos os lançamentos com aquela descrição e o favorito mudam juntos, e o "
     "Pote segue o fundo padrão. A Conferência do Caderno mostra quantas faltam."),
    (r"plano 1|aba meta", "FEITO ✔",
     "A aba Metas saiu e entraram Plano 1 e Plano 2, como na Finanças Família 3.0. Plano 1: premissas em latão (CDI, renda a "
     "cobrir, aporte planejado ou média real dos 3 últimos meses); com a carteira da aba Investimentos calcula % do CDI médio, "
     "IR médio real, juro líquido por mês, patrimônio-alvo, % do caminho e a projeção de 48 meses até INDEPENDENTE. Plano 2: "
     "renda durável = salário pela carreira (tabela editável) + aluguéis + juro líquido, com projeção de 96 meses até a meta. "
     "O % do Plano 1 aparece no Início."),
    (r"investimento", "FEITO ✔",
     "Investimentos agora rende sozinho. Para lançar: primeira linha vazia da tabela → Data (Alt+↓ = hoje) · Banco · % do CDI "
     "(ex.: 110%) · Aplicado; Isento = sim para LCI/LCA. Dias, Alíquota IR (22,5% → 15%), Saldo hoje (CDI × % do CDI em dias "
     "úteis), IR e Líquido se calculam. Reaplicação: escreva \"Reaplicação\" na Observação (não conta como aporte). Resgate total: "
     "Resgatado = sim; parcial: linha com valor negativo. Não precisa lançar de novo em Lançamentos: o Aporte do mês vem daqui. "
     "As aplicações vindas do app partem da posição calculada pelo app e seguem rendendo."),
]
PEDIDO_JEITO_PLINIO = ("Deixar a planilha no jeito do Plínio: como eu visualizo e controlo tudo (pedido no chat, 29/09/2026)", "FEITO ✔",
                       "Tema do Plínio (latão, sereno, brasa); Caderno do mês em três blocos com Mês de trabalho, resumo do mês e "
                       "semáforo (verde pago, amarelo vence em até 3 dias, vermelho atrasado); Lançamentos na ordem da grade "
                       "(Data · Descrição · Parcela · Valor · Vencimento · Situação); Aporte e Saldo livre; aba Potes 80/10/7/3 com "
                       "guardado × gasto e sugestão de reequilíbrio; mapa completo Descrição → Categoria → Fundo; conferência e "
                       "fechamento do mês no Histórico.")
# Colunas da tabela de Investimentos (as do app + a posição que o app calculou, de onde a linha segue rendendo)
CAB_INV = ["Data", "Banco", "% do CDI", "Aplicado", "Dias", "Isento", "Resgatado", "Alíquota IR", "Saldo hoje",
           "IR", "Líquido", "Vencimento", "Observação", "Base (app)", "Data da base"]
# Plano 1 e Plano 2 (Finanças Família 3.0): os valores pessoais vêm de config_pessoal.json → "planos";
# sem ele, valores genéricos de exemplo. Depois ficam na própria planilha (tbPremissas / tbCarreira).
PLANOS_PADRAO = dict(cdi=0.139, rotulo_renda="Renda a cobrir com os juros", renda=6000.0, aporte=3000.0,
                     aporte_modo="Planejado", meta=15000.0, rotulo_alugueis="Aluguéis / outras rendas (por mês)",
                     alugueis=0.0, carreira=[(None, "Salário hoje", 7000.0), (2030, "Promoção (exemplo)", 8500.0)])


def planos_padrao(exemplo=False):
    pl = dict(PLANOS_PADRAO)
    if not exemplo:
        pl.update({k: v for k, v in (_CFG.get("planos") or {}).items() if v is not None})
    pl["carreira"] = [tuple(x) for x in pl["carreira"]]
    return pl


def investimentos_do_app(cab, linhas, data_pos):
    """Converte as linhas do app para CAB_INV: Saldo hoje do app vira a Base (a posição de onde a linha segue rendendo)."""
    out = []
    for r in linhas:
        d = dict(zip(cab, r))
        d["Base (app)"] = None if str(d.get("Resgatado") or "").lower() == "sim" else d.get("Saldo hoje")
        d["Data da base"] = data_pos if d["Base (app)"] else None
        out.append([d.get(h) if h not in ("Dias", "Alíquota IR", "Saldo hoje", "IR", "Líquido") else None for h in CAB_INV])
    return out


# Colunas da tabela de Favoritos
COLS_ATALHO = ["Descrição", "Tipo", "Categoria", "Pote", "Quem", "Conta", "Forma de pagamento",
               "Valor padrão", "Recorrente?", "Dia do vencimento"]
SITUACOES = ["Pago", "Recebido", "Pendente", "Descontar Depois"]
FORMAS = ["PIX", "Cartão de crédito", "Cartão de débito", "Boleto", "Débito automático",
          "Transferência", "Dinheiro"]


def _mais_meses(ano, mes, n):
    t = ano * 12 + (mes - 1) + n
    return t // 12, t % 12 + 1


def _competencia(d):
    comp = d.get("Competência") or d["Data"]
    return comp.year, comp.month


def _conta_no_mes(d):
    """Mesma regra da fórmula da coluna 'Conta no mês?'."""
    return d["Tipo"] != "Investimento" and (d.get("Situação") or "") in ("", "Pago", "Recebido", "Descontar Depois")


def _orcamento_pela_media(lanc, cat_despesa, ano, mes, meses=12):
    """Meta inicial = média mensal dos últimos `meses` meses (até o mês padrão), arredondada p/ cima em R$ 50."""
    janela = {_mais_meses(ano, mes, -i) for i in range(meses)}
    soma = collections.Counter()
    for d in lanc:
        if d["Tipo"] != "Receita" and _conta_no_mes(d) and _competencia(d) in janela:
            soma[d["Categoria"]] += d["Valor"]
    orc = {}
    for cat, *_ in cat_despesa:
        media = soma.get(cat, 0) / meses
        if media >= 1:
            orc[cat] = int(math.ceil(media / 50.0) * 50)
    return dict(sorted(orc.items(), key=lambda kv: -kv[1]))


def _pessoa_na_descricao(desc, pessoas):
    for p in pessoas:
        if p != "Família" and re.search(rf"\b{re.escape(p)}\b", desc or "", re.I):
            return p
    return None


def _detectar_pessoas(lanc, preencher=True):
    """Pessoas da família: nomes que aparecem em "Salário X" / "Extras X" com frequência.
    Preenche a coluna Quem das linhas que citam o nome na descrição (se ainda estiver vazia)."""
    nomes = collections.Counter(m.group(1) for d in lanc
                                for m in [re.match(r"(?:Sal[aá]rio|Extras?)\s+([A-ZÁ-Ú][a-zá-ú]+)$", d.get("Descrição") or "")] if m)
    pessoas = ["Família"] + [n for n, q in nomes.most_common() if q >= 5]
    if preencher:
        for d in lanc:
            if not d.get("Quem"):
                d["Quem"] = _pessoa_na_descricao(d.get("Descrição"), pessoas)
    return pessoas


def _preencher_pote(lanc, cat_despesa):
    """Gasto sem pote recebe o Fundo padrão da categoria — a mesma regra da coluna Fundo da Base_Dados
    (mapa da aba Categorias). "Sem categoria" continua sem pote: chutar seria inventar."""
    fundo = {c: f for c, f, *_ in cat_despesa if f}
    for d in lanc:
        if not d.get("Pote") and d.get("Tipo") not in ("Receita", "Investimento") and d.get("Categoria") in fundo:
            d["Pote"] = fundo[d["Categoria"]]


def _fundo_padrao(lanc, categoria):
    """Pote (fundo) mais usado pela categoria no histórico — o "Fundo Padrão" da aba Categorias."""
    cont = collections.Counter(d.get("Pote") for d in lanc if d.get("Categoria") == categoria and d.get("Pote"))
    return cont.most_common(1)[0][0] if cont else None


def derivar_atalhos(lanc, ano_p, mes_p, pessoas, min_usos=3, janela=24):
    """Favoritos a partir do histórico: descrições usadas >= min_usos vezes nos últimos 24 meses,
    com o tipo/categoria/pote mais comuns, valor fixo (se os 3 últimos forem iguais) e se é recorrente."""
    grupos = collections.defaultdict(list)
    for d in lanc:
        if d.get("Descrição"):
            grupos[d["Descrição"].strip().lower()].append(d)
    janela = lambda n: {_mais_meses(ano_p, mes_p, -i) for i in range(n)}
    rec6, rec12, rec24 = janela(6), janela(12), janela(24)
    atalhos = []
    for ds in grupos.values():
        if len(ds) < min_usos or (janela and not any(_competencia(d) in rec24 for d in ds)):
            continue
        ds = sorted(ds, key=lambda d: (_competencia(d), d.get("Data") or dt.date.min))
        base = [d for d in ds if _competencia(d) in rec12] or ds

        def comum(campo):
            cont = collections.Counter(d.get(campo) for d in base if d.get(campo) and d.get(campo) != "Sem categoria")
            return cont.most_common(1)[0][0] if cont else None

        nome = collections.Counter(d["Descrição"].strip() for d in ds).most_common(1)[0][0]
        ult3 = [round(d["Valor"], 2) for d in ds[-3:]]
        dias = collections.Counter(d["Vencimento"].day for d in ds if d.get("Vencimento"))
        meses6 = {_competencia(d) for d in ds if _competencia(d) in rec6 and _conta_no_mes(d)}
        recente = any(_competencia(d) in janela(2) for d in ds)
        parcelado = any(d.get("Parcela") for d in ds if _competencia(d) in rec6)
        atalhos.append({
            "Descrição": nome, "Tipo": comum("Tipo"), "Categoria": comum("Categoria") or "Sem categoria",
            "Pote": comum("Pote"), "Quem": _pessoa_na_descricao(nome, pessoas) or comum("Quem"),
            "Conta": comum("Conta"), "Forma de pagamento": comum("Forma de pagamento"),
            "Valor padrão": ult3[-1] if len(ult3) == 3 and len(set(ult3)) == 1 else None,
            "Recorrente?": "sim" if len(meses6) >= 4 and recente and not parcelado else "não",
            "Dia do vencimento": dias.most_common(1)[0][0] if dias else None,
        })
    atalhos.sort(key=lambda a: a["Descrição"].lower())
    return atalhos


# =============================================================================
# DADOS DE EXEMPLO (Abr–Set/2026) — marcados como "EXEMPLO" na Observação
# =============================================================================
def dados_exemplo():
    R, D = "Receita", "Despesa"
    itau, nu, cart = "Conta corrente Itaú", "Conta Nubank", "Carteira"
    cnu, citau = "Cartão Nubank", "Cartão Itaú"
    brutas = []
    variaveis = {
        4: [(18, "Farmácia", D, "Saúde", "Farmácia", cnu, "Cartão de crédito", 142.35),
            (26, "Pizza de sábado", D, "Restaurantes e delivery", "Delivery", cnu, "Cartão de crédito", 118.00)],
        5: [(12, "Projeto freelance — site", R, "Freelance / Extras", "Projeto", nu, "PIX", 1800.00),
            (24, "Presente dia das mães", D, "Compras", "Presentes", citau, "Cartão de crédito", 289.90)],
        6: [(14, "Festa junina da escola", D, "Lazer", "Passeios", cart, "Dinheiro", 95.00),
            (21, "Revisão do carro", D, "Transporte", "Manutenção do carro", citau, "Cartão de crédito", 780.00)],
        7: [(8, "Viagem de férias — hotel", D, "Lazer", "Viagens", citau, "Cartão de crédito", 1650.00),
            (19, "Reembolso plano de saúde", R, "Reembolsos", "Consulta", itau, "Transferência", 320.00)],
        8: [(9, "Projeto freelance — logo", R, "Freelance / Extras", "Projeto", nu, "PIX", 2400.00),
            (22, "Material escolar 2º semestre", D, "Educação", "Material escolar", cnu, "Cartão de crédito", 356.70)],
        9: [(13, "Aniversário — restaurante", D, "Restaurantes e delivery", "Restaurante", citau, "Cartão de crédito", 412.00),
            (20, "Show no fim de semana", D, "Lazer", "Hobbies", cnu, "Cartão de crédito", 540.00)],
    }
    energia = {4: 312.40, 5: 286.15, 6: 274.80, 7: 301.22, 8: 329.64, 9: 347.10}
    mercado = {4: 842.37, 5: 915.60, 6: 788.45, 7: 1012.30, 8: 876.90, 9: 934.18}
    combust = {4: 310.00, 5: 285.50, 6: 342.10, 7: 398.00, 8: 305.40, 9: 362.80}
    lazer_f = {4: 64.00, 5: 180.00, 6: 72.50, 7: 145.00, 8: 96.00, 9: 188.00}
    for m in range(4, 10):
        d = lambda day: dt.date(2026, m, day)
        brutas += [
            (d(5), "Salário — titular", R, "Salário", "Mensal", itau, "Transferência", 7800.00),
            (d(5), "Salário — cônjuge", R, "Salário", "Mensal", nu, "Transferência", 5200.00),
            (d(10), "Aluguel do apartamento", D, "Moradia", "Aluguel/Financiamento", itau, "Boleto", 2350.00),
            (d(10), "Conta de energia", D, "Contas da casa", "Energia", itau, "Débito automático", energia[m]),
            (d(7), "Mercado do mês", D, "Alimentação", "Mercado", cnu, "Cartão de crédito", mercado[m]),
            (d(15), "Combustível", D, "Transporte", "Combustível", citau, "Cartão de crédito", combust[m]),
            (d(8), "Mensalidade escolar", D, "Educação", "Escola", itau, "Boleto", 1450.00),
            (d(3), "Plano de saúde familiar", D, "Saúde", "Plano de saúde", itau, "Débito automático", 890.00),
            (d(12), "Streaming + academia", D, "Assinaturas", "Streaming", cnu, "Cartão de crédito", 174.80),
            (d(27), "Cinema / passeio", D, "Lazer", "Cinema", nu, "PIX", lazer_f[m]),
        ]
        for day, *rest in variaveis[m]:
            brutas.append((d(day), *rest))
    brutas.sort(key=lambda x: x[0])
    lanc = []
    for data, desc, tipo, cat, sub, conta, forma, valor in brutas:
        lanc.append({"Data": data, "Descrição": desc, "Tipo": tipo, "Categoria": cat, "Subcategoria": sub,
                     "Pote": "Essenciais" if tipo == D and cat not in ("Lazer", "Restaurantes e delivery") else ("Lazer" if tipo == D else None),
                     "Conta": conta, "Forma de pagamento": forma, "Valor": valor,
                     "Situação": "Recebido" if tipo == R else "Pago", "Observação": "EXEMPLO"})
    cat_despesa = [
        ("Moradia", "Essenciais", "Aluguel/Financiamento, Condomínio, IPTU, Manutenção"),
        ("Contas da casa", "Essenciais", "Energia, Água, Gás, Internet, Celular"),
        ("Alimentação", "Essenciais", "Mercado, Feira, Padaria, Açougue"),
        ("Restaurantes e delivery", "Lazer", "Restaurante, Delivery, Lanches"),
        ("Transporte", "Essenciais", "Combustível, Estacionamento, App de transporte, Manutenção do carro"),
        ("Educação", "Essenciais", "Escola, Material escolar, Cursos, Livros"),
        ("Saúde", "Essenciais", "Plano de saúde, Farmácia, Consultas, Exames"),
        ("Lazer", "Lazer", "Passeios, Cinema, Viagens, Hobbies"),
        ("Assinaturas", "Essenciais", "Streaming, Academia, Aplicativos, Música"),
        ("Compras", "Essenciais", "Roupas, Casa, Eletrônicos, Presentes"),
        ("Impostos e taxas", "Essenciais", "IPVA, Tarifas bancárias, Anuidades"),
        ("Outros", "Essenciais", "Diversos"),
    ]
    return dict(
        exemplo=True, lanc=lanc, mes_padrao=9, ano_padrao=2026,
        tipos=["Receita", "Despesa", "Investimento"], situacoes=SITUACOES,
        potes=["Essenciais", "Lazer", "Reserva"],
        cat_receita=["Salário", "Freelance / Extras", "Rendimentos", "Reembolsos", "Outras receitas"],
        cat_despesa=cat_despesa, cat_outras=["Investimento"],
        contas=[("Conta corrente Itaú", "Conta corrente", "Itaú"), ("Conta Nubank", "Conta digital", "Nubank"),
                ("Poupança Caixa", "Poupança", "Caixa"), ("Carteira", "Dinheiro", "—")],
        cartoes=[("Cartão Nubank", "Mastercard", 3, 10, 8000), ("Cartão Itaú", "Visa", 25, 5, 12000)],
        formas=FORMAS, anos=list(range(2024, 2033)),
        orcamento={"Moradia": 3100, "Contas da casa": 520, "Alimentação": 1700,
                   "Restaurantes e delivery": 450, "Transporte": 700, "Educação": 1500,
                   "Saúde": 1100, "Lazer": 600, "Assinaturas": 260, "Compras": 500,
                   "Impostos e taxas": 200, "Outros": 300},
        nota_orcamento="Metas de exemplo: ajuste para a sua realidade.",
        metas=[("Reserva de emergência", 36000, 14200, dt.date(2027, 12, 31)),
               ("Viagem de férias (família)", 12000, 4300, dt.date(2027, 1, 31)),
               ("Troca do carro", 45000, 9800, dt.date(2028, 12, 31)),
               ("Notebook novo", 6500, 5200, dt.date(2026, 12, 15))],
        dividas=[("Financiamento do carro", 1180.00, 22, 48, dt.date(2028, 11, 10)),
                 ("Geladeira (cartão Itaú)", 349.90, 7, 10, dt.date(2026, 12, 5)),
                 ("Curso de inglês", 420.00, 3, 12, dt.date(2027, 6, 10))],
        investimentos=[[dt.date(2025, 3, 10), "Banco A", 1.0, 10000, None, None, None, None, None, None, None, dt.date(2028, 3, 10), "EXEMPLO", None, None],
                       [dt.date(2026, 1, 15), "Banco B", 1.1, 6000, None, None, None, None, None, None, None, dt.date(2027, 1, 15), "EXEMPLO", None, None],
                       [dt.date(2026, 6, 5), "Banco A", 0.95, 4000, None, "sim", None, None, None, None, None, dt.date(2027, 6, 5), "EXEMPLO (LCI)", None, None]],
        planos=planos_padrao(exemplo=True), acertos=None,
        subtitulo_lanc="Linhas com \"EXEMPLO\" na Observação são dados fictícios: filtre e apague quando quiser.",
        pessoas=["Família"],
        atalhos=derivar_atalhos(lanc, 2026, 9, ["Família"]),
        potes_saldo=[("Reserva", 0.80, None), ("Lazer", 0.20, None)],
        correcoes=[], conferidos={}, conta_a_partir=dt.date(2026, 4, 1),
    )


# =============================================================================
# IMPORTAÇÃO DOS DADOS REAIS (arquivo exportado pelo app: Financas_Plinio_*.xlsx)
# =============================================================================
def _data(v):
    if isinstance(v, dt.datetime):
        return v.date()
    return v if isinstance(v, dt.date) else None


def _txt(v):
    if v is None:
        return None
    s = str(v).strip()
    return s if s and s != "-" else None


def importar_plinio(caminho):
    wb = load_workbook(caminho, data_only=True)

    # ---- Lançamentos
    ws = wb["Lançamentos"]
    cab = [c.value for c in ws[1]]
    brutas = [dict(zip(cab, r)) for r in ws.iter_rows(min_row=2, values_only=True) if any(v is not None for v in r)]

    # Categoria sugerida para linhas sem categoria: a mais usada para a mesma descrição e tipo no histórico
    hist = collections.defaultdict(collections.Counter)
    for r in brutas:
        if _txt(r["Categoria"]):
            hist[(r["Tipo"], _txt(r["Descrição"]).lower())][_txt(r["Categoria"])] += 1

    lanc, sugeridas, sem_cat = [], 0, 0
    for r in brutas:
        data = _data(r["Data"])
        ano_c, mes_c = (int(x) for x in str(r["Mês"]).split("-"))
        comp = dt.date(ano_c, mes_c, 1)
        obs = []
        cat = _txt(r["Categoria"])
        if not cat:
            sug = hist.get((r["Tipo"], _txt(r["Descrição"]).lower()))
            if sug:
                cat = sug.most_common(1)[0][0]
                obs.append("categoria sugerida pelo histórico")
                sugeridas += 1
            else:
                cat = "Sem categoria"
                sem_cat += 1
        if _txt(r.get("Projetada")) == "sim":
            obs.append("projetada (previsão de parcela)")
        lanc.append({
            "Data": data, "Descrição": _txt(r["Descrição"]), "Tipo": r["Tipo"], "Categoria": cat,
            "Pote": _txt(r["Pote"]), "Parcela": _txt(r["Parcela"]), "Valor": float(r["Valor"]),
            "Situação": _txt(r["Situação"]), "Vencimento": _data(r["Vencimento"]),
            "Descontar de": _txt(r["Descontar de"]),
            # Competência só quando o mês do lançamento é diferente do mês da data
            "Competência": comp if (data.year, data.month) != (ano_c, mes_c) else None,
            "Observação": "; ".join(obs) or None,
        })

    # Mês padrão do Dashboard: último mês com lançamento efetivado (não projetado/pendente)
    efetivos = [_competencia(d) for d in lanc if _conta_no_mes(d)]
    ano_p, mes_p = max(efetivos)

    if PESSOAS:                         # Correção IA #1: lista fixa de pessoas (config_pessoal.json)
        pessoas = list(PESSOAS)
        for d in lanc:
            d["Quem"] = PESSOAS[0]
    else:
        pessoas = _detectar_pessoas(lanc)

    def por_frequencia(itens):
        return [k for k, _ in collections.Counter(itens).most_common()]

    cats_rec = por_frequencia(d["Categoria"] for d in lanc if d["Tipo"] == "Receita")
    cats_out = por_frequencia(d["Categoria"] for d in lanc if d["Tipo"] == "Investimento")
    cats_des = por_frequencia(d["Categoria"] for d in lanc if d["Tipo"] not in ("Receita", "Investimento"))
    for lst in (cats_rec, cats_des):           # "Sem categoria" sempre no fim da lista
        if "Sem categoria" in lst:
            lst.remove("Sem categoria")
            lst.append("Sem categoria")
    exemplos = collections.defaultdict(collections.Counter)
    for d in lanc:
        if d["Tipo"] not in ("Receita", "Investimento"):
            exemplos[d["Categoria"]][d["Descrição"]] += 1
    cat_despesa = [(c, _fundo_padrao(lanc, c), ", ".join(k for k, _ in exemplos[c].most_common(5))) for c in cats_des]
    _preencher_pote(lanc, cat_despesa)

    tipos = por_frequencia(d["Tipo"] for d in lanc)
    tipos = ["Receita"] + [t for t in tipos if t not in ("Receita", "Investimento")] + ["Investimento"]
    potes = por_frequencia(d["Pote"] for d in lanc if d["Pote"])
    anos = list(range(min(d["Data"].year for d in lanc), max(max(y for y, _ in efetivos), dt.date.today().year) + 7))

    # Parcelamentos em andamento (Metas & Dívidas): parcelas k/n do mês padrão com k < n já pagas
    dividas = []
    for d in lanc:
        m = re.match(r"(\d+)/(\d+)$", d["Parcela"] or "")
        if m and _competencia(d) == (ano_p, mes_p) and _conta_no_mes(d):
            k, n = int(m.group(1)), int(m.group(2))
            if k < n:
                fa, fm = _mais_meses(ano_p, mes_p, n - k)
                dividas.append((d["Descrição"], d["Valor"], k, n, dt.date(fa, fm, 1)))

    # ---- Investimentos (posição calculada pelo app na data da exportação)
    investimentos, cab_inv = None, None
    if "Investimentos" in wb.sheetnames:
        wi = wb["Investimentos"]
        cab_inv = [c.value for c in wi[1]]
        investimentos = []
        for r in wi.iter_rows(min_row=2, values_only=True):
            if not isinstance(r[0], (dt.datetime, dt.date)):
                break
            investimentos.append([_data(v) if isinstance(v, (dt.datetime, dt.date)) else v for v in r])

    # ---- Acertos a cobrar
    acertos, rateios, totais_acerto = None, {}, {}
    if "Acertos a cobrar" in wb.sheetnames:
        acertos = []
        for r in wb["Acertos a cobrar"].iter_rows(min_row=2, values_only=True):
            if not any(v is not None for v in r):
                continue
            pessoa, mes, data, desc, parc, cheio, pago, falta = r[:8]
            desc_l = (desc or "").strip().lower()
            if data is not None:
                acertos.append([pessoa, mes, _data(data), (desc or "").strip(), parc, cheio, pago or 0])
            elif desc_l == "total a cobrar":
                totais_acerto[pessoa] = falta
            elif desc_l in ("aluguel dele", "aluguel dela"):
                rateios.setdefault(pessoa, {})[desc_l] = falta

    data_pos = None
    if "Sobre este arquivo" in wb.sheetnames:
        m = re.search(r"(\d{2})/(\d{2})/(\d{4})", str(wb["Sobre este arquivo"]["A1"].value or ""))
        if m:
            data_pos = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))

    bancos = por_frequencia(r[1] for r in investimentos or [] if r[1])
    cartoes = sorted({m.group(1) for d in lanc for m in [re.match(r"Cartão de Crédito (\w+)", d["Descrição"] or "")] if m})

    return dict(
        exemplo=False, lanc=lanc, mes_padrao=mes_p, ano_padrao=ano_p,
        tipos=tipos, situacoes=SITUACOES, potes=potes,
        cat_receita=cats_rec, cat_despesa=cat_despesa, cat_outras=cats_out,
        contas=[(b, "Banco", b) for b in bancos] + [("Carteira", "Dinheiro", "—")],
        cartoes=[(f"Cartão {c}", None, None, None, None) for c in cartoes],
        formas=FORMAS, anos=anos,
        orcamento=_orcamento_pela_media(lanc, cat_despesa, ano_p, mes_p),
        nota_orcamento=(f"Metas iniciais = média mensal dos gastos de cada categoria nos 12 meses até "
                        f"{MESES[mes_p - 1].lower()}/{ano_p}, arredondada para cima (múltiplos de R$ 50). Ajuste à vontade."),
        metas=[], dividas=dividas,
        investimentos=investimentos_do_app(cab_inv, investimentos, data_pos) if investimentos else None,
        cab_investimentos=CAB_INV, data_posicao=data_pos, planos=planos_padrao(),
        acertos=acertos, rateios=rateios, totais_acerto=totais_acerto,
        subtitulo_lanc=(f"Importado do app em {data_pos:%d/%m/%Y}. " if data_pos else "") +
                       "Só entra nos totais o que tem 'Conta no mês?' = sim (receita recebida; gasto pago ou a descontar; nunca investimento, pendente ou projetado).",
        info_importacao=dict(linhas=len(lanc), sugeridas=sugeridas, sem_categoria=sem_cat),
        pessoas=pessoas, atalhos=derivar_atalhos(lanc, ano_p, mes_p, pessoas, min_usos=1, janela=None),
        potes_saldo=list(POTES_PLINIO) or [(p, None, None) for p in potes if p not in ("Corrente", "-")],
        correcoes=[], conferidos={},
        conta_a_partir=dt.date(*CONTA_A_PARTIR_PLINIO, 1) if CONTA_A_PARTIR_PLINIO else dt.date(ano_p, mes_p, 1),
    )


# =============================================================================
# REIMPORTAÇÃO DA PRÓPRIA PLANILHA (para gerar uma versão nova sem perder o que foi lançado)
# =============================================================================
def _ler_tabela(wb_val, nome):
    """Lê uma Tabela (ListObject) pelo nome e devolve (cabeçalhos, linhas como dicts), sem a linha de totais."""
    for ws in wb_val.worksheets:
        if nome in ws.tables:
            t = ws.tables[nome]
            linhas = list(ws[t.ref])
            cab = [c.value for c in linhas[0]]
            corpo = linhas[1:len(linhas) - (t.totalsRowCount or 0)]
            dados = [dict(zip(cab, [c.value for c in r])) for r in corpo]
            return cab, [d for d in dados if any(v not in (None, "") for v in d.values())]
    return None, []


def importar_propria(caminho):
    wb = load_workbook(caminho, data_only=True)
    num = lambda v: float(v) if isinstance(v, (int, float)) else None

    def lista(tab, col=None):
        cab, rows = _ler_tabela(wb, tab)
        return [r[col or cab[0]] for r in rows if r.get(col or cab[0])] if cab else []

    def tuplas(tab):
        cab, rows = _ler_tabela(wb, tab)
        return [tuple(r[c] for c in cab) for r in rows] if cab else []

    pessoas = lista("tbPessoas")
    exemplo_arq = False
    _, at_rows = _ler_tabela(wb, "tbAtalhos")
    atalhos_existentes = {(a.get("Descrição") or "").strip().lower(): a for a in at_rows}

    _, rows = _ler_tabela(wb, "tbLancamentos")
    _, rows_f = _ler_tabela(load_workbook(caminho), "tbLancamentos")      # mesmas linhas, com as fórmulas
    lanc = []
    for r, rf in zip(rows, rows_f):
        desc = _txt(r.get("Descrição"))
        at = atalhos_existentes.get((desc or "").lower(), {})
        d = {h: r.get(h) for h in COLS_LANC if h not in ("Mês", "Ano", "Conta no mês?")}
        for h in ("Data", "Vencimento", "Competência"):
            d[h] = _data(d.get(h))
        for h in d:
            if isinstance(d[h], str):
                d[h] = _txt(d[h])
        # colunas automáticas com fórmula e sem valor em cache (arquivo nunca aberto no Excel): usa o favorito
        for h in COLS_AUTO:
            if d.get(h) in (None, "") and isinstance(rf.get(h), str) and rf[h].startswith("="):
                d[h] = at.get("Valor padrão" if h == "Valor" else h)
        d["_rev"] = isinstance(rf.get("Categoria"), str) and "tbRevisaoCat" in rf["Categoria"]   # linha da aba Categorizar
        d["Descrição"] = desc
        d["Valor"] = num(d.get("Valor"))
        if not d.get("Tipo"):
            d["Tipo"] = "Gasto Extra"
        if not d.get("Situação"):
            d["Situação"] = "Recebido" if d["Tipo"] == "Receita" else "Pago"
        if d["Data"] is None and d["Competência"] is None:
            if not desc or d["Valor"] is None:
                continue
            # linha digitada sem data (Correção IA: data automática) — entra com a data de hoje, marcada
            d["Data"] = dt.date.today()
            d["Observação"] = ((d.get("Observação") or "") + " (data vazia: preenchida com o dia da atualização)").strip()
        lanc.append(d)

    exemplo_arq = any((d.get("Observação") or "") == "EXEMPLO" for d in lanc)
    if PESSOAS and not exemplo_arq:     # Correção IA #1: lista fixa de pessoas (config_pessoal.json)
        pessoas = list(PESSOAS)
    else:
        pessoas = pessoas or _detectar_pessoas(lanc)
    if not exemplo_arq:
        for d in lanc:
            if d.get("Quem") not in pessoas:
                d["Quem"] = pessoas[0]
    # Mês de trabalho: no Caderno (versão nova) ou no Dashboard (versões anteriores)
    mes_sel = ano_sel = None
    for aba in ("Caderno", "Dashboard"):
        if aba in wb.sheetnames and not mes_sel:
            linha3 = [c.value for row in wb[aba].iter_rows(min_row=3, max_row=3) for c in row]
            mes_sel = next((v for v in linha3 if v in MESES), None)
            ano_sel = next((v for v in linha3 if isinstance(v, int) and 2000 < v < 2100), None)
    efetivos = [_competencia(d) for d in lanc if _conta_no_mes(d)]
    ano_p, mes_p = (ano_sel, MESES.index(mes_sel) + 1) if mes_sel and ano_sel else max(efetivos)

    _, orc = _ler_tabela(wb, "tbOrcamento")
    _, metas = _ler_tabela(wb, "tbMetas")
    _, divs = _ler_tabela(wb, "tbDividas")
    cab_inv, inv = _ler_tabela(wb, "tbInvestimentos")
    _, acs = _ler_tabela(wb, "tbAcertos")
    data_pos = None
    if "Investimentos" in wb.sheetnames:
        m = re.search(r"(\d{2})/(\d{2})/(\d{4})", str(wb["Investimentos"]["B3"].value or ""))
        if m:
            data_pos = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    # Investimentos: versão nova (com Base) mantém só o que se digita; versão antiga = posição fixa do app
    investimentos = None
    if cab_inv:
        linhas_inv = [[_data(r[h]) if isinstance(r[h], (dt.datetime, dt.date)) else r[h] for h in cab_inv]
                      for r in inv if r.get("Data") or r.get("Aplicado")]
        if "Base (app)" in cab_inv:
            investimentos = [[None if h in ("Dias", "Alíquota IR", "Saldo hoje", "IR", "Líquido") else
                              dict(zip(cab_inv, r)).get(h) for h in CAB_INV] for r in linhas_inv]
        else:
            investimentos = investimentos_do_app(cab_inv, linhas_inv, data_pos)
    # Plano 1 e 2: premissas e carreira como estão na planilha
    planos = planos_padrao(exemplo_arq)
    _, prem = _ler_tabela(wb, "tbPremissas")
    chaves = ["cdi", "renda", "aporte", "aporte_modo", "meta", "alugueis"]
    rotulos = {"renda": "rotulo_renda", "alugueis": "rotulo_alugueis"}
    for k, pr_ in zip(chaves, prem):
        if pr_.get("Valor") not in (None, ""):
            planos[k] = pr_["Valor"]
        if k in rotulos and pr_.get("Premissa"):
            planos[rotulos[k]] = pr_["Premissa"]
    _, cdis = _ler_tabela(wb, "tbCDI")
    if cdis:
        planos["cdi_hist"] = [(_data(c["A partir de"]), c["CDI ao ano"]) for c in cdis if c.get("A partir de") and c.get("CDI ao ano")]
    _, car = _ler_tabela(wb, "tbCarreira")
    if car:
        planos["carreira"] = [(_data(c["A partir de"]) if c.get("A partir de") else None, c.get("Posto / etapa"), c.get("Salário"))
                              for c in car]
    rateios = {}
    if "Acertos" in wb.sheetnames:
        wa = wb["Acertos"]
        atual = {}                     # coluna → pessoa do bloco de rateio (o bloco pode estar em qualquer coluna)
        for row in wa.iter_rows():
            for i, c in enumerate(row[:-1]):
                rot = c.value
                if isinstance(rot, str) and rot.startswith("RATEIO — "):
                    atual[i] = next((a["Pessoa"] for a in acs if str(a["Pessoa"]).upper() == rot[9:]), rot[9:].title())
                elif i in atual and rot in ("Aluguel dele (digite)", "Aluguel dela (digite)"):
                    rateios.setdefault(atual[i], {})["aluguel dele" if "dele" in rot else "aluguel dela"] = row[i + 1].value
    _, recs = _ler_tabela(wb, "tbRecebidos")

    # Favoritos: os que já existem (com as suas edições) + toda descrição do histórico que ainda não
    # está no mapa — o mapa completo Descrição → Categoria, como a aba Categorias da Finanças Família 3.0
    atalhos = []
    for a in at_rows:
        a = {h: a.get(h) for h in COLS_ATALHO}
        if a.get("Quem") and a["Quem"] not in pessoas:
            a["Quem"] = None
        atalhos.append(a)
    ja = {(a.get("Descrição") or "").strip().lower() for a in atalhos}
    atalhos += [a for a in derivar_atalhos(lanc, ano_p, mes_p, pessoas, min_usos=1, janela=None)
                if a["Descrição"].strip().lower() not in ja]
    atalhos.sort(key=lambda a: (a.get("Descrição") or "").lower())

    # Categorias de gasto com o Fundo padrão (lido por nome de coluna; derivado do histórico se faltar)
    cab_cd, rows_cd = _ler_tabela(wb, "tbCatDespesa")
    cat_despesa = []
    for r in rows_cd:
        nome_cat = r.get(cab_cd[0])
        if nome_cat:
            cat_despesa.append((nome_cat, r.get("Fundo padrão") or _fundo_padrao(lanc, nome_cat),
                                r.get("Subcategorias / exemplos")))

    if not exemplo_arq:
        _preencher_pote(lanc, cat_despesa)

    # Potes: aba Potes (versão nova) ou os valores do Plínio
    _, pts = _ler_tabela(wb, "tbPotesSaldo")
    potes_saldo = [(p["Pote"], p.get("% alvo"), p.get("Guardado")) for p in pts if p.get("Pote")]
    if not potes_saldo:
        potes_saldo = (list(POTES_PLINIO) or [(p, None, None) for p in lista("tbPotes") if p not in ("Corrente", "-")]) \
            if not exemplo_arq else [("Reserva", 0.80, None), ("Lazer", 0.20, None)]

    # Correções IA / Melhorias IA: seus pedidos (tabela tbCorrecoes ou a aba solta que você criou)
    correcoes = []
    _, crs = _ler_tabela(wb, "tbCorrecoes")
    if crs:
        correcoes = [(c.get("Pedido"), c.get("Status"), c.get("Como ficou")) for c in crs if c.get("Pedido")]
    else:
        for aba in wb.sheetnames:
            if re.search(r"(corre[cç][oõ]es|melhorias)\s*ia", aba, re.I):
                for row in wb[aba].iter_rows(values_only=True):
                    ped = _txt(row[0]) if row else None
                    if ped and not re.match(r"(🤖|pedido$|dica)", ped, re.I):
                        correcoes.append((ped, _txt(row[1]) if len(row) > 1 else None, _txt(row[2]) if len(row) > 2 else None))

    # Meses conferidos (coluna "Conferido em" do Histórico) e o "Conta a partir de" do Caderno
    conferidos, conta_a_partir = {}, None
    if "Historico" in wb.sheetnames:
        wh = wb["Historico"]
        cab_h = [c.value for c in wh[5]]
        if "Conferido em" in cab_h:
            jc = cab_h.index("Conferido em")
            for row in wh.iter_rows(min_row=6, values_only=True):
                if isinstance(row[1], (dt.datetime, dt.date)) and row[jc] not in (None, ""):
                    conferidos[(row[1].year, row[1].month)] = row[jc]
    if "Caderno" in wb.sheetnames:
        for row in wb["Caderno"].iter_rows(values_only=True):
            for i, v in enumerate(row[:-1]):
                if v == "Conta a partir de" and isinstance(row[i + 1], (dt.datetime, dt.date)):
                    conta_a_partir = _data(row[i + 1])
    if not conta_a_partir:
        conta_a_partir = dt.date(*CONTA_A_PARTIR_PLINIO, 1) if (CONTA_A_PARTIR_PLINIO and not exemplo_arq) else min(
            dt.date(*_competencia(d), 1) for d in lanc)

    return dict(
        exemplo=exemplo_arq, lanc=lanc,
        mes_padrao=mes_p, ano_padrao=ano_p,
        tipos=lista("tbTipos"), situacoes=lista("tbSituacoes") or SITUACOES, potes=lista("tbPotes"),
        cat_receita=lista("tbCatReceita"), cat_despesa=cat_despesa,
        cat_outras=lista("tbCatOutras"), contas=tuplas("tbContas"),
        cartoes=[t for t in tuplas("tbCartoes") if t[0]], formas=lista("tbFormas") or FORMAS,
        anos=lista("tbAnos"),
        orcamento={o["Categoria"]: o["Meta mensal"] for o in orc if o.get("Categoria")},
        nota_orcamento="Metas definidas por você (mantidas da versão anterior da planilha).",
        metas=[(m["Nome"], m["Valor alvo"], m["Valor atual"], _data(m["Prazo"])) for m in metas if m.get("Nome")],
        dividas=[(v["Descrição"], v["Valor da parcela"], v["Parcelas pagas"], v["Total de parcelas"], _data(v["Data de término"]))
                 for v in divs if v.get("Descrição")],
        investimentos=investimentos, cab_investimentos=CAB_INV, data_posicao=data_pos, planos=planos,
        acertos=[[a["Pessoa"], a["Mês"], _data(a["Data"]), a["Descrição"], a["Parcela"], a["Valor cheio"], a["Já pago"] or 0]
                 for a in acs if "Valor cheio" in a] or None,               # formato antigo (lista do app)
        acertos_pessoas=[a["Pessoa"] for a in acs if "Valor cheio" not in a and a.get("Pessoa")],
        recebidos=[[_data(r.get("Data")), r.get("Pessoa"), r.get("Valor"), _data(r.get("Referente a (mês)"))]
                   for r in recs if r.get("Valor")],
        rateios=rateios,
        subtitulo_lanc="Só entra nos totais o que tem 'Conta no mês?' = sim (receita recebida; gasto pago ou a descontar; nunca investimento, pendente ou projetado).",
        info_importacao=dict(linhas=len(lanc), sugeridas=0, sem_categoria=sum(1 for d in lanc if d.get("Categoria") == "Sem categoria")),
        pessoas=pessoas, atalhos=atalhos, potes_saldo=potes_saldo, correcoes=correcoes,
        conferidos=conferidos, conta_a_partir=conta_a_partir,
    )


def importar(caminho):
    """Detecta o formato: planilha gerada por este script (tem tbLancamentos) ou exportação do app."""
    wb = load_workbook(caminho, read_only=False)
    propria = any("tbLancamentos" in ws.tables for ws in wb.worksheets)
    D = importar_propria(caminho) if propria else importar_plinio(caminho)
    _, resp = _ler_tabela(load_workbook(caminho, data_only=True), "tbRevisaoCat")
    preparar_revisao(D, {(r.get("Descrição") or "").strip().lower(): r for r in resp if r.get("Descrição")})
    preparar_acertos(D)
    return D


def preparar_acertos(D):
    """Pessoas da aba Acertos = nomes usados em "Descontar de" (+ os já cadastrados). A lista antiga do app
    (pessoa do acerto ≠ nome em Descontar de, ex.: o rateio) é casada com os lançamentos para achar o nome."""
    lanc = D["lanc"]
    nomes = collections.Counter(d["Descontar de"] for d in lanc if d.get("Descontar de"))
    mapa = {}
    for a in D.get("acertos") or []:
        pessoa, desc_a, valor = a[0], (a[3] or "").strip().lower(), a[5]
        achados = collections.Counter(d.get("Descontar de") for d in lanc if d.get("Situação") == "Descontar Depois"
                                      and (d.get("Descrição") or "").strip().lower() == desc_a
                                      and d.get("Descontar de") and abs((d.get("Valor") or 0) - (valor or 0)) < 0.01)
        if achados:
            mapa.setdefault(pessoa, collections.Counter()).update(achados)
    mapa = {k: v.most_common(1)[0][0] for k, v in mapa.items()}
    D["rateios"] = {mapa.get(k, k): v for k, v in (D.get("rateios") or {}).items()}
    pessoas = list(D.get("acertos_pessoas") or [])
    for n in [n for n, _ in nomes.most_common()] + list(D["rateios"]):
        if n and n.lower() not in {p.lower() for p in pessoas}:
            pessoas.append(n)
    D["acertos_pessoas"] = pessoas
    D["recebidos"] = list(D.get("recebidos") or []) + [
        [None, mapa.get(a[0], a[0]), a[6], dt.date(*map(int, str(a[1]).split("-")[:2]), 1) if a[1] else None]
        for a in D.get("acertos") or [] if a[6]]


# =============================================================================
# CATEGORIZAR (Correção IA: "não deixar nada sem categoria; se tiver Outros, pergunte o que colocar")
# =============================================================================
PENDENTES_CAT = ("Outros", "Sem categoria", "", None)


def _sugerir_categoria(desc, tipo, lanc, cat_despesa):
    """Sugestão genérica: 1) a categoria que a mesma descrição já teve; 2) os exemplos de cada categoria
    (aba Config); 3) receita → Renda."""
    k = desc.strip().lower()
    usadas = collections.Counter(d["Categoria"] for d in lanc
                                 if (d.get("Descrição") or "").strip().lower() == k and d.get("Categoria") not in PENDENTES_CAT)
    if usadas:
        return usadas.most_common(1)[0][0]
    if tipo == "Receita":
        return "Renda"
    palavras = set(re.findall(r"\w{4,}", k))
    melhor = None
    for cat, _fundo, ex in cat_despesa:
        if cat in PENDENTES_CAT:
            continue
        exemplos = [e.strip().lower() for e in str(ex or "").split(",") if e.strip()] + [cat.lower()]
        for e in exemplos:
            if e == k or (len(e) >= 4 and e in k) or (palavras & set(re.findall(r"\w{4,}", e))):
                return cat
    return melhor


def preparar_revisao(D, respostas):
    """Aplica as respostas já dadas na aba Categorizar e monta a lista do que ainda está em Outros/Sem categoria.
    Respostas da versão anterior viram categoria fixa; sugestões novas ficam na aba para você conferir."""
    lanc, cat_despesa = D["lanc"], D["cat_despesa"]
    fundo = {c: f for c, f, _ in cat_despesa}
    for d in lanc:
        r = respostas.get((d.get("Descrição") or "").strip().lower())
        if r and r.get("Categoria certa") and (d.get("_rev") or d.get("Categoria") in PENDENTES_CAT + (r.get("Categoria atual"),)):
            d["Categoria"] = r["Categoria certa"]
            if d.get("Tipo") != "Receita" and fundo.get(d["Categoria"]):
                d["Pote"] = fundo[d["Categoria"]]
        elif d.get("Categoria") in ("", None) and r:
            d["Categoria"] = r.get("Categoria atual")
        if r and not d.get("Pote") and d.get("Tipo") != "Receita":      # arquivo nunca aberto no Excel
            d["Pote"] = fundo.get(d.get("Categoria"))
    for a in D.get("atalhos") or []:
        r = respostas.get((a.get("Descrição") or "").strip().lower())
        if r and r.get("Categoria certa"):
            a["Categoria"] = r["Categoria certa"]
            if a.get("Tipo") != "Receita" and fundo.get(a["Categoria"]):
                a["Pote"] = None            # o Pote segue o Fundo padrão da categoria
        elif r and not a.get("Categoria"):
            a["Categoria"] = r.get("Categoria atual")
    extra = {k.strip().lower(): v for k, v in ((_CFG.get("revisao_categoria") or {}) if not D.get("exemplo") else {}).items()}
    grupos = collections.OrderedDict()
    for d in sorted(lanc, key=lambda x: (x.get("Descrição") or "").lower()):
        desc = d.get("Descrição")
        if not desc or d.get("Categoria") not in PENDENTES_CAT or d.get("Tipo") == "Investimento":
            continue
        g = grupos.setdefault(desc.strip().lower(), dict(desc=desc.strip(), tipo=d.get("Tipo"), vezes=0, total=0.0,
                                                         ultimo=None, atual=d.get("Categoria") or "Sem categoria"))
        g["vezes"] += 1
        g["total"] += d.get("Valor") or 0
        dd = d.get("Data") or (d.get("Competência"))
        if dd and (g["ultimo"] is None or dd > g["ultimo"]):
            g["ultimo"] = dd
    revisao = []
    for k, g in grupos.items():
        cfg = extra.get(k) or {}
        g["sugestao"] = cfg.get("sugestao") or _sugerir_categoria(g["desc"], g["tipo"], lanc, cat_despesa)
        g["certa"] = cfg.get("certa")
        g["pergunta"] = cfg.get("pergunta") or ("" if g["certa"] else
                                                ("Confirme a sugestão ou escolha outra." if g["sugestao"] else "Do que se trata?"))
        revisao.append(g)
    D["revisao"] = revisao


# =============================================================================
# HELPERS DE ESTILO
# =============================================================================
def fill(hex_):
    return PatternFill("solid", start_color="FF" + hex_, end_color="FF" + hex_)


def fnt(size=10, bold=False, color=None, italic=False):
    return Font(name=FONTE, size=size, bold=bold, italic=italic, color="FF" + (color or P["texto"]))


def side(color=None, style="thin"):
    return Side(style=style, color="FF" + (color or P["borda"]))


BG = fill(P["fundo"])
CARD = fill(P["card"])
CARD2 = fill(P["card2"])
HEAD = fill(P["cab_fundo"])
ALIGN_L = Alignment(horizontal="left", vertical="center", indent=1)
ALIGN_R = Alignment(horizontal="right", vertical="center", indent=1)
ALIGN_C = Alignment(horizontal="center", vertical="center")


def pintar_fundo(ws, ultima_col=60, ultima_linha=0):
    """Fundo da aba: estilo de coluna (vale para linhas infinitas) + células já criadas."""
    for c in range(1, ultima_col + 1):
        cd = ws.column_dimensions[get_column_letter(c)]
        cd.fill = BG
        cd.font = fnt()
    for r in range(1, ultima_linha + 1):
        for c in range(1, ultima_col + 1):
            cell = ws.cell(r, c)
            cell.fill = BG
            cell.font = fnt()
    ws.sheet_view.showGridLines = False


def larguras(ws, mapa):
    for col, w in mapa.items():
        ws.column_dimensions[col].width = w


def put(ws, ref, value=None, font=None, fill_=None, fmt=None, align=None, border=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    if font is not None:
        c.font = font
    if fill_ is not None:
        c.fill = fill_
    if fmt is not None:
        c.number_format = fmt
    if align is not None:
        c.alignment = align
    if border is not None:
        c.border = border
    return c


def area(ws, rng, fill_=None, font=None, border=None, fmt=None, align=None):
    for row in ws[rng]:
        for c in row:
            if fill_ is not None:
                c.fill = fill_
            if font is not None:
                c.font = font
            if border is not None:
                c.border = border
            if fmt is not None:
                c.number_format = fmt
            if align is not None:
                c.alignment = align


def mesclar(ws, rng, value=None, **kw):
    area(ws, rng, fill_=kw.get("fill_"), font=kw.get("font"), fmt=kw.get("fmt"),
         align=kw.get("align"), border=kw.get("border"))
    if rng.split(":")[0] != rng.split(":")[-1]:
        ws.merge_cells(rng)
    first = rng.split(":")[0]
    if value is not None:
        ws[first].value = value
    return ws[first]


def titulo_aba(ws, titulo, subtitulo, ultima_col):
    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 18
    put(ws, "B2", titulo, font=fnt(18, True), align=Alignment(vertical="center"))
    put(ws, "B3", subtitulo, font=fnt(9, color=P["suave"]), align=Alignment(vertical="center"))


def nova_tabela(ws, nome, ref, cabecalhos, formulas=None, totais=None):
    """Cria uma Tabela (ListObject). formulas = {coluna: fórmula calculada},
    totais = {coluna: ('sum' | rótulo texto)}."""
    formulas = formulas or {}
    t = Table(displayName=nome, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=False,
                                      showColumnStripes=False, showFirstColumn=False,
                                      showLastColumn=False)
    cols = []
    for i, h in enumerate(cabecalhos, start=1):
        tc = TableColumn(id=i, name=h)
        if h in formulas:
            tc.calculatedColumnFormula = TableFormula(attr_text=formulas[h].lstrip("="))
        if totais and h in totais:
            if totais[h] == "sum":
                tc.totalsRowFunction = "sum"
            else:
                tc.totalsRowLabel = totais[h]
        cols.append(tc)
    t.tableColumns = cols
    if totais:
        t.totalsRowCount = 1
        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
        t.autoFilter = AutoFilter(ref=f"{m.group(1)}{m.group(2)}:{m.group(3)}{int(m.group(4)) - 1}")
    else:
        t.autoFilter = AutoFilter(ref=ref)
    ws.add_table(t)
    return t


def estilo_cabecalho(ws, linha, col_ini, col_fim):
    ws.row_dimensions[linha].height = 24
    for c in range(col_ini, col_fim + 1):
        cell = ws.cell(linha, c)
        cell.fill = HEAD
        cell.font = fnt(10, True, P["cabecalho_txt"])
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
        cell.border = Border()


def estilo_corpo(ws, lin_ini, lin_fim, col_ini, col_fim, fill_=None):
    b = Border(bottom=side(P["borda"]))
    for r in range(lin_ini, lin_fim + 1):
        ws.row_dimensions[r].height = 20
        for c in range(col_ini, col_fim + 1):
            cell = ws.cell(r, c)
            cell.fill = fill_ or CARD
            cell.font = fnt(10)
            cell.border = b
            cell.alignment = ALIGN_L


def dv_lista(ws, formula, rng, titulo="Valor inválido", msg="Escolha um item da lista."):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
    dv.error, dv.errorTitle = msg, titulo
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(rng)
    return dv


def caixa_ajuda(ws, col_a, col_b, lin, linhas, titulo="COMO FUNCIONA"):
    """Caixa de explicação curta ("como funciona") mesclada entre col_a e col_b a partir da linha lin."""
    larg = sum((ws.column_dimensions[get_column_letter(c)].width or 9)
               for c in range(ws[col_a + "1"].column, ws[col_b + "1"].column + 1))
    cabe = max(20, int(larg * 1.05))
    mesclar(ws, f"{col_a}{lin}:{col_b}{lin}", titulo, fill_=CARD2, font=fnt(10, True, P["destaque"]), align=ALIGN_L,
            border=Border(top=side(P["destaque"], "medium")))
    ws.row_dimensions[lin].height = 22
    for i, t in enumerate(linhas, start=1):
        r = lin + i
        mesclar(ws, f"{col_a}{r}:{col_b}{r}", t, fill_=CARD2, font=fnt(9.5, color=P["texto"]),
                align=Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1))
        ws.row_dimensions[r].height = 15 * max(1, math.ceil(len(t) / cabe)) + 4
    return lin + len(linhas) + 1


def regras_percentual(ws, rng, primeira):
    """Barra de dados + cor: verde até 80%, amarelo 80–100%, vermelho acima de 100%."""
    ws.conditional_formatting.add(rng, DataBarRule(start_type="num", start_value=0, end_type="num",
                                                   end_value=1.2, color="FF" + BARRA,
                                                   showValue=True))
    verde = DifferentialStyle(font=Font(color="FF" + P["positivo"], bold=True))
    amar = DifferentialStyle(font=Font(color="FF" + P["alerta"], bold=True))
    verm = DifferentialStyle(font=Font(color="FF" + P["negativo"], bold=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({primeira}),{primeira}>1)"], font=verm.font, stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({primeira}),{primeira}>0.8,{primeira}<=1)"], font=amar.font, stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f"AND(ISNUMBER({primeira}),{primeira}<=0.8)"], font=verde.font, stopIfTrue=False))


# ---- helpers de gráficos --------------------------------------------------------
def _cp(color, size=900, bold=False):
    return CharacterProperties(sz=size, b=bold, solidFill=color, latin=DFont(typeface=FONTE))


def txpr(color, size=900, bold=False):
    cp = _cp(color, size, bold)
    return RichText(bodyPr=RichTextProperties(), p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def titulo_grafico(texto):
    cp = _cp(P["texto"], 1200, True)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=texto)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def estilizar_grafico(ch, titulo):
    ch.title = titulo_grafico(titulo)
    ch.graphical_properties = GraphicalProperties(solidFill=P["card"], ln=LineProperties(noFill=True))
    ch.roundedCorners = False
    ch.plot_area.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
    if ch.legend is not None:
        ch.legend.position = "b"
        ch.legend.txPr = txpr(P["texto"], 900)


def estilizar_eixos(ch, fmt_valor='"R$" #,##0'):
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
        ax.txPr = txpr(P["suave"], 850)
        ax.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=P["borda"]))
        ax.majorGridlines = None
    ch.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=P["grade"], w=6350)))
    ch.y_axis.numFmt = NumFmt(formatCode=fmt_valor, sourceLinked=False)
    ch.y_axis.majorTickMark = "none"
    ch.x_axis.majorTickMark = "none"


def ancorar(ws, ch, col_ini, lin_ini, col_fim, lin_fim):
    """Ancora o gráfico exatamente entre células (índices base 0)."""
    ch.anchor = TwoCellAnchor(_from=AnchorMarker(col=col_ini, row=lin_ini),
                              to=AnchorMarker(col=col_fim, row=lin_fim))
    ws.add_chart(ch)


# =============================================================================
# CONSTRUÇÃO
# =============================================================================
def construir(D, caminho):
    wb = Workbook()
    ws_ini = wb.active
    ws_ini.title = "Inicio"
    ws_corr = wb.create_sheet("Correções IA")      # 2ª aba: onde você deixou
    ws_cad = wb.create_sheet("Caderno")
    ws_lanc = wb.create_sheet("Lancamentos")
    ws_rev = wb.create_sheet("Categorizar")
    ws_contas = wb.create_sheet("ContasMes")
    ws_parc = wb.create_sheet("Parcelar")
    ws_orc = wb.create_sheet("Orcamento")
    ws_potes = wb.create_sheet("Potes")
    ws_anual = wb.create_sheet("Anual")
    ws_hist = wb.create_sheet("Historico")
    ws_p1 = wb.create_sheet("Planos")      # Plano 1 e Plano 2 na mesma aba (Correção IA)
    ws_p2 = None
    ws_inv = wb.create_sheet("Investimentos")
    ws_ac = wb.create_sheet("Acertos")
    ws_at = wb.create_sheet("Atalhos")
    ws_cfg = wb.create_sheet("Config")
    ws_calc = wb.create_sheet("Calc")
    L = "tbLancamentos"
    CL = get_column_letter
    CONTA = f'{L}[Conta no mês?],"sim"'
    rev_keys = {g["desc"].lower() for g in D.get("revisao") or []}

    def f_rev(desc_ref, orig):
        """Categoria que vem da resposta na aba Categorizar (sem resposta: a de antes)."""
        o = (orig or "Sem categoria").replace('"', '""')
        busca = f"INDEX(tbRevisaoCat[Categoria certa],MATCH({desc_ref},tbRevisaoCat[Descrição],0))"
        return f'=IFERROR(IF({busca}&""="","{o}",{busca}&""),"{o}")'

    # ---------------------------------------------------------------- Config
    ws = ws_cfg
    pintar_fundo(ws, 45, 3)
    titulo_aba(ws, "Configurações", "Listas editáveis. Para adicionar um item, digite na primeira linha vazia logo abaixo da tabela: ela cresce sozinha e os menus suspensos acompanham.", 26)
    ws.column_dimensions["A"].width = 2
    L0 = 5  # linha de cabeçalho das tabelas de config
    cfg = [
        ("tbCatReceita", "RECEITAS", ["Categoria de receita"], [[c] for c in D["cat_receita"]], [24]),
        ("tbCatDespesa", "GASTOS", ["Categoria de gasto", "Fundo padrão", "Subcategorias / exemplos"], D["cat_despesa"], [26, 14, 56]),
        ("tbCatOutras", "INVESTIMENTOS / OUTRAS", ["Categoria (fora dos gastos)"], [[c] for c in D["cat_outras"]], [26]),
        ("tbPotes", "POTES", ["Pote"], [[p] for p in D["potes"]], [16]),
        ("tbPessoas", "FAMÍLIA (QUEM)", ["Quem"], [[p] for p in D.get("pessoas") or ["Família"]], [14]),
        ("tbContas", "CONTAS / BANCOS", ["Conta", "Tipo de conta", "Banco"], D["contas"], [20, 15, 12]),
        ("tbCartoes", "CARTÕES", ["Cartão", "Bandeira", "Dia fechamento", "Dia vencimento", "Limite"], D["cartoes"], [18, 12, 11, 11, 14]),
        ("tbFormas", "FORMAS DE PAGAMENTO", ["Forma de pagamento"], [[f] for f in D["formas"]], [20]),
        ("tbTipos", "TIPOS", ["Tipo"], [[t] for t in D["tipos"]], [14]),
        ("tbSituacoes", "SITUAÇÕES", ["Situação"], [[s] for s in D["situacoes"]], [18]),
        ("tbMeses", "MESES", ["Nº", "Mês", "Abrev."], [(i + 1, MESES[i], MESES_ABREV[i]) for i in range(12)], [6, 12, 8]),
        ("tbAnos", "ANOS", ["Ano"], [[a] for a in D["anos"]], [9]),
    ]
    col_cfg = {}
    col = 2
    for nome, secao, cabec, linhas, largs in cfg:
        linhas = list(linhas) or [[None] * len(cabec)]
        col_cfg[nome] = col
        put(ws, f"{CL(col)}{L0 - 1}", secao, font=fnt(10, True, P["suave"]))
        for j, h in enumerate(cabec):
            ws.cell(L0, col + j, h)
            ws.column_dimensions[CL(col + j)].width = largs[j]
        for i, lin in enumerate(linhas):
            for j, v in enumerate(lin):
                ws.cell(L0 + 1 + i, col + j, v)
        cf = col + len(cabec) - 1
        estilo_cabecalho(ws, L0, col, cf)
        estilo_corpo(ws, L0 + 1, L0 + len(linhas), col, cf)
        nova_tabela(ws, nome, f"{CL(col)}{L0}:{CL(cf)}{L0 + len(linhas)}", cabec)
        if nome == "tbCartoes":
            for r in range(L0 + 1, L0 + 1 + len(linhas)):
                ws.cell(r, cf).number_format = FMT_MOEDA
        ws.column_dimensions[CL(cf + 1)].width = 3
        col = cf + 2
    ws.freeze_panes = f"A{L0 + 1}"

    def rng_cfg(tab, desloc=0):
        c = CL(col_cfg[tab] + desloc)
        return f"Config!${c}${L0 + 1}:${c}$1000"

    def nome_dinamico(nome, tab):
        c = CL(col_cfg[tab])
        wb.defined_names[nome] = DefinedName(
            nome, attr_text=f"OFFSET(Config!${c}${L0 + 1},0,0,MAX(1,SUMPRODUCT(--(LEN({rng_cfg(tab)})>0))),1)")

    for nome, tab in [("CatReceita", "tbCatReceita"), ("CatDespesa", "tbCatDespesa"), ("CatOutras", "tbCatOutras"),
                      ("Potes", "tbPotes"), ("Pessoas", "tbPessoas"), ("ListaContas", "tbContas"), ("ListaCartoes", "tbCartoes"),
                      ("FormasPagamento", "tbFormas"), ("Tipos", "tbTipos"), ("Situacoes", "tbSituacoes"),
                      ("Anos", "tbAnos")]:
        nome_dinamico(nome, tab)
    cp = CL(col_cfg["tbPessoas"])
    wb.defined_names["QuemPadrao"] = DefinedName("QuemPadrao", attr_text=f"Config!${cp}${L0 + 1}")
    cfd = CL(col_cfg["tbCatDespesa"] + 1)
    dv_lista(ws, "=Potes", f"{cfd}{L0 + 1}:{cfd}{L0 + 300}", msg="Escolha um pote cadastrado (tabela Potes).")
    cm = col_cfg["tbMeses"]
    wb.defined_names["Meses"] = DefinedName("Meses", attr_text=f"Config!${CL(cm + 1)}${L0 + 1}:${CL(cm + 1)}${L0 + 12}")
    wb.defined_names["MesesAbrev"] = DefinedName("MesesAbrev", attr_text=f"Config!${CL(cm + 2)}${L0 + 1}:${CL(cm + 2)}${L0 + 12}")

    # ---------------------------------------------------------------- Lançamentos
    ws = ws_lanc
    lanc = D["lanc"]
    n = len(lanc)
    H = 4
    ult = H + n
    ncol = len(COLS_LANC)
    ci = {h: 2 + i for i, h in enumerate(COLS_LANC)}          # índice da coluna por nome
    cl = {h: CL(c) for h, c in ci.items()}                     # letra da coluna por nome
    pintar_fundo(ws, ncol + 3, 3)
    titulo_aba(ws, "Lançamentos", "Digite Data, Descrição e Valor: com um favorito, Tipo, Categoria, Situação e Pote se preenchem sozinhos. "
               "Clique no [+] acima das colunas para ver os detalhes (conta, parcela, vencimento...). " + D["subtitulo_lanc"], ncol)
    ws.column_dimensions["A"].width = 2
    for h, w in zip(COLS_LANC, LARG_LANC):
        ws.column_dimensions[cl[h]].width = w
    tr = lambda c: f"{L}[[#This Row],[{c}]]"
    desc = tr("Descrição")
    casa = f"MATCH({desc},tbAtalhos[Descrição],0)"
    tipo_padrao = next((t for t in D["tipos"] if "extra" in t.lower()), D["tipos"][1] if len(D["tipos"]) > 1 else "Receita")

    def f_auto(col):
        return f'=IF({desc}="","",IFERROR(INDEX(tbAtalhos[{col}],{casa})&"",""))'

    f_mes = f'=IF(AND({tr("Data")}="",{tr("Competência")}=""),"",MONTH(IF({tr("Competência")}<>"",{tr("Competência")},{tr("Data")})))'
    f_ano = f'=IF(AND({tr("Data")}="",{tr("Competência")}=""),"",YEAR(IF({tr("Competência")}<>"",{tr("Competência")},{tr("Data")})))'
    f_conta = (f'=IF({tr("Valor")}="","",IF(AND({tr("Tipo")}<>"Investimento",OR({tr("Situação")}="",{tr("Situação")}="Pago",'
               f'{tr("Situação")}="Recebido",{tr("Situação")}="Descontar Depois")),"sim","não"))')
    formulas_lanc = {
        "Mês": f_mes, "Ano": f_ano, "Conta no mês?": f_conta,
        # colunas automáticas: vêm do favorito (tbAtalhos); digitar por cima vale só para aquela linha
        "Valor": f'=IF({desc}="","",IFERROR(1/(1/INDEX(tbAtalhos[Valor padrão],{casa})),""))',
        "Tipo": f'=IF({desc}="","",IFERROR(INDEX(tbAtalhos[Tipo],{casa})&"","{tipo_padrao}"))',
        "Categoria": f_auto("Categoria"),
        # Pote: o do favorito; se não houver, o "Fundo padrão" da categoria (aba Config)
        "Pote": (f'=IF({desc}="","",IF(IFERROR(INDEX(tbAtalhos[Pote],{casa})&"","")<>"",INDEX(tbAtalhos[Pote],{casa})&"",'
                 f'IFERROR(INDEX(tbCatDespesa[Fundo padrão],MATCH({tr("Categoria")},tbCatDespesa[Categoria de gasto],0))&"","")))'),
        # Quem: o do favorito ou, sem ele, a primeira pessoa da lista (Correção IA #1)
        "Quem": f'=IF({desc}="","",IF(IFERROR(INDEX(tbAtalhos[Quem],{casa})&"","")<>"",INDEX(tbAtalhos[Quem],{casa})&"",QuemPadrao))',
        "Conta": f_auto("Conta"), "Forma de pagamento": f_auto("Forma de pagamento"),
        "Situação": f'=IF({desc}="","",IF({tr("Tipo")}="Receita","Recebido","Pago"))',
    }
    so_linhas_novas = set(COLS_AUTO)          # nas linhas importadas essas colunas ficam com o valor fixo
    for h in COLS_LANC:
        ws.cell(H, ci[h], h)
    estilo_cabecalho(ws, H, 2, ncol + 1)
    for h in COLS_AUTO:                        # cabeçalho das colunas automáticas com um tom diferente
        ws.cell(H, ci[h]).fill = fill(misturar(P["roxo"], P["card"], 0.55 if TEMA == "plinio_escuro" else 0.15))
    estilo_corpo(ws, H + 1, ult, 2, ncol + 1)
    fmts = {"Data": FMT_DATA, "Vencimento": FMT_DATA, "Competência": FMT_COMPETENCIA, "Valor": FMT_MOEDA}
    suaves = ("Mês", "Ano", "Conta no mês?", "Parcela", "Competência")
    for i, d in enumerate(lanc):
        r = H + 1 + i
        for h in COLS_LANC:
            c = ws.cell(r, ci[h])
            if h in formulas_lanc and h not in so_linhas_novas:
                c.value = formulas_lanc[h]
            else:
                c.value = d.get(h)
            if h in fmts:
                c.number_format = fmts[h]
            if h in suaves:
                c.alignment = ALIGN_C
                c.font = fnt(10, color=P["suave"])
        ws.cell(r, ci["Valor"]).alignment = ALIGN_R
        ws.cell(r, ci["Observação"]).font = fnt(9, color=P["suave"], italic=True)
        if ((d.get("Descrição") or "").strip().lower() in rev_keys and d.get("Tipo") != "Investimento"
                and d.get("Categoria") in PENDENTES_CAT):
            ws.cell(r, ci["Categoria"]).value = f_rev(desc, d.get("Categoria"))
            if d.get("Tipo") != "Receita":
                po = (d.get("Pote") or "").replace('"', '""')
                ws.cell(r, ci["Pote"]).value = (f'=IFERROR(INDEX(tbCatDespesa[Fundo padrão],MATCH({tr("Categoria")},'
                                                f'tbCatDespesa[Categoria de gasto],0))&"","{po}")')
    nova_tabela(ws, L, f"B{H}:{CL(ncol + 1)}{ult}", COLS_LANC, formulas=formulas_lanc)
    for r in range(ult + 1, ult + 1001):           # formatos prontos para as próximas linhas
        for h, f in fmts.items():
            ws.cell(r, ci[h]).number_format = f
        ws.cell(r, ci["Valor"]).alignment = ALIGN_R
    ws.freeze_panes = f"{cl['Parcela']}{H + 1}"
    # Detalhes agrupados (recolhidos): clique no [+] para mostrar
    ws.column_dimensions.group(cl["Conta"], cl["Observação"], outline_level=1, hidden=True)
    ws.sheet_properties.outlinePr.summaryRight = False
    # Atalhos no topo
    put(ws, f"{cl['Parcela']}2", f'=HYPERLINK("#Lancamentos!B"&(ROWS({L}[Data])+{H + 1}),"✚  Ir para a próxima linha vazia")',
        font=fnt(11, True, P["cabecalho_txt"]), fill_=HEAD, align=ALIGN_C)
    ws.merge_cells(f"{cl['Parcela']}2:{cl['Vencimento']}2")
    put(ws, f"{cl['Tipo']}2", '=HYPERLINK("#Caderno!A1","📒  Caderno do mês")', font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_C)
    ws.merge_cells(f"{cl['Tipo']}2:{cl['Categoria']}2")
    put(ws, f"{cl['Pote']}2", '=HYPERLINK("#Inicio!A1","⌂  Início")', font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_C)
    ws.merge_cells(f"{cl['Pote']}2:{cl['Quem']}2")
    D["_linhas_importadas"] = (H + 1, ult)
    D["_col_lanc"] = cl

    LIM = ult + 20000
    faixa = lambda h, ini=H + 1: f"{cl[h]}{ini}:{cl[h]}{LIM}"
    dv_desc = DataValidation(type="list", formula1="=ListaFavoritos", allow_blank=True, showErrorMessage=False,
                             showInputMessage=True, promptTitle="Descrição",
                             prompt="Escolha um favorito (Alt+↓) ou digite. Com favorito, Tipo, Categoria, Pote e Situação se preenchem sozinhos.")
    ws.add_data_validation(dv_desc)
    dv_desc.add(faixa("Descrição"))
    dv_lista(ws, "=Tipos", faixa("Tipo"), msg="Escolha um tipo cadastrado na aba Config.")
    dv_lista(ws, "=Categorias", faixa("Categoria"), msg="Escolha uma categoria cadastrada na aba Config.")
    dv_lista(ws, "=Potes", faixa("Pote"), msg="Escolha um pote cadastrado na aba Config.")
    dv_lista(ws, "=Pessoas", faixa("Quem"), msg="Escolha uma pessoa cadastrada na aba Config.")
    dv_lista(ws, "=Contas", faixa("Conta"), msg="Escolha uma conta ou cartão cadastrado na aba Config.")
    dv_lista(ws, "=FormasPagamento", faixa("Forma de pagamento"), msg="Escolha uma forma de pagamento cadastrada na aba Config.")
    dv_sit = dv_lista(ws, "=Situacoes", faixa("Situação"), msg="Escolha uma situação cadastrada na aba Config.")
    dv_sit.showInputMessage, dv_sit.promptTitle = True, "Situação"
    dv_sit.prompt = ("Pago/Recebido entra nos totais. Use Pendente para contas futuras e troque para Pago quando pagar. "
                     "Descontar Depois: escreva de quem descontar na coluna ao lado de Quem.")
    dv_dd = DataValidation(type="list", formula1="=PessoasDescontar", allow_blank=True, showErrorMessage=False,
                           showInputMessage=True, promptTitle="Descontar de",
                           prompt="Com Situação = Descontar Depois: de quem descontar (Alt+↓). Soma sozinho na aba Acertos.")
    ws.add_data_validation(dv_dd)
    dv_dd.add(faixa("Descontar de"))
    dv_val = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv_val.error, dv_val.errorTitle, dv_val.showErrorMessage = "Digite um valor positivo: o Tipo define se é entrada ou saída.", "Valor inválido", True
    ws.add_data_validation(dv_val)
    dv_val.add(faixa("Valor"))
    # Data: lista com hoje / ontem / anteontem (Correção IA: "data automática com o dia de hoje").
    # Uma data que se preenche sozinha e fica fixa exige macro (TODAY() numa coluna mudaria todo dia);
    # a lista resolve com Alt+↓ Enter, e continua aceitando qualquer data digitada.
    for k, rot in enumerate(("hoje", "ontem", "anteontem")):
        ws_calc.cell(2 + k, 63, f"=TODAY()-{k}").number_format = FMT_DATA
        ws_calc.cell(2 + k, 64, rot)
    ws_calc.cell(1, 63, "Datas rápidas (lista da coluna Data)")
    wb.defined_names["DatasRapidas"] = DefinedName("DatasRapidas", attr_text="Calc!$BK$2:$BK$4")
    dv_hoje = DataValidation(type="list", formula1="=DatasRapidas", allow_blank=True, showErrorMessage=False,
                             showInputMessage=True, promptTitle="Data",
                             prompt="Alt+↓ e Enter = HOJE (ou ontem/anteontem). Também vale Ctrl+; ou digitar a data.")
    ws.add_data_validation(dv_hoje)
    dv_hoje.add(faixa("Data"))
    dv_dt = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
    dv_dt.error, dv_dt.errorTitle, dv_dt.showErrorMessage = "Digite uma data válida (dd/mm/aaaa). Dica: Ctrl+; insere a data de hoje.", "Data inválida", True
    dv_dt.showInputMessage, dv_dt.promptTitle, dv_dt.prompt = True, "Data", "Ctrl+;  insere a data de hoje."
    ws.add_data_validation(dv_dt)
    for h in ("Vencimento", "Competência"):
        dv_dt.add(faixa(h))

    # Formatação condicional (fórmulas de FC não aceitam referência estruturada: usa intervalos absolutos)
    abs_ = lambda h: f"${cl[h]}${H + 1}:${cl[h]}${LIM}"
    lin = lambda h: f"${cl[h]}{H + 1}"
    t1, s1, d1 = lin("Tipo"), lin("Situação"), lin("Descrição")
    ws.conditional_formatting.add(faixa("Valor"), FormulaRule(formula=[f'{t1}="Receita"'], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(faixa("Tipo"), FormulaRule(formula=[f'{t1}="Receita"'], font=Font(color="FF" + P["positivo"])))
    ws.conditional_formatting.add(faixa("Tipo"), FormulaRule(formula=[f'{t1}="Investimento"'], font=Font(color="FF" + P["destaque"])))
    ws.conditional_formatting.add(faixa("Tipo"), FormulaRule(formula=[f'AND({t1}<>"",{t1}<>"Receita",{t1}<>"Investimento")'], font=Font(color="FF" + P["negativo"])))
    falta = fill(misturar(P["alerta"], P["card"], 0.55))
    atraso = fill(misturar(P["negativo"], P["card"], 0.45))
    venc = f'IF({lin("Vencimento")}<>"",{lin("Vencimento")},{lin("Data")})'
    ws.conditional_formatting.add(faixa("Situação"), FormulaRule(formula=[f'AND({s1}="Pendente",{venc}<>"",{venc}<TODAY())'], fill=atraso, font=Font(color="FF" + P["texto"], bold=True), stopIfTrue=True))
    vence3 = fill(misturar(P["alerta"], P["card"], 0.55))
    ws.conditional_formatting.add(faixa("Situação"), FormulaRule(formula=[f'AND({s1}="Pendente",{venc}<>"",{venc}-TODAY()<=3)'], fill=vence3, font=Font(color="FF" + P["texto"], bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(faixa("Situação"), FormulaRule(formula=[f'{s1}="Pendente"'], font=Font(color="FF" + P["alerta"], bold=True)))
    ws.conditional_formatting.add(faixa("Situação"), FormulaRule(formula=[f'OR({s1}="Pago",{s1}="Recebido")'], font=Font(color="FF" + P["positivo"])))
    ws.conditional_formatting.add(faixa("Situação"), FormulaRule(formula=[f'{s1}="Descontar Depois"'], font=Font(color="FF" + P["roxo"], bold=True)))
    for h in ("Data", "Valor", "Categoria"):
        ws.conditional_formatting.add(faixa(h), FormulaRule(formula=[f'AND({d1}<>"",OR({lin(h)}="",{lin(h)}="Sem categoria"))'], fill=falta))
    ws.conditional_formatting.add(faixa("Descontar de"), FormulaRule(
        formula=[f'AND({s1}="Descontar Depois",{lin("Descontar de")}="")'], fill=atraso))
    # possível lançamento em dobro (mesma data, descrição e valor) — só nas linhas novas
    ini_novas = ult + 1
    ws.conditional_formatting.add(faixa("Descrição", ini_novas), FormulaRule(
        formula=[f'AND(${cl["Descrição"]}{ini_novas}<>"",COUNTIFS({abs_("Data")},${cl["Data"]}{ini_novas},{abs_("Descrição")},${cl["Descrição"]}{ini_novas},{abs_("Valor")},${cl["Valor"]}{ini_novas})>1)'],
        fill=falta, font=Font(color="FF" + P["alerta"], bold=True)))

    # ---------------------------------------------------------------- Favoritos (tbAtalhos)
    ws = ws_at
    atalhos = D["atalhos"] or [{h: None for h in COLS_ATALHO}]
    cab_at = COLS_ATALHO + ["Usos", "Último valor"]
    nat = len(cab_at)
    pintar_fundo(ws, nat + 3, 3)
    titulo_aba(ws, "Favoritos", "Descrições que você usa sempre. Ao escolher uma delas em Lançamentos, as colunas roxas se preenchem sozinhas. "
               "Recorrente? = sim entra no checklist Contas do mês. Adicione na linha vazia abaixo da tabela.", nat)
    ws.column_dimensions["A"].width = 2
    for j, (h, w) in enumerate(zip(cab_at, [30, 13, 24, 13, 11, 18, 18, 14, 12, 12, 8, 14])):
        ws.column_dimensions[CL(2 + j)].width = w
    HAt = 5
    for j, h in enumerate(cab_at):
        ws.cell(HAt, 2 + j, h)
    estilo_cabecalho(ws, HAt, 2, nat + 1)
    ws.row_dimensions[HAt].height = 30
    ult_at = HAt + len(atalhos)
    estilo_corpo(ws, HAt + 1, ult_at, 2, nat + 1)
    f_usos = "=COUNTIF(tbLancamentos[Descrição],tbAtalhos[[#This Row],[Descrição]])"
    f_ultimo = '=IFERROR(LOOKUP(2,1/(tbLancamentos[Descrição]=tbAtalhos[[#This Row],[Descrição]]),tbLancamentos[Valor]),"")'
    for i, a in enumerate(atalhos):
        r = HAt + 1 + i
        for j, h in enumerate(COLS_ATALHO):
            ws.cell(r, 2 + j, a.get(h))
        if (a.get("Descrição") or "").strip().lower() in rev_keys and a.get("Categoria") in PENDENTES_CAT:
            ws.cell(r, 2 + COLS_ATALHO.index("Categoria"), f_rev(f"$B{r}", a.get("Categoria")))
        ws.cell(r, 2 + len(COLS_ATALHO), f_usos)
        ws.cell(r, 3 + len(COLS_ATALHO), f_ultimo)
        for colx in (9, 13):
            ws.cell(r, colx).number_format = FMT_MOEDA
            ws.cell(r, colx).alignment = ALIGN_R
        for colx in (10, 11, 12):
            ws.cell(r, colx).alignment = ALIGN_C
        ws.cell(r, 12).font = fnt(10, color=P["suave"])
        ws.cell(r, 13).font = fnt(10, color=P["suave"])
    nova_tabela(ws, "tbAtalhos", f"B{HAt}:{CL(nat + 1)}{ult_at}", cab_at, formulas={"Usos": f_usos, "Último valor": f_ultimo})
    LA = ult_at + 500
    at_col = {h: CL(2 + j) for j, h in enumerate(cab_at)}
    fa = lambda h: f"{at_col[h]}{HAt + 1}:{at_col[h]}{LA}"
    dv_lista(ws, "=Tipos", fa("Tipo"))
    dv_lista(ws, "=Categorias", fa("Categoria"))
    dv_lista(ws, "=Potes", fa("Pote"))
    dv_lista(ws, "=Pessoas", fa("Quem"))
    dv_lista(ws, "=Contas", fa("Conta"))
    dv_lista(ws, "=FormasPagamento", fa("Forma de pagamento"))
    dv_lista(ws, '"sim,não"', fa("Recorrente?"), msg="Use sim ou não.")
    for r in range(HAt + 1, LA):
        ws[f"{at_col['Valor padrão']}{r}"].number_format = FMT_MOEDA
    ws.conditional_formatting.add(fa("Recorrente?"), FormulaRule(formula=[f'${at_col["Recorrente?"]}{HAt + 1}="sim"'], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.freeze_panes = f"C{HAt + 1}"
    wb.defined_names["ListaFavoritos"] = DefinedName("ListaFavoritos", attr_text=f"OFFSET(Atalhos!$B${HAt + 1},0,0,MAX(1,SUMPRODUCT(--(LEN(Atalhos!$B${HAt + 1}:$B${LA})>0))),1)")
    D["_atalhos_pos"] = (HAt, at_col)

    # ---------------------------------------------------------------- Layout do Histórico + regras de aporte
    tipos_gasto = [t for t in D["tipos"] if t not in ("Receita", "Investimento")]
    cab_h = (["Mês", "Receitas"] + tipos_gasto + (["Total de gastos"] if len(tipos_gasto) > 1 else [])
             + ["Resultado", "Aporte", "Resgate", "Saldo livre", "% comprometido", "Conferido em"])
    hist_col = {h: CL(2 + i) for i, h in enumerate(cab_h)}
    comps = [_competencia(d) for d in lanc]
    (a0, m0), (a1, m1) = min(comps), max(comps)
    a1, m1 = max((a1, 12), (dt.date.today().year, 12))
    meses_hist = []
    a, m = a0, m0
    while (a, m) <= (a1, m1):
        meses_hist.append((a, m))
        a, m = _mais_meses(a, m, 1)
    tem_inv = bool(D.get("investimentos"))

    def aporte_f(mes, ano):
        """Aporte do mês = aplicações novas da aba Investimentos naquele mês — sem as "Reaplicação"
        (dinheiro que já estava investido) e sem linha de "correção". É a mesma regra do aporte da Tela do
        Dinheiro do Plínio (conferido). Sem a aba, usa os lançamentos de Tipo Investimento."""
        if not tem_inv:
            return f'SUMIFS({L}[Valor],{L}[Tipo],"Investimento",{L}[Mês],{mes},{L}[Ano],{ano})'
        return (f'SUMPRODUCT((YEAR(tbInvestimentos[Data])={ano})*(MONTH(tbInvestimentos[Data])={mes})'
                f'*(tbInvestimentos[Aplicado]>0)*ISERROR(SEARCH("reaplica",tbInvestimentos[Observação]))'
                f'*ISERROR(SEARCH("correção",tbInvestimentos[Observação]))*tbInvestimentos[Aplicado])')

    def resgate_f(mes, ano):
        """Resgate do mês = valores negativos lançados na aba Investimentos (a sua convenção de resgate parcial)."""
        if not tem_inv:
            return "0"
        return (f'SUMPRODUCT((YEAR(tbInvestimentos[Data])={ano})*(MONTH(tbInvestimentos[Data])={mes})'
                f'*(tbInvestimentos[Aplicado]<0)*(-tbInvestimentos[Aplicado]))')

    # ---------------------------------------------------------------- Calc (oculta)
    ws = c = ws_calc
    ws.sheet_state = "hidden"
    larguras(ws, {"A": 30, "B": 14})

    def sif(tipo, mes, ano, extra=""):
        crit_tipo = f'{L}[Tipo],"Receita"' if tipo == "Receita" else f'{L}[Tipo],"<>Receita"'
        return f'SUMIFS({L}[Valor],{crit_tipo},{CONTA},{L}[Mês],{mes},{L}[Ano],{ano}{extra})'

    conta_cfg = lambda tab: f"=SUMPRODUCT(--(LEN({rng_cfg(tab)})>0))"
    param = [
        ("Mês selecionado (nº)", "=MATCH(MesSelecionado,Meses,0)"),                   # B2
        ("Ano selecionado", "=AnoSelecionado"),                                        # B3
        ("Data de referência", "=DATE(B3,B2,1)"),                                      # B4
        ("Mês anterior (data)", "=DATE(B3,B2-1,1)"),                                   # B5
        ("Mês anterior (nº)", "=MONTH(B5)"),                                           # B6
        ("Ano do mês anterior", "=YEAR(B5)"),                                          # B7
        ("", None),                                                                    # B8
        ("Receitas do mês", "=" + sif("Receita", "$B$2", "$B$3")),                     # B9
        ("Gastos do mês", "=" + sif("Gasto", "$B$2", "$B$3")),                         # B10
        ("Receitas mês anterior", "=" + sif("Receita", "$B$6", "$B$7")),               # B11
        ("Gastos mês anterior", "=" + sif("Gasto", "$B$6", "$B$7")),                   # B12
        ("Saldo do mês", "=B9-B10"),                                                   # B13
        ("Saldo mês anterior", "=B11-B12"),                                            # B14
        ("Taxa de poupança", "=IF(B9>0,B13/B9,0)"),                                    # B15
        ("Variação gastos vs. anterior", "=IF(B12>0,B10/B12-1,0)"),                    # B16
        ("Orçamento total (metas)", "=SUM(tbOrcamento[Meta mensal])"),                 # B17
        ("% do orçamento utilizado", "=IF(B17>0,B10/B17,0)"),                          # B18
        ("Taxa de poupança mês anterior", "=IF(B11>0,B14/B11,0)"),                     # B19
        ("Qtde categorias receita", conta_cfg("tbCatReceita")),                        # B20
        ("Qtde categorias gasto", conta_cfg("tbCatDespesa")),                          # B21
        ("Qtde contas", conta_cfg("tbContas")),                                        # B22
        ("Qtde cartões", conta_cfg("tbCartoes")),                                      # B23
        ("Qtde categorias no orçamento", "=COUNTA(tbOrcamento[Categoria])"),           # B24
        ("Diferença gastos vs. anterior", "=B10-B12"),                                 # B25
        ("Meses com dados no ano", "=SUMPRODUCT(--(L2:L13+M2:M13>0))"),                # B26
        ("Qtde categorias outras", conta_cfg("tbCatOutras")),                          # B27
    ]
    put(c, "A1", "PARÂMETROS", font=fnt(10, True))
    for i, (rot, f) in enumerate(param):
        c.cell(2 + i, 1, rot)
        if f:
            c.cell(2 + i, 2, f)
    for ref in ("B4", "B5"):
        c[ref].number_format = FMT_DATA
    for ref in ("B15", "B16", "B18", "B19"):
        c[ref].number_format = FMT_PCT
    wb.defined_names["MesNum"] = DefinedName("MesNum", attr_text="Calc!$B$2")

    # Últimos 12 meses (D:H)
    for j, h in enumerate(["Data", "Rótulo", "Receitas", "Gastos", "Saldo"]):
        c.cell(1, 4 + j, h)
    for i in range(12):
        r = 2 + i
        c.cell(r, 4, f"=DATE($B$3,$B$2-{11 - i},1)").number_format = FMT_DATA
        c.cell(r, 5, f'=INDEX(MesesAbrev,MONTH(D{r}))&"/"&RIGHT(YEAR(D{r}),2)')
        c.cell(r, 6, "=" + sif("Receita", f"MONTH(D{r})", f"YEAR(D{r})"))
        c.cell(r, 7, "=" + sif("Gasto", f"MONTH(D{r})", f"YEAR(D{r})"))
        c.cell(r, 8, f"=F{r}-G{r}")

    # Ano selecionado mês a mês (J:O)
    for j, h in enumerate(["Mês", "Rótulo", "Receitas", "Gastos", "Saldo", "Saldo acumulado"]):
        c.cell(1, 10 + j, h)
    for i in range(12):
        r = 2 + i
        c.cell(r, 10, i + 1)
        c.cell(r, 11, f"=INDEX(MesesAbrev,J{r})")
        c.cell(r, 12, "=" + sif("Receita", f"J{r}", "$B$3"))
        c.cell(r, 13, "=" + sif("Gasto", f"J{r}", "$B$3"))
        c.cell(r, 14, f"=L{r}-M{r}")
        c.cell(r, 15, f"=SUM($N$2:N{r})")

    # Gastos por categoria no mês (Q:S) + ranking (T:X) p/ a rosca
    NSLOT = 60
    rng_desp = rng_cfg("tbCatDespesa")
    for j, h in enumerate(["Categoria", "Realizado", "Chave ordenação"]):
        c.cell(1, 17 + j, h)
    for i in range(NSLOT):
        r = 2 + i
        c.cell(r, 17, f'=IF({i + 1}<=$B$21,INDEX({rng_desp},{i + 1}),"")')
        c.cell(r, 18, f'=IF(Q{r}="",0,SUMIFS({L}[Valor],{L}[Categoria],Q{r},{L}[Tipo],"<>Receita",{CONTA},{L}[Mês],$B$2,{L}[Ano],$B$3))')
        c.cell(r, 19, f'=IF(Q{r}="",-1,ROUND(R{r},2)+({NSLOT + 1 - (i + 1)})/10000000)')
    for j, h in enumerate(["Posição", "Chave", "Linha", "Categoria (rosca)", "Valor (rosca)"]):
        c.cell(1, 20 + j, h)
    for k in range(1, 8):
        r = 1 + k
        c.cell(r, 20, k)
        c.cell(r, 21, f"=LARGE($S$2:$S${NSLOT + 1},T{r})")
        c.cell(r, 22, f"=MATCH(U{r},$S$2:$S${NSLOT + 1},0)")
        c.cell(r, 23, f'=IF(U{r}<=0,"",INDEX($Q$2:$Q${NSLOT + 1},V{r}))')
        c.cell(r, 24, f"=IF(U{r}<=0,0,INDEX($R$2:$R${NSLOT + 1},V{r}))")
    c.cell(9, 20, 8)
    c.cell(9, 23, "Demais categorias")
    c.cell(9, 24, "=MAX(0,$B$10-SUM(X2:X8))")

    # Orçado x realizado (Z:AC) — espelha tbOrcamento
    N_ORC = max(1, len(D["orcamento"]))
    for j, h in enumerate(["Categoria", "Orçado", "Realizado", "% usado"]):
        c.cell(1, 26 + j, h)
    for i in range(N_ORC):
        r = 2 + i
        c.cell(r, 26, f'=IFERROR(INDEX(tbOrcamento[Categoria],{i + 1})&"","")')
        c.cell(r, 27, f"=IFERROR(INDEX(tbOrcamento[Meta mensal],{i + 1})+0,0)")
        c.cell(r, 28, f"=IFERROR(INDEX(tbOrcamento[Realizado no mês selecionado],{i + 1})+0,0)")
        c.cell(r, 29, f"=IF(AA{r}>0,AB{r}/AA{r},0)")

    # Top 5 maiores gastos do mês (AE:AK) — LARGE + INDEX/MATCH com critério de mês
    for j, h in enumerate(["Posição", "Chave (valor+linha)", "Linha", "Data", "Descrição", "Categoria", "Valor"]):
        c.cell(1, 31 + j, h)
    cond = f'({L}[Tipo]<>"Receita")*({L}[Conta no mês?]="sim")*({L}[Mês]=$B$2)*({L}[Ano]=$B$3)'
    for k in range(1, 6):
        r = 1 + k
        c.cell(r, 31, k)
        c[f"AF{r}"] = ArrayFormula(f"AF{r}", f"=IFERROR(LARGE(IF({cond},ROUND({L}[Valor],2)+ROW({L}[Valor])/1000000000),AE{r}),0)")
        c[f"AG{r}"] = ArrayFormula(f"AG{r}", f"=IF(AF{r}<=0,0,MATCH(ROUND((AF{r}-ROUND(AF{r},2))*1000000000,0),ROW({L}[Valor]),0))")
        c.cell(r, 34, f'=IF(AG{r}=0,"",INDEX({L}[Data],AG{r}))')
        c.cell(r, 35, f'=IF(AG{r}=0,"",INDEX({L}[Descrição],AG{r}))')
        c.cell(r, 36, f'=IF(AG{r}=0,"",INDEX({L}[Categoria],AG{r}))')
        c.cell(r, 37, f'=IF(AG{r}=0,"",INDEX({L}[Valor],AG{r}))')

    # Listas combinadas para os menus suspensos (AM: categorias, AO: contas+cartões)
    rr, rd, ro = rng_cfg("tbCatReceita"), rng_cfg("tbCatDespesa"), rng_cfg("tbCatOutras")
    c.cell(1, 39, "Categorias (receitas + gastos + outras)")
    for i in range(150):
        k = i + 1
        c.cell(2 + i, 39, f'=IF({k}<=$B$20,INDEX({rr},{k}),IF({k}-$B$20<=$B$21,INDEX({rd},MAX(1,{k}-$B$20)),'
                          f'IF({k}-$B$20-$B$21<=$B$27,INDEX({ro},MAX(1,{k}-$B$20-$B$21)),"")))')
    rc_, rk = rng_cfg("tbContas"), rng_cfg("tbCartoes")
    c.cell(1, 41, "Contas + cartões")
    for i in range(60):
        k = i + 1
        c.cell(2 + i, 41, f'=IF({k}<=$B$22,INDEX({rc_},{k}),IF({k}-$B$22<=$B$23,INDEX({rk},MAX(1,{k}-$B$22)),""))')
    wb.defined_names["Categorias"] = DefinedName("Categorias", attr_text="OFFSET(Calc!$AM$2,0,0,MAX(1,Calc!$B$20+Calc!$B$21+Calc!$B$27),1)")
    wb.defined_names["Contas"] = DefinedName("Contas", attr_text="OFFSET(Calc!$AO$2,0,0,MAX(1,Calc!$B$22+Calc!$B$23),1)")

    # Apoio ao Início e ao checklist (AQ em diante)
    HAt, at_col = D["_atalhos_pos"]
    NREC = max(10, sum(1 for a in D["atalhos"] if a.get("Recorrente?") == "sim") + 15)
    c.cell(1, 43, "Nº recorrente")
    c.cell(1, 44, "Linha do favorito recorrente")
    for k in range(1, NREC + 1):
        r = 1 + k
        c.cell(r, 43, k)
        c[f"AR{r}"] = ArrayFormula(f"AR{r}", f'=IFERROR(SMALL(IF(tbAtalhos[Recorrente?]="sim",ROW(tbAtalhos[Recorrente?])),AQ{r}),0)')
    lc = D["_col_lanc"]
    for j, h in enumerate(["Nº", "Chave (vencimento+linha)", "Linha", "Vence", "Descrição", "Valor", "Status"]):
        c.cell(1, 46 + j, h)
    cond_pend = f'({L}[Situação]="Pendente")*({L}[Tipo]<>"Receita")'
    for k in range(1, 9):
        r = 1 + k
        c.cell(r, 46, k)
        c[f"AU{r}"] = ArrayFormula(f"AU{r}", f'=IFERROR(SMALL(IF({cond_pend},IF({L}[Vencimento]<>"",{L}[Vencimento],{L}[Data])+ROW({L}[Valor])/1000000),AT{r}),0)')
        c.cell(r, 48, f"=IF(AU{r}=0,0,ROUND((AU{r}-INT(AU{r}))*1000000,0))")
        c.cell(r, 49, f'=IF(AU{r}=0,"",INT(AU{r}))').number_format = FMT_DATA
        c.cell(r, 50, f'=IF(AV{r}=0,"",INDEX(Lancamentos!${lc["Descrição"]}:${lc["Descrição"]},AV{r}))')
        c.cell(r, 51, f'=IF(AV{r}=0,"",INDEX(Lancamentos!${lc["Valor"]}:${lc["Valor"]},AV{r}))')
        c.cell(r, 52, f'=IF(AW{r}="","",IF(AW{r}<TODAY(),"Atrasada",IF(AW{r}-TODAY()<=7,"Esta semana","Em breve")))')
    for j, h in enumerate(["Nº", "Data", "Descrição", "Valor", "Categoria", "Tipo"]):
        c.cell(1, 54 + j, h)
    for k in range(1, 9):
        r = 1 + k
        c.cell(r, 54, k)
        pos = f"ROWS({L}[Data])-BB{r}+1"
        c.cell(r, 55, f'=IFERROR(IF(INDEX({L}[Data],{pos})="","",INDEX({L}[Data],{pos})),"")').number_format = FMT_DATA
        c.cell(r, 56, f'=IFERROR(INDEX({L}[Descrição],{pos})&"","")')
        c.cell(r, 57, f'=IFERROR(IF(INDEX({L}[Valor],{pos})="","",INDEX({L}[Valor],{pos})),"")')
        c.cell(r, 58, f'=IFERROR(INDEX({L}[Categoria],{pos})&"","")')
        c.cell(r, 59, f'=IFERROR(INDEX({L}[Tipo],{pos})&"","")')
    hs = hist_col
    extras = [
        (30, "Aporte no mês", "=" + aporte_f("$B$2", "$B$3"), FMT_MOEDA),
        (31, "Resgate no mês", "=" + resgate_f("$B$2", "$B$3"), FMT_MOEDA),
        (32, "Saldo livre do mês", "=B13-B30+B31", FMT_MOEDA),
        (33, "% comprometido da renda", "=IF(B9>0,B10/B9,0)", FMT_PCT),
        (34, "(livre — antes: saldo que veio do mês anterior, retirado a pedido)", "=0", FMT_MOEDA),
        (35, "Aporte mês anterior", "=" + aporte_f("$B$6", "$B$7"), FMT_MOEDA),
        (36, "Saldo livre mês anterior", "=B14-B35+" + resgate_f("$B$6", "$B$7"), FMT_MOEDA),
        (37, "% comprometido mês anterior", "=IF(B11>0,B12/B11,0)", FMT_PCT),
        (38, "Total investido (bruto)", "=SUM(tbInvestimentos[Saldo hoje])" if tem_inv else "=0", FMT_MOEDA),
        (39, "IR estimado se resgatar hoje", "=SUM(tbInvestimentos[IR])" if tem_inv else "=0", FMT_MOEDA),
        (40, "Investimento líquido", "=SUM(tbInvestimentos[Líquido])" if tem_inv else "=0", FMT_MOEDA),
        (41, "(livre — antes: saldo acumulado, retirado a pedido)", "=0", FMT_MOEDA),
        (42, "Gastos pendentes no mês (previsto, fora do total)",
         f'=SUMIFS({L}[Valor],{L}[Situação],"Pendente",{L}[Tipo],"<>Receita",{L}[Tipo],"<>Investimento",{L}[Mês],$B$2,{L}[Ano],$B$3)', FMT_MOEDA),
        (43, "Receitas pendentes no mês (a receber)",
         f'=SUMIFS({L}[Valor],{L}[Situação],"Pendente",{L}[Tipo],"Receita",{L}[Mês],$B$2,{L}[Ano],$B$3)', FMT_MOEDA),
    ]
    for rr_, rot, f, fmt_ in extras:
        c.cell(rr_, 1, rot)
        c.cell(rr_, 2, f).number_format = fmt_
    c.cell(28, 1, "A pagar no mês (pendentes)")
    c.cell(28, 2, f'=SUMIFS({L}[Valor],{L}[Situação],"Pendente",{L}[Tipo],"<>Receita",{L}[Mês],$B$2,{L}[Ano],$B$3)')
    c.cell(29, 1, "Pendências atrasadas (qtde)")
    c["B29"] = ArrayFormula("B29", f'=SUM(IF({cond_pend},IF(IF({L}[Vencimento]<>"",{L}[Vencimento],{L}[Data])<TODAY(),1,0),0))')

    # ---------------------------------------------------------------- Categorizar (nada fica em Outros)
    ws = ws_rev
    rev = D.get("revisao") or []
    linhas_rev = rev or [dict(desc=None, tipo=None, vezes=None, total=None, ultimo=None, atual=None, sugestao=None, certa=None, pergunta=None)]
    pintar_fundo(ws, 12, 3)
    titulo_aba(ws, "Categorizar", "Tudo o que estava em Outros ou Sem categoria, agrupado por descrição. Escolha a Categoria certa (Alt+↓): "
               "todos os lançamentos com essa descrição mudam juntos, e o Pote segue o fundo padrão da categoria.", 10)
    larguras(ws, {"A": 2, "B": 30, "C": 12, "D": 7, "E": 14, "F": 12, "G": 15, "H": 22, "I": 24, "J": 46})
    cab_r = ["Descrição", "Tipo", "Vezes", "Total", "Último", "Categoria atual", "Sugestão", "Categoria certa", "Pergunta / do que se trata"]
    HR = 9
    cards_r = [("B", "C", "FALTAM RESPONDER", '=COUNTIFS(tbRevisaoCat[Descrição],"<>",tbRevisaoCat[Categoria certa],"")', P["negativo"]),
               ("E", "F", "RESPONDIDAS", '=COUNTIFS(tbRevisaoCat[Descrição],"<>",tbRevisaoCat[Categoria certa],"<>")', P["positivo"]),
               ("H", "H", "LANÇAMENTOS AFETADOS", "=SUM(tbRevisaoCat[Vezes])", P["destaque"])]
    for a_, b_, rot, f, cor in cards_r:
        area(ws, f"{a_}5:{b_}7", fill_=CARD)
        mesclar(ws, f"{a_}5:{b_}5", rot, fill_=CARD, font=fnt(9, True, P["suave"]), align=Alignment(horizontal="left", vertical="bottom", indent=1),
                border=Border(top=side(cor, "thick")))
        mesclar(ws, f"{a_}6:{b_}6", f, fill_=CARD, font=fnt(15, True, cor), fmt="0", align=Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[6].height = 30
    caixa_ajuda(ws, "J", "J", 4, [
        "1. Em Categoria certa, escolha a categoria (Alt+↓). A Sugestão é um palpite: confira.",
        "2. Pronto: todos os lançamentos com essa descrição (e o favorito) passam para ela.",
        "3. Não sabe o que é? Escreva na coluna Pergunta e me mande pedir em Correções IA.",
        "Na próxima atualização, o que foi respondido fica fixo e sai desta lista."])
    for j, h in enumerate(cab_r):
        ws.cell(HR, 2 + j, h)
    estilo_cabecalho(ws, HR, 2, 10)
    ws.cell(HR, 9).fill = fill(misturar(P["destaque"], P["card"], 0.3))
    ws.cell(HR, 9).font = fnt(10, True, P["botao_txt"])
    estilo_corpo(ws, HR + 1, HR + len(linhas_rev), 2, 10)
    for i, g in enumerate(linhas_rev):
        r = HR + 1 + i
        for j, k in enumerate(("desc", "tipo", "vezes", "total", "ultimo", "atual", "sugestao")):
            ws.cell(r, 2 + j, g[k])
        put(ws, f"I{r}", g["certa"], font=fnt(10, True, P["destaque"]), fill_=CARD2, border=Border(bottom=side(P["destaque"])))
        put(ws, f"J{r}", g["pergunta"], font=fnt(9, color=P["suave"], italic=True),
            align=Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1))
        ws.cell(r, 5).number_format, ws.cell(r, 5).alignment = FMT_MOEDA, ALIGN_R
        ws.cell(r, 6).number_format, ws.cell(r, 6).alignment = FMT_DATA, ALIGN_C
        ws.cell(r, 4).alignment = ALIGN_C
        ws.cell(r, 7).font = fnt(10, color=P["suave"])
        ws.cell(r, 8).font = fnt(10, color=P["suave"])
    nova_tabela(ws, "tbRevisaoCat", f"B{HR}:J{HR + len(linhas_rev)}", cab_r)
    dv_lista(ws, "=Categorias", f"I{HR + 1}:I{HR + len(linhas_rev) + 200}", msg="Escolha uma categoria da aba Config (ou cadastre uma nova lá).")
    ws.conditional_formatting.add(f"I{HR + 1}:I{HR + len(linhas_rev)}", FormulaRule(
        formula=[f'AND($B{HR + 1}<>"",$I{HR + 1}="")'], fill=fill(misturar(P["negativo"], P["card"], 0.8))))
    put(ws, "H2", '=HYPERLINK("#Inicio!A1","⌂  Início")', font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_C)
    ws.freeze_panes = f"C{HR + 1}"

    # ---------------------------------------------------------------- Contas do mês (checklist dos recorrentes)
    ws = ws_contas
    pintar_fundo(ws, 26, 3)
    titulo_aba(ws, "Contas do mês", "Checklist das contas e receitas recorrentes (marcadas em Favoritos). O status vem dos lançamentos do mês de trabalho escolhido no Caderno.", 20)
    put(ws, "E2", '="Mês: "&MesSelecionado&" / "&AnoSelecionado', font=fnt(13, True, P["destaque"]), align=Alignment(horizontal="right", vertical="center"))
    ws.merge_cells("E2:G2")
    put(ws, "I2", '=HYPERLINK("#Inicio!A1","⌂  Início")', font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_C)
    larguras(ws, {"A": 2, "B": 28, "C": 22, "D": 7, "E": 21, "F": 21, "G": 26, "H": 6, "I": 12, "J": 6, "K": 6, "L": 3})
    for colx in ("H", "I", "J", "K"):
        ws.column_dimensions[colx].hidden = True
    HC = 16          # linhas 9–14: caixa "como funciona" (recolhida no [+], Correção IA: mais espaço para a lista)
    ini_c, fim_c = HC + 1, HC + NREC
    # cards de resumo
    cards = [("B", "C", "CONTAS LANÇADAS", f'=SUMPRODUCT(--(I{ini_c}:I{fim_c}>0))&" de "&SUMPRODUCT(--(LEN(B{ini_c}:B{fim_c})>0))', None, P["destaque"]),
             ("E", "E", "PREVISTO", f"=SUM(E{ini_c}:E{fim_c})", FMT_MOEDA, P["suave"]),
             ("F", "F", "JÁ LANÇADO", f"=SUM(F{ini_c}:F{fim_c})", FMT_MOEDA, P["positivo"]),
             ("G", "G", "FALTA LANÇAR (PREVISTO)", f'=SUMPRODUCT((LEFT(G{ini_c}:G{fim_c},1)="✖")*N(+E{ini_c}:E{fim_c}))', FMT_MOEDA, P["negativo"])]
    for a, b, rot, f, fmt_, cor in cards:
        area(ws, f"{a}5:{b}7", fill_=CARD)
        area(ws, f"{a}5:{b}5", border=Border(top=side(cor, "thick")))
        mesclar(ws, f"{a}5:{b}5", rot, fill_=CARD, font=fnt(9, True, P["suave"]), align=Alignment(horizontal="left", vertical="bottom", indent=1),
                border=Border(top=side(cor, "thick")))
        mesclar(ws, f"{a}6:{b}6", f, fill_=CARD, font=fnt(13, True), fmt=fmt_, align=Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 30
    ws.row_dimensions[7].height = 16
    D["_contas_resumo"] = ("B6", f"G6")
    caixa_ajuda(ws, "B", "G", 9, [
        "É o checklist das contas que se repetem todo mês: tudo que está marcado Recorrente? = sim na aba Atalhos (Favoritos).",
        "Para cada conta, olha os lançamentos do mês de trabalho (Caderno) e diz: ✔ Pago · ⏳ Pendente · ✖ Falta lançar.",
        "Previsto = o Valor padrão do favorito (ou o que você pagou no mês anterior). Lançado = o que já está em Lançamentos.",
        "À direita, o bloco pronto para colar traz só o que falta lançar: copie, vá à 1ª linha vazia e cole só os valores.",
        "Conta nova que se repete? Na aba Atalhos, marque Recorrente? = sim e o dia do vencimento. Deixou de pagar? Marque não."])
    for r_box in range(9, 15):
        ws.row_dimensions[r_box].outlineLevel = 1
        ws.row_dimensions[r_box].hidden = True
    ws.sheet_properties.outlinePr.summaryBelow = False
    put(ws, "B8", "ⓘ  Como funciona: clique no [+] à esquerda para abrir a explicação.", font=fnt(9, color=P["suave"], italic=True))
    ws.row_dimensions[8].height = 16
    cab_c = ["Conta / receita", "Categoria", "Dia", "Previsto", "Lançado no mês", "Status", "Linha", "Lançamentos", "Pendentes", "Falta nº"]
    for j, h in enumerate(cab_c):
        ws.cell(HC, 2 + j, h)
    estilo_cabecalho(ws, HC, 2, 7)
    estilo_corpo(ws, ini_c, fim_c, 2, 7)
    A = lambda col, r: f"INDEX(Atalhos!${at_col[col]}:${at_col[col]},$H{r})"
    for k in range(1, NREC + 1):
        r = HC + k
        cr = 1 + k
        ws.cell(r, 8, f"=Calc!AR{cr}")
        ws.cell(r, 2, f'=IF($H{r}=0,"",{A("Descrição", r)}&"")')
        ws.cell(r, 3, f'=IF($H{r}=0,"",{A("Categoria", r)}&"")')
        ws.cell(r, 4, f'=IF($H{r}=0,"",IF({A("Dia do vencimento", r)}="","",{A("Dia do vencimento", r)}))')
        ws.cell(r, 5, f'=IF(B{r}="","",IF(N({A("Valor padrão", r)})>0,{A("Valor padrão", r)},'
                      f'SUMIFS({L}[Valor],{L}[Descrição],B{r},{L}[Mês],Calc!$B$6,{L}[Ano],Calc!$B$7)))')
        ws.cell(r, 6, f'=IF(B{r}="","",SUMIFS({L}[Valor],{L}[Descrição],B{r},{L}[Mês],MesNum,{L}[Ano],AnoSelecionado))')
        ws.cell(r, 9, f'=IF(B{r}="",0,COUNTIFS({L}[Descrição],B{r},{L}[Mês],MesNum,{L}[Ano],AnoSelecionado))')
        ws.cell(r, 10, f'=IF(B{r}="",0,COUNTIFS({L}[Descrição],B{r},{L}[Mês],MesNum,{L}[Ano],AnoSelecionado,{L}[Situação],"Pendente"))')
        vence = f'DATE(AnoSelecionado,MesNum,MIN(IF(D{r}="",1,D{r}),DAY(EOMONTH(DATE(AnoSelecionado,MesNum,1),0))))'
        ws.cell(r, 7, f'=IF(B{r}="","",IF(I{r}=0,IF(AND(D{r}<>"",{vence}<TODAY()),"✖ Atrasada — falta lançar","✖ Falta lançar"),'
                      f'IF(J{r}>0,IF({vence}<TODAY(),"⏳ Pendente — vencida","⏳ Pendente"),IF({A("Tipo", r)}="Receita","✔ Recebido","✔ Pago"))))')
        ws.cell(r, 11, f'=IF(LEFT(G{r},1)="✖",SUMPRODUCT(--(LEFT($G${ini_c}:G{r},1)="✖")),0)')
        for colx in (5, 6):
            ws.cell(r, colx).number_format = FMT_MOEDA
            ws.cell(r, colx).alignment = ALIGN_R
        ws.cell(r, 4).alignment = ALIGN_C
        ws.cell(r, 7).font = fnt(10, True)
    rngG = f"G{ini_c}:G{fim_c}"
    ws.conditional_formatting.add(rngG, FormulaRule(formula=[f'LEFT(G{ini_c},1)="✔"'], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(rngG, FormulaRule(formula=[f'LEFT(G{ini_c},1)="⏳"'], font=Font(color="FF" + P["alerta"], bold=True)))
    ws.conditional_formatting.add(rngG, FormulaRule(formula=[f'LEFT(G{ini_c},1)="✖"'], font=Font(color="FF" + P["negativo"], bold=True)))
    # bloco pronto para colar em Lançamentos (só o que falta lançar)
    BC = 13   # coluna M
    for j, h in enumerate(COLS_BLOCO):
        ws.column_dimensions[CL(BC + j)].width = LARG_BLOCO[h]
        ws.cell(HC, BC + j, h)
    estilo_cabecalho(ws, HC, BC, BC + len(COLS_BLOCO) - 1)
    estilo_corpo(ws, ini_c, fim_c, BC, BC + len(COLS_BLOCO) - 1, fill_=CARD2)
    put(ws, f"{CL(BC)}4", "BLOCO PRONTO PARA COLAR — só as contas que faltam lançar", font=fnt(10, True, P["suave"]))
    instr = ["1. Selecione as linhas preenchidas do bloco abaixo e copie (Ctrl+C).",
             "2. Clique no botão ao lado para ir à primeira linha vazia de Lançamentos.",
             "3. Cole só os valores: Ctrl+Shift+V (ou Ctrl+Alt+V ➜ Valores).",
             "4. Elas entram como Pendente: troque para Pago quando pagar."]
    for i, t in enumerate(instr):
        put(ws, f"{CL(BC)}{5 + i}", t, font=fnt(9, color=P["suave"]))
    put(ws, f"{CL(BC + 6)}5", f'=HYPERLINK("#Lancamentos!B"&(ROWS({L}[Data])+{D["_linhas_importadas"][0]}),"✚  Ir para a primeira linha vazia")',
        font=fnt(11, True, P["cabecalho_txt"]), fill_=HEAD, align=ALIGN_C)
    ws.merge_cells(f"{CL(BC + 6)}5:{CL(BC + 9)}6")
    for k in range(1, NREC + 1):
        r = HC + k
        idx = f"MATCH({k},$K${ini_c}:$K${fim_c},0)"
        lin_at = f"INDEX($H${ini_c}:$H${fim_c},{idx})"
        Ai = lambda col: f'INDEX(Atalhos!${at_col[col]}:${at_col[col]},{lin_at})&""'
        dia = f"INDEX($D${ini_c}:$D${fim_c},{idx})"
        data_f = f'IFERROR(DATE(AnoSelecionado,MesNum,MIN(IF({dia}="",1,{dia}),DAY(EOMONTH(DATE(AnoSelecionado,MesNum,1),0)))),"")'
        mapa = {
            "Data": "=" + data_f, "Vencimento": "=" + data_f,
            "Descrição": f'=IFERROR(INDEX($B${ini_c}:$B${fim_c},{idx}),"")',
            "Valor": f'=IFERROR(IF(N(INDEX($E${ini_c}:$E${fim_c},{idx}))>0,INDEX($E${ini_c}:$E${fim_c},{idx}),""),"")',
            "Tipo": f'=IFERROR({Ai("Tipo")},"")', "Categoria": f'=IFERROR({Ai("Categoria")},"")',
            "Situação": f'=IFERROR(IF({idx}>0,"Pendente",""),"")',
            "Pote": f'=IFERROR(IF({Ai("Pote")}<>"",{Ai("Pote")},INDEX(tbCatDespesa[Fundo padrão],MATCH({Ai("Categoria")},tbCatDespesa[Categoria de gasto],0))&""),"")',
            "Quem": f'=IFERROR(IF({Ai("Quem")}<>"",{Ai("Quem")},QuemPadrao),"")',
            "Conta": f'=IFERROR({Ai("Conta")},"")', "Forma de pagamento": f'=IFERROR({Ai("Forma de pagamento")},"")',
            "Parcela": '=""', "Descontar de": '=""',
        }
        for j, h in enumerate(COLS_BLOCO):
            cc = ws.cell(r, BC + j, mapa[h])
            if h in ("Data", "Vencimento"):
                cc.number_format = FMT_DATA
            elif h == "Valor":
                cc.number_format, cc.alignment = FMT_MOEDA, ALIGN_R
    ws.freeze_panes = f"A{HC + 1}"

    # ---------------------------------------------------------------- Parcelar / repetir
    ws = ws_parc
    pintar_fundo(ws, 20, 3)
    titulo_aba(ws, "Parcelar ou repetir", "Preencha os campos à esquerda: o bloco à direita gera uma linha por parcela (ou por mês). Copie, vá até a primeira linha vazia de Lançamentos e cole só os valores.", 16)
    larguras(ws, {"A": 2, "B": 30, "C": 22, "D": 3})
    put(ws, "B4", "COMPRA / LANÇAMENTO", font=fnt(10, True, P["suave"]))
    campos = [
        ("Modo", "Parcelado", '"Parcelado,Repetir todo mês"', "Parcelado: divide em parcelas 1/n, 2/n...  Repetir: o mesmo valor todo mês (aluguel, mesada, assinatura)."),
        ("Descrição", "Geladeira nova (exemplo)", "=ListaFavoritos", "Pode escolher um favorito: tipo, categoria e pote vêm dele se ficarem vazios abaixo."),
        ("Valor (R$)", 3600, None, None),
        ("O valor informado é", "Total da compra", '"Total da compra,Valor de cada parcela"', None),
        ("Nº de parcelas / meses", 10, None, "Entre 1 e 60."),
        ("Data da 1ª parcela", dt.date(D["ano_padrao"] + (D["mes_padrao"] // 12), D["mes_padrao"] % 12 + 1, 10), None, None),
        ("Tipo", tipo_padrao, "=Tipos", None),
        ("Categoria (vazio = do favorito)", None, "=Categorias", None),
        ("Pote (vazio = do favorito)", None, "=Potes", None),
        ("Quem", None, "=Pessoas", None),
        ("Conta / cartão", None, "=Contas", None),
        ("Forma de pagamento", "Cartão de crédito" if "Cartão de crédito" in D["formas"] else None, "=FormasPagamento", None),
        ("1ª parcela já foi paga?", "não", '"sim,não"', "sim: a 1ª entra como Pago; as demais como Pendente."),
    ]
    ref = {}
    for i, (rot, val, lista_dv, dica) in enumerate(campos):
        r = 5 + i
        ws.row_dimensions[r].height = 22
        put(ws, f"B{r}", rot, font=fnt(10, color=P["suave"]), fill_=CARD, align=ALIGN_L)
        put(ws, f"C{r}", val, font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_L,
            border=Border(bottom=side(P["destaque"])))
        ref[rot] = f"$C${r}"
        if lista_dv:
            dv = DataValidation(type="list", formula1=lista_dv, allow_blank=True, showErrorMessage=lista_dv.startswith('"'))
            if dica:
                dv.showInputMessage, dv.prompt = True, dica
            ws.add_data_validation(dv)
            dv.add(f"C{r}")
    ws[ref["Valor (R$)"].replace("$", "")].number_format = FMT_MOEDA
    ws[ref["Data da 1ª parcela"].replace("$", "")].number_format = FMT_DATA
    dvn = DataValidation(type="whole", operator="between", formula1="1", formula2="60", showErrorMessage=True,
                         error="Use de 1 a 60.", errorTitle="Nº de parcelas")
    ws.add_data_validation(dvn)
    dvn.add(ref["Nº de parcelas / meses"].replace("$", ""))
    modo, dsc, val, tv, nn, d1 = (ref[k] for k in ("Modo", "Descrição", "Valor (R$)", "O valor informado é", "Nº de parcelas / meses", "Data da 1ª parcela"))
    r0 = 5 + len(campos) + 1
    put(ws, f"B{r0}", "RESUMO", font=fnt(10, True, P["suave"]))
    parcela_f = f'IF({modo}="Repetir todo mês",{val},IF({tv}="Total da compra",ROUND({val}/MAX(1,{nn}),2),{val}))'
    resumo = [("Valor de cada parcela", f"={parcela_f}", FMT_MOEDA),
              ("Total", f'=IF(AND({modo}="Parcelado",{tv}="Total da compra"),{val},{parcela_f}*{nn})', FMT_MOEDA),
              ("Última parcela em", f'=IFERROR(DATE(YEAR({d1}),MONTH({d1})+{nn}-1,1),"")', FMT_COMPETENCIA)]
    for i, (rot, f, fmt_) in enumerate(resumo):
        r = r0 + 1 + i
        put(ws, f"B{r}", rot, font=fnt(10, color=P["suave"]), fill_=CARD, align=ALIGN_L)
        put(ws, f"C{r}", f, font=fnt(11, True), fill_=CARD, fmt=fmt_, align=ALIGN_L)
    put(ws, f"B{r0 + 5}", '=HYPERLINK("#Inicio!A1","⌂  Início")', font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_C)
    PC = 5   # bloco a partir da coluna E
    NPARC = 60
    for j, h in enumerate(COLS_BLOCO):
        ws.column_dimensions[CL(PC + j)].width = LARG_BLOCO[h]
        ws.cell(4, PC + j, h)
    estilo_cabecalho(ws, 4, PC, PC + len(COLS_BLOCO) - 1)
    put(ws, f"{CL(PC)}2", f'=HYPERLINK("#Lancamentos!B"&(ROWS({L}[Data])+{D["_linhas_importadas"][0]}),"✚  Ir para a primeira linha vazia de Lançamentos")',
        font=fnt(11, True, P["cabecalho_txt"]), fill_=HEAD, align=ALIGN_C)
    ws.merge_cells(f"{CL(PC)}2:{CL(PC + 4)}2")
    estilo_corpo(ws, 5, 4 + NPARC, PC, PC + len(COLS_BLOCO) - 1, fill_=CARD2)
    at_de = lambda col: f'IFERROR(INDEX(tbAtalhos[{col}],MATCH({dsc},tbAtalhos[Descrição],0))&"","")'
    esc = lambda campo, col: f'IF({ref[campo]}<>"",{ref[campo]},{at_de(col)})'
    for i in range(1, NPARC + 1):
        r = 4 + i
        on = f"{i}<={nn}"
        mes_i = f"MONTH({d1})+{i - 1}"
        tipo_i = f'IF({ref["Tipo"]}<>"",{ref["Tipo"]},IFERROR(INDEX(tbAtalhos[Tipo],MATCH({dsc},tbAtalhos[Descrição],0))&"","{tipo_padrao}"))'
        valor_i = (f'IF({modo}="Repetir todo mês",{val},IF({tv}="Total da compra",IF({i}<{nn},ROUND({val}/{nn},2),'
                   f'ROUND({val}-ROUND({val}/{nn},2)*({nn}-1),2)),{val}))')
        data_i = f'DATE(YEAR({d1}),{mes_i},MIN(DAY({d1}),DAY(EOMONTH(DATE(YEAR({d1}),{mes_i},1),0))))'
        cat_i = esc("Categoria (vazio = do favorito)", "Categoria")
        pote_i = esc("Pote (vazio = do favorito)", "Pote")
        mapa = {
            "Data": f'=IF({on},{data_i},"")', "Vencimento": f'=IF({on},{data_i},"")',
            "Descrição": f'=IF({on},{dsc},"")',
            "Valor": f'=IF({on},{valor_i},"")',
            "Tipo": f'=IF({on},{tipo_i},"")',
            "Categoria": f'=IF({on},{cat_i},"")',
            "Situação": f'=IF({on},IF(AND({i}=1,{ref["1ª parcela já foi paga?"]}="sim"),IF({tipo_i}="Receita","Recebido","Pago"),"Pendente"),"")',
            "Pote": f'=IF({on},IF({pote_i}<>"",{pote_i},IFERROR(INDEX(tbCatDespesa[Fundo padrão],MATCH({cat_i},tbCatDespesa[Categoria de gasto],0))&"","")),"")',
            "Quem": f'=IF({on},IF({esc("Quem", "Quem")}<>"",{esc("Quem", "Quem")},QuemPadrao),"")',
            "Conta": f'=IF({on},{esc("Conta / cartão", "Conta")},"")',
            "Forma de pagamento": f'=IF({on},{esc("Forma de pagamento", "Forma de pagamento")},"")',
            "Parcela": f'=IF(AND({on},{modo}="Parcelado"),"{i}/"&{nn},"")',
            "Descontar de": '=""',
        }
        for j, h in enumerate(COLS_BLOCO):
            cc = ws.cell(r, PC + j, mapa[h])
            if h in ("Data", "Vencimento"):
                cc.number_format = FMT_DATA
            elif h == "Valor":
                cc.number_format, cc.alignment = FMT_MOEDA, ALIGN_R
            elif h == "Parcela":
                cc.alignment = ALIGN_C
    ws.freeze_panes = "A5"

    # ---------------------------------------------------------------- Caderno do mês (o Painel Mensal)
    ws = ws_cad
    lc = D["_col_lanc"]
    pintar_fundo(ws, 16, 3)
    larguras(ws, {"A": 2, "B": 13.5, "C": 30, "D": 10, "E": 18, "F": 14, "G": 18, "H": 22, "I": 3, "J": 33, "K": 18, "L": 2, "N": 14, "O": 6})
    ws.column_dimensions["N"].hidden = True
    ws.column_dimensions["O"].hidden = True
    ws.row_dimensions[1].height = 10
    ws.row_dimensions[2].height = 30
    ws.row_dimensions[3].height = 28
    put(ws, "B2", "Caderno do mês", font=fnt(20, True), align=Alignment(vertical="bottom"))
    put(ws, "B3", "Receitas, gastos fixos e gastos extras do mês de trabalho — como o Painel Mensal. Pendente aparece, mas não entra no total.",
        font=fnt(9, color=P["suave"]), align=Alignment(vertical="center"))
    put(ws, "J2", "MÊS DE TRABALHO  (todo o resto segue este mês)", font=fnt(9, True, P["suave"]), align=Alignment(vertical="bottom", indent=1))
    ws.merge_cells("J2:K2")
    sel_borda = Border(bottom=side(P["destaque"], "medium"))
    put(ws, "J3", MESES[D["mes_padrao"] - 1], font=fnt(14, True, P["destaque"]), fill_=CARD2, align=ALIGN_C, border=sel_borda)
    put(ws, "K3", D["ano_padrao"], font=fnt(14, True, P["destaque"]), fill_=CARD2, align=ALIGN_C, border=sel_borda)
    wb.defined_names["MesSelecionado"] = DefinedName("MesSelecionado", attr_text="Caderno!$J$3")
    wb.defined_names["AnoSelecionado"] = DefinedName("AnoSelecionado", attr_text="Caderno!$K$3")
    dv_lista(ws, "=Meses", "J3", "Mês inválido", "Escolha um mês da lista.")
    dv_lista(ws, "=Anos", "K3", "Ano inválido", "Escolha um ano da lista (cadastre novos anos na aba Config).")
    ws.freeze_panes = "A4"

    # ---- blocos: Receitas | Gastos fixos | Gastos extras (um por tipo de gasto)
    nomes_bloco = {"Receita": "RECEITAS", "Gasto Fixo": "GASTOS FIXOS", "Gasto Extra": "GASTOS EXTRAS"}
    # linhas por bloco com folga sobre o recorde do histórico (62 fixos em set/2026, 60 extras em dez/2022)
    blocos = [("Receita", 25)] + [(t, 100 if "fixo" in t.lower() else 80) for t in tipos_gasto]
    cols_cad = [("Data", "B"), ("Descrição", "C"), ("Parcela", "D"), ("Valor", "E"), ("Vencimento", "F"), ("Situação", "G"), ("Categoria", "H")]
    r = 5
    faixas_slots = []
    for tipo, nslots in blocos:
        titulo = nomes_bloco.get(tipo, tipo.upper() + "S")
        ws.row_dimensions[r].height = 24
        area(ws, f"B{r}:H{r}", fill_=CARD, border=Border(top=side(P["destaque"])))
        put(ws, f"B{r}", titulo, font=fnt(10, True, P["destaque"]), fill_=CARD, align=ALIGN_L)
        crit_ok = f'{L}[Tipo],"{tipo}",{L}[Mês],MesNum,{L}[Ano],AnoSelecionado'
        put(ws, f"C{r}", f'=COUNTIFS({crit_ok})&" lançamento(s)"', font=fnt(9, color=P["suave"]), fill_=CARD, align=ALIGN_R)
        put(ws, f"E{r}", f'=SUMIFS({L}[Valor],{crit_ok},{CONTA})', font=fnt(11, True, P["positivo"] if tipo == "Receita" else P["texto"]),
            fill_=CARD, fmt=FMT_MOEDA, align=ALIGN_R)
        put(ws, f"F{r}", "+ previsto", font=fnt(9, color=P["suave"]), fill_=CARD, align=ALIGN_R)
        put(ws, f"G{r}", f'=SUMIFS({L}[Valor],{crit_ok},{L}[Situação],"Pendente")', font=fnt(9, True, P["alerta"]),
            fill_=CARD, fmt=FMT_MOEDA, align=ALIGN_L)
        r += 1
        for rot, colx in cols_cad:
            ws[f"{colx}{r}"] = rot
        estilo_cabecalho(ws, r, 2, 8)
        ws[f"E{r}"].alignment = Alignment(horizontal="right", vertical="center", indent=1)
        r += 1
        ini = r
        for k in range(1, nslots + 1):
            ws.row_dimensions[r].height = 19
            ws[f"N{r}"] = ArrayFormula(f"N{r}", f'=IFERROR(SMALL(IF(({L}[Tipo]="{tipo}")*({L}[Mês]=MesNum)*({L}[Ano]=AnoSelecionado),'
                                                  f'IF({L}[Data]="",0,{L}[Data])*100000+ROW({L}[Data])),{k}),0)')
            lin = f"MOD($N{r},100000)"
            for rot, colx in cols_cad:
                fonte = f"INDEX(Lancamentos!${lc[rot]}:${lc[rot]},{lin})"
                if rot in ("Data", "Vencimento", "Valor"):
                    f = f'=IF($N{r}=0,"",IF({fonte}="","",{fonte}))'
                else:
                    f = f'=IF($N{r}=0,"",{fonte}&"")'
                c_ = ws[f"{colx}{r}"]
                c_.value = f
                c_.fill, c_.font = CARD, fnt(10, color=P["suave"] if rot in ("Data", "Parcela", "Categoria") else P["texto"])
                c_.border = Border(bottom=side(P["borda"]))
                c_.alignment = ALIGN_R if rot == "Valor" else (ALIGN_C if rot in ("Parcela", "Vencimento") else ALIGN_L)
                if rot in ("Data", "Vencimento"):
                    c_.number_format = FMT_DATA
                elif rot == "Valor":
                    c_.number_format = FMT_MOEDA
            r += 1
        fim = r - 1
        faixas_slots.append((ini, fim, tipo))
        put(ws, f"B{r}", f'=IF(COUNTIFS({crit_ok})>{nslots},"… e mais "&(COUNTIFS({crit_ok})-{nslots})&" lançamento(s) deste tipo — veja em Lançamentos","")',
            font=fnt(9, color=P["suave"], italic=True))
        r += 2
    fim_cad = r
    prim, ult_s = faixas_slots[0][0], faixas_slots[-1][1]
    # possível duplicidade dentro do mês (mesma data, descrição e valor) — marca na coluna O (oculta)
    for ini, fim, _ in faixas_slots:
        for rr_ in range(ini, fim + 1):
            ws[f"O{rr_}"] = (f'=IF($C{rr_}="",0,IF(COUNTIFS($B${prim}:$B${ult_s},$B{rr_},$C${prim}:$C${ult_s},$C{rr_},'
                             f'$E${prim}:$E${ult_s},$E{rr_})>1,1,0))')
    # semáforo (legenda da Finanças Família 3.0): vermelho atrasado, amarelo vence em até 3 dias, verde pago
    for ini, fim, tipo in faixas_slots:
        rng = f"B{ini}:H{fim}"
        venc = f'IF($F{ini}<>"",$F{ini},$B{ini})'
        vermelho = fill(misturar(P["negativo"], P["card"], 0.55))
        amarelo = fill(misturar(P["alerta"], P["card"], 0.62))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND($G{ini}="Pendente",{venc}<>"",{venc}<TODAY())'], fill=vermelho, stopIfTrue=True))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND($G{ini}="Pendente",{venc}<>"",{venc}-TODAY()<=3)'], fill=amarelo, stopIfTrue=True))
        ws.conditional_formatting.add(f"E{ini}:G{fim}", FormulaRule(formula=[f'$G{ini}="Pendente"'], font=Font(color="FF" + P["alerta"], italic=True)))
        ws.conditional_formatting.add(f"G{ini}:G{fim}", FormulaRule(formula=[f'OR($G{ini}="Pago",$G{ini}="Recebido")'], font=Font(color="FF" + P["positivo"])))
        ws.conditional_formatting.add(f"G{ini}:G{fim}", FormulaRule(formula=[f'$G{ini}="Descontar Depois"'], font=Font(color="FF" + P["roxo"], bold=True)))
        if tipo == "Receita":
            ws.conditional_formatting.add(f"E{ini}:E{fim}", FormulaRule(formula=[f'$G{ini}="Recebido"'], font=Font(color="FF" + P["positivo"], bold=True)))
        ws.conditional_formatting.add(f"C{ini}:C{fim}", FormulaRule(formula=[f"$O{ini}=1"], font=Font(color="FF" + P["alerta"], bold=True)))

    # ---- painel da direita: resumo do mês, conferência, semáforo e botões
    pr = 5

    def secao_painel(texto):
        nonlocal pr
        ws.row_dimensions[pr].height = 22
        area(ws, f"J{pr}:K{pr}", fill_=CARD2, border=Border(top=side(P["destaque"])))
        put(ws, f"J{pr}", texto, font=fnt(10, True, P["destaque"]), fill_=CARD2, align=ALIGN_L)
        pr += 1

    def linha_painel(rot, f, fmt_=FMT_MOEDA, forte=False, cor=None, editavel=False):
        nonlocal pr
        ws.row_dimensions[pr].height = 20
        put(ws, f"J{pr}", rot, font=fnt(10, forte, P["texto"] if forte else P["suave"]), fill_=CARD, align=ALIGN_L,
            border=Border(bottom=side(P["borda"])))
        put(ws, f"K{pr}", f, font=fnt(11 if forte else 10, True if (forte or editavel) else False, cor or (P["destaque"] if editavel else P["texto"])),
            fill_=CARD2 if editavel else CARD, fmt=fmt_, align=ALIGN_R, border=Border(bottom=side(P["destaque"] if editavel else P["borda"])))
        pr += 1
        return f"K{pr - 1}"

    secao_painel("RESUMO DO MÊS")
    linha_painel("Recebido", "=Calc!B9", cor=P["positivo"], forte=True)
    for t in tipos_gasto:
        linha_painel(nomes_bloco.get(t, t).capitalize() if t in nomes_bloco else t,
                     f'=SUMIFS({L}[Valor],{L}[Tipo],"{t}",{CONTA},{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)')
    ref_res = linha_painel("Resultado do mês", "=Calc!B13", forte=True)
    ref_comp = linha_painel("% comprometido da renda", "=Calc!B33", fmt_=FMT_PCT)
    linha_painel("Aporte no mês", "=Calc!B30")
    linha_painel("Resgate no mês", "=Calc!B31")
    ref_livre = linha_painel("Saldo livre do mês", "=Calc!B32", forte=True)
    linha_painel("Total investido", "=Calc!B38")
    linha_painel("IR se resgatar hoje", "=Calc!B39")
    linha_painel("Investimento líquido", "=Calc!B40")
    linha_painel("A pagar (previsto, fora do total)", "=Calc!B42", cor=P["alerta"])
    linha_painel("A receber (previsto)", "=Calc!B43", cor=P["alerta"])
    for ref_ in (ref_res, ref_livre):
        ws.conditional_formatting.add(ref_, FormulaRule(formula=[f"{ref_}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
        ws.conditional_formatting.add(ref_, FormulaRule(formula=[f"{ref_}>0"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(ref_comp, FormulaRule(formula=[f"{ref_comp}>1"], font=Font(color="FF" + P["negativo"], bold=True)))
    pr += 1
    secao_painel("CONFERÊNCIA DO MÊS")
    venc_arr = f'IF({L}[Vencimento]<>"",{L}[Vencimento],{L}[Data])'
    ws[f"K{pr}"] = None
    ref_atr = linha_painel("Pendências vencidas", None, fmt_="0")
    ws[ref_atr] = ArrayFormula(ref_atr, f'=SUM(IF(({L}[Situação]="Pendente")*({L}[Mês]=MesNum)*({L}[Ano]=AnoSelecionado),IF({venc_arr}<TODAY(),1,0),0))')
    ref_sem = linha_painel("Sem categoria", f'=COUNTIFS({L}[Categoria],"Sem categoria",{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)'
                                            f'+COUNTIFS({L}[Categoria],"",{L}[Descrição],"<>",{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)', fmt_="0")
    ref_semdata = linha_painel("Lançamentos sem data (não entram em mês nenhum)",
                               f'=COUNTIFS({L}[Data],"",{L}[Competência],"",{L}[Descrição],"<>")', fmt_="0")
    ref_rev = linha_painel("A categorizar (aba Categorizar)", '=COUNTIFS(tbRevisaoCat[Descrição],"<>",tbRevisaoCat[Categoria certa],"")', fmt_="0")
    ref_ddn = linha_painel("Descontar Depois sem nome", f'=COUNTIFS({L}[Situação],"Descontar Depois",{L}[Descontar de],"",{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)', fmt_="0")
    ref_dup = linha_painel("Possíveis em dobro", f"=SUM(O{prim}:O{ult_s})/2", fmt_="0")
    linha_painel("Gastos vs. mês anterior", "=Calc!B16", fmt_=FMT_VAR)
    hc = hist_col
    busca_h = f"MATCH(DATE(AnoSelecionado,MesNum,1),Historico!$B$6:$B$1000,0)"
    ref_conf = linha_painel("Mês conferido em",
                            f'=IFERROR(IF(INDEX(Historico!${hc["Conferido em"]}$6:${hc["Conferido em"]}$1000,{busca_h})="","ainda não",'
                            f'INDEX(Historico!${hc["Conferido em"]}$6:${hc["Conferido em"]}$1000,{busca_h})),"ainda não")', fmt_=FMT_DATA)
    for ref_ in (ref_atr, ref_sem, ref_dup, ref_semdata, ref_rev, ref_ddn):
        ws.conditional_formatting.add(ref_, FormulaRule(formula=[f"{ref_}>0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(ref_conf, FormulaRule(formula=[f'{ref_conf}="ainda não"'], font=Font(color="FF" + P["alerta"], bold=True)))
    put(ws, f"J{pr}", f'=HYPERLINK("#Historico!{hc["Conferido em"]}"&IFERROR({busca_h}+5,6),"✔  Marcar o mês como conferido (Histórico) ➜")',
        font=fnt(10, True, P["destaque"]), fill_=CARD2, align=ALIGN_C)
    ws.merge_cells(f"J{pr}:K{pr}")
    ws.row_dimensions[pr].height = 22
    pr += 2
    secao_painel("SEMÁFORO")
    for txt, cor, fundo in [("Pago / recebido", P["positivo"], None),
                            ("Pendente — vence em até 3 dias", P["texto"], misturar(P["alerta"], P["card"], 0.62)),
                            ("Pendente — vencido (atrasado)", P["texto"], misturar(P["negativo"], P["card"], 0.55)),
                            ("Pendente — mais tarde (não soma no total)", P["alerta"], None),
                            ("Descontar Depois", P["roxo"], None)]:
        put(ws, f"J{pr}", "●  " + txt, font=fnt(10, True, cor), fill_=fill(fundo) if fundo else CARD, align=ALIGN_L)
        ws.merge_cells(f"J{pr}:K{pr}")
        pr += 1
    pr += 1
    for rot_, alvo in [("✚  Novo lançamento", f'"#Lancamentos!B"&(ROWS({L}[Data])+{D["_linhas_importadas"][0]})'),
                       ("✔  Contas do mês", '"#ContasMes!A1"'), ("⟳  Parcelar / repetir", '"#Parcelar!A1"'), ("⌂  Início", '"#Inicio!A1"')]:
        put(ws, f"J{pr}", f'=HYPERLINK({alvo},"{rot_}")', font=fnt(11, True, P["cabecalho_txt"]), fill_=HEAD, align=ALIGN_C)
        ws.merge_cells(f"J{pr}:K{pr}")
        ws.row_dimensions[pr].height = 24
        pr += 2
    D["_caderno"] = (faixas_slots, fim_cad)

    # ---------------------------------------------------------------- Potes (Alocação por Projeto × gasto)
    ws = ws_potes
    pintar_fundo(ws, 16, 3)
    titulo_aba(ws, "Potes", "O dinheiro guardado por finalidade. Categoria diz com o quê; pote (fundo) diz de qual dinheiro. "
               "Você digita o Guardado e o % alvo (latão); o resto se calcula. O gasto vem dos lançamentos do ano do Caderno.", 12)
    larguras(ws, {"A": 2, "B": 18, "C": 10, "D": 17, "E": 10, "F": 11, "G": 17, "H": 17, "I": 18, "J": 3, "K": 34, "L": 19})
    HP = 5
    cab_p = ["Pote", "% alvo", "Guardado", "% real", "Desvio", "Gasto no ano", "Gasto no mês", "Guardou × gastou"]
    for j, h in enumerate(cab_p):
        ws.cell(HP, 2 + j, h)
    estilo_cabecalho(ws, HP, 2, 9)
    potes = D.get("potes_saldo") or [(p, None, None) for p in D["potes"]]
    tr_p = lambda c: f"tbPotesSaldo[[#This Row],[{c}]]"
    f_real = f'=IF(SUM(tbPotesSaldo[Guardado])>0,N({tr_p("Guardado")})/SUM(tbPotesSaldo[Guardado]),0)'
    f_desv = f'={tr_p("% real")}-N({tr_p("% alvo")})'
    f_ano = (f'=SUMIFS({L}[Valor],{L}[Pote],{tr_p("Pote")},{L}[Tipo],"<>Receita",{CONTA},{L}[Ano],AnoSelecionado)')
    f_mes = (f'=SUMIFS({L}[Valor],{L}[Pote],{tr_p("Pote")},{L}[Tipo],"<>Receita",{CONTA},{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)')
    f_gxg = f'=IF(N({tr_p("Guardado")})+{tr_p("Gasto no ano")}>0,N({tr_p("Guardado")})/(N({tr_p("Guardado")})+{tr_p("Gasto no ano")}),0)'
    ultp = HP + len(potes)
    estilo_corpo(ws, HP + 1, ultp + 1, 2, 9)
    for i, (pote, alvo, guard) in enumerate(potes):
        r = HP + 1 + i
        ws.row_dimensions[r].height = 24
        vals_ = [pote, alvo, guard, f_real, f_desv, f_ano, f_mes, f_gxg]
        for j, v in enumerate(vals_):
            ws.cell(r, 2 + j, v)
        for colx, fm in (("C", FMT_PCT), ("D", FMT_MOEDA), ("E", FMT_PCT), ("F", '+0.0%;-0.0%;0.0%'), ("G", FMT_MOEDA), ("H", FMT_MOEDA), ("I", FMT_PCT)):
            ws[f"{colx}{r}"].number_format = fm
            ws[f"{colx}{r}"].alignment = ALIGN_R
        for colx in ("C", "D"):                   # campos que você digita: latão, como no Plano ("amarelo = você edita")
            ws[f"{colx}{r}"].fill = CARD2
            ws[f"{colx}{r}"].font = fnt(10, True, P["destaque"])
        ws[f"B{r}"].font = fnt(11, True)
    tp = ultp + 1
    ws.cell(tp, 2, "Total")
    ws.cell(tp, 3, "=SUBTOTAL(109,tbPotesSaldo[% alvo])")
    ws.cell(tp, 4, "=SUBTOTAL(109,tbPotesSaldo[Guardado])")
    ws.cell(tp, 5, "=SUBTOTAL(109,tbPotesSaldo[% real])")
    ws.cell(tp, 7, "=SUBTOTAL(109,tbPotesSaldo[Gasto no ano])")
    ws.cell(tp, 8, "=SUBTOTAL(109,tbPotesSaldo[Gasto no mês])")
    area(ws, f"B{tp}:I{tp}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for colx, fm in (("C", FMT_PCT), ("D", FMT_MOEDA), ("E", FMT_PCT), ("G", FMT_MOEDA), ("H", FMT_MOEDA)):
        ws[f"{colx}{tp}"].number_format, ws[f"{colx}{tp}"].alignment = fm, ALIGN_R
    nova_tabela(ws, "tbPotesSaldo", f"B{HP}:I{tp}", cab_p,
                formulas={"% real": f_real, "Desvio": f_desv, "Gasto no ano": f_ano, "Gasto no mês": f_mes, "Guardou × gastou": f_gxg},
                totais={"Pote": "Total", "% alvo": "sum", "Guardado": "sum", "% real": "sum", "Gasto no ano": "sum", "Gasto no mês": "sum"})
    ws.conditional_formatting.add(f"I{HP + 1}:I{ultp}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                                   color="FF" + BARRA_META, showValue=True))
    ws.conditional_formatting.add(f"F{HP + 1}:F{ultp}", FormulaRule(formula=[f"ABS(F{HP + 1})>0.05"], font=Font(color="FF" + P["alerta"], bold=True)))
    ws.conditional_formatting.add(f"C{tp}", FormulaRule(formula=[f"ROUND(C{tp},4)<>1"], font=Font(color="FF" + P["negativo"], bold=True)))
    dv_lista(ws, "=Potes", f"B{HP + 1}:B{ultp + 20}", msg="Use um pote cadastrado na aba Config.")
    nota_p = ("Guardado inicial = Alocação por Projeto do Plínio (Manutenção já com o valor corrigido). "
              "Atualize quando mudar." if not D.get("exemplo") else "Preencha quanto está guardado em cada pote.")
    put(ws, f"B{tp + 2}", nota_p, font=fnt(9, color=P["suave"], italic=True))
    put(ws, f"B{tp + 3}", '="Gastos do ano sem pote: "', font=fnt(9, color=P["suave"]))
    put(ws, f"D{tp + 3}", f'=SUMIFS({L}[Valor],{L}[Pote],"",{L}[Tipo],"<>Receita",{CONTA},{L}[Ano],AnoSelecionado)',
        font=fnt(9, True, P["alerta"]), fmt=FMT_MOEDA, align=ALIGN_R)
    put(ws, f"B{tp + 4}", "Barra: quanto do dinheiro de cada pote ficou guardado frente ao que foi gasto no ano (cheia = guardou muito mais do que gastou).",
        font=fnt(9, color=P["suave"], italic=True))
    # reequilíbrio (a Regra do Reequilíbrio do seu Guia — sempre como sugestão)
    put(ws, "K4", "REEQUILÍBRIO — SUGESTÃO", font=fnt(10, True, P["suave"]))
    put(ws, "K5", "Gastou algo e pagou com o dinheiro do próprio mês, sem resgatar? Some esse valor nos potes, na proporção do % alvo — "
        "em vez de tirar tudo de um pote só. É só sugestão: você aprova atualizando o Guardado.",
        font=fnt(9, color=P["suave"]), align=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells("K5:L7")
    put(ws, "K8", "Valor gasto e pago no mês", font=fnt(10, color=P["suave"]), fill_=CARD, align=ALIGN_L)
    put(ws, "L8", 1000, font=fnt(11, True, P["destaque"]), fill_=CARD2, fmt=FMT_MOEDA, align=ALIGN_R, border=Border(bottom=side(P["destaque"])))
    put(ws, "K9", "Pote", font=fnt(10, True, P["cabecalho_txt"]), fill_=HEAD, align=ALIGN_L)
    put(ws, "L9", "Somar ao pote", font=fnt(10, True, P["cabecalho_txt"]), fill_=HEAD, align=ALIGN_R)
    for i in range(len(potes)):
        r = 10 + i
        put(ws, f"K{r}", f"=B{HP + 1 + i}&\"  (\"&TEXT(C{HP + 1 + i}*100,\"0\")&\"%)  →  novo guardado \"", font=fnt(10), fill_=CARD, align=ALIGN_L)
        ws[f"K{r}"] = f'=B{HP + 1 + i}'
        put(ws, f"L{r}", f"=ROUND($L$8*N(C{HP + 1 + i}),2)", font=fnt(10, True, P["positivo"]), fill_=CARD, fmt=FMT_MOEDA, align=ALIGN_R)
    r = 10 + len(potes)
    put(ws, f"K{r}", "Novo total guardado", font=fnt(10, True), fill_=CARD2, align=ALIGN_L)
    put(ws, f"L{r}", f"=D{tp}+SUM(L10:L{r - 1})", font=fnt(10, True), fill_=CARD2, fmt=FMT_MOEDA, align=ALIGN_R)
    ws.freeze_panes = f"A{HP + 1}"

    # ---------------------------------------------------------------- Correções IA (o seu canal de pedidos)
    ws = ws_corr
    pintar_fundo(ws, 6, 3)
    titulo_aba(ws, "🤖 Correções IA", "Escreva pedidos aqui que eu implemento: um pedido por linha, na primeira linha vazia abaixo da tabela. "
               "Eu marco FEITO ✔ e explico em \"Como ficou\" — como na aba Melhorias IA da Finanças Família 3.0.", 4)
    larguras(ws, {"A": 2, "B": 58, "C": 12, "D": 86})
    corr = []
    rv = D.get("revisao") or []
    cnt = dict(n_rev=len(rv), n_lanc=sum(g["vezes"] for g in rv), n_certa=sum(1 for g in rv if g.get("certa")))
    for ped, st, como in (D.get("correcoes") or []):
        if not st:
            for padrao, st_, como_ in CORRECOES_FEITAS:
                if re.search(padrao, ped or "", re.I):
                    st, como = st_, como_.format(**cnt) if "{" in como_ else como_
                    break
        corr.append((ped, st, como))
    if not D.get("exemplo") and not any((c[0] or "").startswith(PEDIDO_JEITO_PLINIO[0][:40]) for c in corr):
        corr.append(PEDIDO_JEITO_PLINIO)
    corr = corr or [(None, None, None)]
    HCo = 5
    for j, h in enumerate(["Pedido", "Status", "Como ficou"]):
        ws.cell(HCo, 2 + j, h)
    estilo_cabecalho(ws, HCo, 2, 4)
    estilo_corpo(ws, HCo + 1, HCo + len(corr), 2, 4)
    for i, (ped, st, como) in enumerate(corr):
        r = HCo + 1 + i
        for j, v in enumerate((ped, st, como)):
            c_ = ws.cell(r, 2 + j, v)
            c_.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if j == 1 else "left", indent=0 if j == 1 else 1)
        linhas = max(len(str(ped or "")) / 55, len(str(como or "")) / 82, 1)
        ws.row_dimensions[r].height = max(20, 15 * math.ceil(linhas) + 6)
    nova_tabela(ws, "tbCorrecoes", f"B{HCo}:D{HCo + len(corr)}", ["Pedido", "Status", "Como ficou"])
    for r in range(HCo + len(corr) + 1, HCo + len(corr) + 60):
        for colx in "BCD":
            ws[f"{colx}{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    rs = f"C{HCo + 1}:C{HCo + len(corr) + 60}"
    ws.conditional_formatting.add(rs, FormulaRule(formula=[f'LEFT(C{HCo + 1},4)="FEIT"'], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(f"B{HCo + 1}:B{HCo + len(corr) + 60}",
                                  FormulaRule(formula=[f'AND(B{HCo + 1}<>"",C{HCo + 1}="")'], font=Font(color="FF" + P["alerta"], bold=True)))
    dv_lista(ws, '"FEITO ✔,Em andamento,A fazer"', f"C{HCo + 1}:C{HCo + len(corr) + 60}", msg="Use FEITO ✔, Em andamento ou A fazer.").showErrorMessage = False
    ws.freeze_panes = f"A{HCo + 1}"

    # ---------------------------------------------------------------- Orçamento
    ws = ws_orc
    pintar_fundo(ws, 10, 3)
    titulo_aba(ws, "Orçamento", "Defina a meta mensal de cada categoria de gasto. O realizado acompanha o mês de trabalho escolhido no Caderno.", 7)
    larguras(ws, {"A": 2, "B": 28, "C": 16, "D": 20, "E": 16, "F": 14, "G": 3, "H": 46})
    put(ws, "H2", '="Mês de referência: "&MesSelecionado&" / "&AnoSelecionado', font=fnt(11, True, P["destaque"]), align=Alignment(vertical="center"))
    HO = 5
    cab_orc = ["Categoria", "Meta mensal", "Realizado no mês selecionado", "Diferença", "% usado"]
    for j, h in enumerate(cab_orc):
        ws.cell(HO, 2 + j, h)
    estilo_cabecalho(ws, HO, 2, 6)
    ws.row_dimensions[HO].height = 32
    f_real = (f'=SUMIFS({L}[Valor],{L}[Categoria],tbOrcamento[[#This Row],[Categoria]],{L}[Tipo],"<>Receita",{CONTA},'
              f'{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)')
    f_dif = "=tbOrcamento[[#This Row],[Meta mensal]]-tbOrcamento[[#This Row],[Realizado no mês selecionado]]"
    f_pct = ("=IF(tbOrcamento[[#This Row],[Meta mensal]]>0,"
             "tbOrcamento[[#This Row],[Realizado no mês selecionado]]/tbOrcamento[[#This Row],[Meta mensal]],0)")
    orc_itens = list(D["orcamento"].items()) or [(None, None)]
    ult_o = HO + len(orc_itens)
    estilo_corpo(ws, HO + 1, ult_o + 1, 2, 6)
    for i, (cat, meta) in enumerate(orc_itens):
        r = HO + 1 + i
        ws.cell(r, 2, cat)
        ws.cell(r, 3, meta)
        ws.cell(r, 4, f_real)
        ws.cell(r, 5, f_dif)
        ws.cell(r, 6, f_pct)
    tot = ult_o + 1
    f_pct_tot = "=IF(SUM(tbOrcamento[Meta mensal])>0,SUM(tbOrcamento[Realizado no mês selecionado])/SUM(tbOrcamento[Meta mensal]),0)"
    ws.cell(tot, 2, "Total")
    ws.cell(tot, 3, "=SUBTOTAL(109,tbOrcamento[Meta mensal])")
    ws.cell(tot, 4, "=SUBTOTAL(109,tbOrcamento[Realizado no mês selecionado])")
    ws.cell(tot, 5, "=SUBTOTAL(109,tbOrcamento[Diferença])")
    ws.cell(tot, 6, f_pct_tot)
    area(ws, f"B{tot}:F{tot}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for r in range(HO + 1, tot + 1):
        for colx in "CDE":
            ws[f"{colx}{r}"].number_format = FMT_MOEDA
            ws[f"{colx}{r}"].alignment = ALIGN_R
        ws[f"F{r}"].number_format = FMT_PCT
        ws[f"F{r}"].alignment = ALIGN_R
    nova_tabela(ws, "tbOrcamento", f"B{HO}:F{tot}", cab_orc,
                formulas={"Realizado no mês selecionado": f_real, "Diferença": f_dif, "% usado": f_pct},
                totais={"Categoria": "Total", "Meta mensal": "sum",
                        "Realizado no mês selecionado": "sum", "Diferença": "sum"})
    tco = ws.tables["tbOrcamento"].tableColumns[4]
    tco.totalsRowFunction = "custom"
    tco.totalsRowFormula = TableFormula(attr_text=f_pct_tot.lstrip("="))
    regras_percentual(ws, f"F{HO + 1}:F{tot}", f"F{HO + 1}")
    ws.conditional_formatting.add(f"E{HO + 1}:E{tot}", FormulaRule(formula=[f"E{HO + 1}<0"], font=Font(color="FF" + P["negativo"])))
    dv_lista(ws, "=CatDespesa", f"B{HO + 1}:B{tot + 40}", msg="Use uma categoria de gasto cadastrada na aba Config.")
    put(ws, "H5", "Legenda do % usado", font=fnt(10, True))
    for i, (txt, cor) in enumerate([("até 80% — dentro do planejado", P["positivo"]),
                                    ("80% a 100% — atenção", P["alerta"]),
                                    ("acima de 100% — estourou a meta", P["negativo"])]):
        put(ws, f"H{6 + i}", "●  " + txt, font=fnt(10, color=cor))
    caixa_ajuda(ws, "H", "H", 15, [
        "Meta mensal = teto de gasto da categoria no mês.",
        "Começou pela média dos seus últimos 12 meses.",
        "Realizado = o que foi Pago no mês do Caderno.",
        "Diferença + = ainda cabe; − = estourou.",
        "% usado: verde até 80%, amarelo até 100%.",
        "Mudar a meta: digite por cima do valor.",
        "Categoria nova: Config + linha vazia aqui.",
        "O Início mostra o % do orçamento usado."])
    put(ws, "H10", D["nota_orcamento"], font=fnt(9, color=P["suave"], italic=True),
        align=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells("H10:H13")
    ws.freeze_panes = f"A{HO + 1}"

    # ---------------------------------------------------------------- Anual
    ws = ws_anual
    pintar_fundo(ws, 20, 3)
    titulo_aba(ws, "Visão anual", "Categorias × meses do ano do mês de trabalho (Caderno). Linhas vazias estão reservadas para categorias novas cadastradas na aba Config.", 17)
    put(ws, "O2", '="Ano: "&AnoSelecionado', font=fnt(14, True, P["destaque"]), align=Alignment(horizontal="right", vertical="center"))
    ws.merge_cells("O2:Q2")
    larguras(ws, {"A": 2, "B": 26, **{CL(3 + i): 11.5 for i in range(12)}, "O": 14, "P": 13, "Q": 16, "R": 2})
    HA = 5
    for j, h in enumerate(["Categoria"] + MESES_ABREV + ["Total", "Média mensal", "Tendência"]):
        ws.cell(HA, 2 + j, h)
    estilo_cabecalho(ws, HA, 2, 17)
    for colx in range(3, 17):
        ws.cell(HA, colx).alignment = Alignment(horizontal="right", vertical="center", indent=1)
    ws.cell(HA, 17).alignment = ALIGN_C

    def linha_anual(r, cat_formula, receita):
        crit = f'{L}[Tipo],"Receita"' if receita else f'{L}[Tipo],"<>Receita"'
        ws.row_dimensions[r].height = 20
        ws.cell(r, 2, cat_formula)
        for m in range(12):
            ws.cell(r, 3 + m, f'=IF($B{r}="",0,SUMIFS({L}[Valor],{L}[Categoria],$B{r},{crit},{CONTA},{L}[Mês],{m + 1},{L}[Ano],AnoSelecionado))')
        ws.cell(r, 15, f"=SUM(C{r}:N{r})")
        ws.cell(r, 16, f"=IF(Calc!$B$26>0,O{r}/Calc!$B$26,0)")
        area(ws, f"B{r}:Q{r}", fill_=CARD, font=fnt(10), border=Border(bottom=side(P["borda"])))
        area(ws, f"C{r}:P{r}", fmt=FMT_MOEDA_ZERO_TRACO, align=ALIGN_R)
        ws.cell(r, 2).alignment = ALIGN_L

    def secao(r, texto, cor):
        ws.row_dimensions[r].height = 22
        area(ws, f"B{r}:Q{r}", fill_=CARD2, font=fnt(10, True, cor))
        ws.cell(r, 2, texto)
        ws.cell(r, 2).alignment = ALIGN_L

    def linha_total(r, rotulo, ini, fim, cor):
        ws.cell(r, 2, rotulo)
        for colx in range(3, 17):
            cc = CL(colx)
            ws.cell(r, colx, f"=SUM({cc}{ini}:{cc}{fim})")
        area(ws, f"B{r}:Q{r}", fill_=CARD2, font=fnt(10, True, cor), fmt=FMT_MOEDA_ZERO_TRACO, border=Border(top=side(cor)))
        area(ws, f"C{r}:P{r}", align=ALIGN_R)
        ws.cell(r, 2).alignment = ALIGN_L
        ws.cell(r, 16, f"=IF(Calc!$B$26>0,O{r}/Calc!$B$26,0)")

    r = HA + 1
    secao(r, "RECEITAS", P["positivo"])
    rec_ini = r + 1
    for i in range(len(D["cat_receita"]) + 2):
        r += 1
        linha_anual(r, f'=IF({i + 1}<=Calc!$B$20,INDEX({rr},{i + 1}),"")', True)
    rec_fim = r
    r += 1
    tot_rec = r
    linha_total(r, "Total de receitas", rec_ini, rec_fim, P["positivo"])
    r += 2
    secao(r, "GASTOS", P["negativo"])
    des_ini = r + 1
    for i in range(len(D["cat_despesa"]) + 3):
        r += 1
        linha_anual(r, f'=IF({i + 1}<=Calc!$B$21,INDEX({rd},{i + 1}),"")', False)
    des_fim = r
    r += 1
    tot_des = r
    linha_total(r, "Total de gastos", des_ini, des_fim, P["negativo"])
    r += 2
    saldo_l = r
    ws.cell(r, 2, "Saldo do mês")
    for colx in range(3, 17):
        cc = CL(colx)
        ws.cell(r, colx, f"={cc}{tot_rec}-{cc}{tot_des}")
    area(ws, f"B{r}:Q{r}", fill_=CARD2, font=fnt(11, True), fmt=FMT_MOEDA_ZERO_TRACO, border=Border(top=side(P["destaque"])))
    area(ws, f"C{r}:P{r}", align=ALIGN_R)
    ws.cell(r, 2).alignment = ALIGN_L
    ws.row_dimensions[r].height = 22
    r += 1
    ws.cell(r, 2, "Taxa de poupança")
    for colx in range(3, 17):
        cc = CL(colx)
        ws.cell(r, colx, f"=IF({cc}{tot_rec}>0,{cc}{saldo_l}/{cc}{tot_rec},0)")
    area(ws, f"B{r}:Q{r}", fill_=CARD, font=fnt(10, color=P["suave"]), fmt=FMT_PCT)
    area(ws, f"C{r}:P{r}", align=ALIGN_R)
    ws.cell(r, 2).alignment = ALIGN_L
    ws.cell(r, 16).value = None
    ws.freeze_panes = f"C{HA + 1}"
    ws.conditional_formatting.add(f"C{des_ini}:N{des_fim}", ColorScaleRule(
        start_type="num", start_value=0, start_color="FF" + P["card"],
        end_type="max", end_color="FF" + P["heat_max"]))
    ws.conditional_formatting.add(f"C{saldo_l}:O{saldo_l}", FormulaRule(formula=[f"C{saldo_l}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    linhas_spark = list(range(rec_ini, rec_fim + 1)) + [tot_rec] + list(range(des_ini, des_fim + 1)) + [tot_des, saldo_l]

    # ---------------------------------------------------------------- Histórico (mês a mês, todos os anos)
    ws = ws_hist
    nh = len(cab_h)
    pintar_fundo(ws, nh + 3, 3)
    titulo_aba(ws, "Histórico mês a mês", "Todos os meses, com a mesma regra do caderno (só entra o que tem 'Conta no mês?' = sim). "
               "Fechamento: depois de conferir o mês no Caderno, escreva a data em \"Conferido em\" — é o seu ritual de declarar que revisou.", nh)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 12
    for j in range(1, nh):
        ws.column_dimensions[CL(2 + j)].width = 17
    HH = 5
    for j, h in enumerate(cab_h):
        ws.cell(HH, 2 + j, h)
    estilo_cabecalho(ws, HH, 2, nh + 1)
    ws.row_dimensions[HH].height = 30
    for j in range(1, nh):
        ws.cell(HH, 2 + j).alignment = Alignment(horizontal="right", vertical="center", indent=1, wrap_text=True)
    estilo_corpo(ws, HH + 1, HH + len(meses_hist), 2, nh + 1)
    hc = hist_col
    for i, (a, m) in enumerate(meses_hist):
        r = HH + 1 + i
        ws.cell(r, 2, dt.date(a, m, 1)).number_format = FMT_COMPETENCIA
        base = f'{CONTA},{L}[Mês],MONTH($B{r}),{L}[Ano],YEAR($B{r})'
        ws[f"{hc['Receitas']}{r}"] = f'=SUMIFS({L}[Valor],{L}[Tipo],"Receita",{base})'
        for t in tipos_gasto:
            ws[f"{hc[t]}{r}"] = f'=SUMIFS({L}[Valor],{L}[Tipo],"{t}",{base})'
        if "Total de gastos" in hc:
            ws[f"{hc['Total de gastos']}{r}"] = f"=SUM({hc[tipos_gasto[0]]}{r}:{hc[tipos_gasto[-1]]}{r})"
        col_tot = hc.get("Total de gastos", hc[tipos_gasto[0]])
        ws[f"{hc['Resultado']}{r}"] = f"={hc['Receitas']}{r}-{col_tot}{r}"
        ws[f"{hc['Aporte']}{r}"] = "=" + aporte_f(f"MONTH($B{r})", f"YEAR($B{r})")
        ws[f"{hc['Resgate']}{r}"] = "=" + resgate_f(f"MONTH($B{r})", f"YEAR($B{r})")
        ws[f"{hc['Saldo livre']}{r}"] = f"={hc['Resultado']}{r}-{hc['Aporte']}{r}+{hc['Resgate']}{r}"
        ws[f"{hc['% comprometido']}{r}"] = f"=IF({hc['Receitas']}{r}>0,{col_tot}{r}/{hc['Receitas']}{r},0)"
        conf = D.get("conferidos", {}).get((a, m))
        ws[f"{hc['Conferido em']}{r}"] = conf
        for h, colx in hc.items():
            c_ = ws[f"{colx}{r}"]
            if h in ("Mês", "Conferido em"):
                continue
            c_.number_format = FMT_PCT if h == "% comprometido" else FMT_MOEDA_ZERO_TRACO
            c_.alignment = ALIGN_R
        ws[f"{hc['Resultado']}{r}"].font = fnt(10, True)
        ws[f"{hc['Saldo livre']}{r}"].font = fnt(10, True)
        cc = ws[f"{hc['Conferido em']}{r}"]
        cc.number_format, cc.alignment = FMT_DATA, ALIGN_C
        cc.fill, cc.font = CARD2, fnt(10, True, P["destaque"])
        cc.border = Border(bottom=side(P["latao_fraco"] if "latao_fraco" in P else P["borda"]))
        if m == 12:
            for colx in range(2, nh + 2):
                ws.cell(r, colx).border = Border(bottom=side(P["destaque"]))
    fim_h = HH + len(meses_hist)
    for h in ("Resultado", "Saldo livre"):
        c0 = hc[h]
        ws.conditional_formatting.add(f"{c0}{HH + 1}:{c0}{fim_h}", FormulaRule(formula=[f"{c0}{HH + 1}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
        ws.conditional_formatting.add(f"{c0}{HH + 1}:{c0}{fim_h}", FormulaRule(formula=[f"{c0}{HH + 1}>0"], font=Font(color="FF" + P["positivo"], bold=True)))
    cconf = hc["Conferido em"]
    dvc = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, showErrorMessage=True,
                         error="Digite a data em que você conferiu o mês (dd/mm/aaaa).", errorTitle="Conferido em",
                         showInputMessage=True, promptTitle="Fechamento do mês",
                         prompt="Depois de conferir o mês no Caderno, digite a data (Ctrl+; põe a de hoje).")
    ws.add_data_validation(dvc)
    dvc.add(f"{cconf}{HH + 1}:{cconf}{fim_h}")
    ws.freeze_panes = f"C{HH + 1}"
    D["_hist"] = (HH, fim_h)

    # ---------------------------------------------------------------- Parcelamentos em andamento (na aba Parcelar)
    # Correção IA: a aba Metas saiu (virou Plano 1 e Plano 2); os parcelamentos ficaram junto do Parcelar.
    ws = ws_parc
    HM = 4 + NPARC + 4
    put(ws, f"{CL(PC)}{HM - 1}", "PARCELAMENTOS EM ANDAMENTO — atualize \"Parcelas pagas\" quando pagar uma parcela",
        font=fnt(10, True, P["suave"]))
    cab_d = ["Descrição", "Valor da parcela", "Parcelas pagas", "Total de parcelas", "Saldo devedor", "Data de término"]
    f_sd = "=N(tbDividas[[#This Row],[Valor da parcela]])*MAX(0,N(tbDividas[[#This Row],[Total de parcelas]])-N(tbDividas[[#This Row],[Parcelas pagas]]))"
    dividas = D["dividas"] or [(None, None, None, None, None)]
    c0 = PC
    for j, h in enumerate(cab_d):
        ws.cell(HM, c0 + j, h)
    estilo_cabecalho(ws, HM, c0, c0 + 5)
    ws.row_dimensions[HM].height = 32
    estilo_corpo(ws, HM + 1, HM + len(dividas) + 1, c0, c0 + 5)
    for i, (desc_d, parc, pagas, total, fim) in enumerate(dividas):
        r = HM + 1 + i
        for j, v in enumerate((desc_d, parc, pagas, total)):
            ws.cell(r, c0 + j, v)
        ws.cell(r, c0 + 4, f_sd)
        ws.cell(r, c0 + 5, fim)
    td = HM + len(dividas) + 1
    ws.cell(td, c0, "Total")
    ws.cell(td, c0 + 1, "=SUBTOTAL(109,tbDividas[Valor da parcela])")
    ws.cell(td, c0 + 4, "=SUBTOTAL(109,tbDividas[Saldo devedor])")
    area(ws, f"{CL(c0)}{td}:{CL(c0 + 5)}{td}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for r in range(HM + 1, td + 1):
        for j in (1, 4):
            ws.cell(r, c0 + j).number_format = FMT_MOEDA
            ws.cell(r, c0 + j).alignment = ALIGN_R
        for j in (2, 3):
            ws.cell(r, c0 + j).alignment = ALIGN_C
        ws.cell(r, c0 + 5).number_format = FMT_DATA
        ws.cell(r, c0 + 5).alignment = ALIGN_C
    nova_tabela(ws, "tbDividas", f"{CL(c0)}{HM}:{CL(c0 + 5)}{td}", cab_d, formulas={"Saldo devedor": f_sd},
                totais={"Descrição": "Total", "Valor da parcela": "sum", "Saldo devedor": "sum"})
    put(ws, f"B{r0 + 7}", f'=HYPERLINK("#Parcelar!{CL(c0)}{HM}","▼  Parcelamentos em andamento")', font=fnt(11, True, P["destaque"]),
        fill_=CARD2, align=ALIGN_C)
    caixa_ajuda(ws, "B", "C", r0 + 9, [
        "Compra parcelada sem digitar linha por linha.",
        "1. Preencha os campos em latão (acima).",
        "2. À direita saem as linhas 1/n, 2/n, 3/n…",
        "3. Copie as linhas preenchidas (Ctrl+C).",
        "4. ✚ leva à 1ª linha vazia de Lançamentos.",
        "5. Cole só os valores: Ctrl+Shift+V.",
        "Futuras = Pendente: só contam quando pagas.",
        "Repetir todo mês: aluguel, mesada, assinatura.",
        "Parcelamentos em andamento: link acima."])

    # ---------------------------------------------------------------- Investimentos (vivo: rende sozinho)
    # Correção IA: "como ficaria a inserção dos investimentos?" — uma linha por aplicação; Dias, Alíquota IR,
    # Saldo hoje, IR e Líquido se calculam (CDI × % do CDI, dias úteis, IR regressivo). As linhas do app
    # partem da posição que o app calculou (Base) e seguem rendendo a partir da data da base.
    ws = ws_inv
    cab_i = CAB_INV
    inv = D.get("investimentos") or [[None] * len(CAB_INV)]
    ni = len(cab_i)
    pintar_fundo(ws, ni + 8, 3)
    titulo_aba(ws, "Investimentos", "Uma linha por aplicação. Digite Data, Banco, % do CDI e Aplicado (latão); Saldo hoje, IR e Líquido "
               "se calculam sozinhos com o Histórico do CDI da aba Planos. Resgate total: Resgatado = sim. Resgate parcial: linha com valor negativo.", ni)
    ws.column_dimensions["A"].width = 2
    larg_i = {"Data": 12, "Banco": 11, "% do CDI": 10, "Aplicado": 15, "Dias": 7, "Isento": 8, "Resgatado": 10,
              "Alíquota IR": 10, "Saldo hoje": 15, "IR": 12, "Líquido": 15, "Vencimento": 12, "Observação": 40,
              "Base (app)": 14, "Data da base": 12}
    HI = 5
    ti = lambda c: f"tbInvestimentos[[#This Row],[{c}]]"
    pct_i = f'IF(N({ti("% do CDI")})>5,N({ti("% do CDI")})/100,N({ti("% do CDI")}))'
    base_i = f'IF(N({ti("Base (app)")})<>0,{ti("Base (app)")},N({ti("Aplicado")}))'
    d0_i = f'IF(N({ti("Data da base")})>0,{ti("Data da base")},{ti("Data")})'
    def L_cdi(t):
        """Log-índice acumulado do CDI até a data t (pela tabela tbCDI)."""
        k = f"IFERROR(MATCH({t},tbCDI[A partir de],1),1)"
        return (f"(INDEX(tbCDI[Acumulado],{k})+(NETWORKDAYS(DATE(2000,1,1),{t})-INDEX(tbCDI[Dias úteis],{k}))"
                f"*INDEX(tbCDI[Log diário],{k}))")

    f_inv = {
        "Dias": f'=IF(N({ti("Data")})=0,"",TODAY()-{ti("Data")})',
        "Alíquota IR": (f'=IF(N({ti("Data")})=0,"",IF({ti("Isento")}="sim",0,IF(TODAY()-{ti("Data")}<=180,0.225,'
                        f'IF(TODAY()-{ti("Data")}<=360,0.2,IF(TODAY()-{ti("Data")}<=720,0.175,0.15)))))'),
        # rendimento = base × e^(% do CDI × Σ ln(1+CDI diário)) nos dias úteis desde a base, com a taxa de cada período
        # do Histórico do CDI (aba Planos): mudar a Selic hoje não altera o que já rendeu.
        "Saldo hoje": (f'=IF(OR(N({ti("Data")})=0,{ti("Resgatado")}="sim"),0,IF(N({ti("Aplicado")})<0,{ti("Aplicado")},'
                       f'{base_i}*EXP({pct_i}*MAX(0,{L_cdi("TODAY()")}-{L_cdi(d0_i)}))))'),
        "IR": f'=IF(N({ti("Saldo hoje")})<=0,0,MAX(0,{ti("Saldo hoje")}-N({ti("Aplicado")}))*N({ti("Alíquota IR")}))',
        "Líquido": f'=N({ti("Saldo hoje")})-N({ti("IR")})',
    }
    for j, h in enumerate(cab_i):
        ws.cell(HI, 2 + j, h)
        ws.column_dimensions[CL(2 + j)].width = larg_i.get(h, 12)
    estilo_cabecalho(ws, HI, 2, ni + 1)
    for h in f_inv:
        ws.cell(HI, 2 + cab_i.index(h)).fill = fill(misturar(P["roxo"], P["card"], 0.55 if TEMA == "plinio_escuro" else 0.15))
    estilo_corpo(ws, HI + 1, HI + len(inv), 2, ni + 1)
    fm = {"Data": FMT_DATA, "Vencimento": FMT_DATA, "Data da base": FMT_DATA, "% do CDI": FMT_PCT, "Alíquota IR": FMT_PCT,
          "Aplicado": FMT_MOEDA, "Saldo hoje": FMT_MOEDA, "IR": FMT_MOEDA, "Líquido": FMT_MOEDA, "Base (app)": FMT_MOEDA}
    for i, row in enumerate(inv):
        r = HI + 1 + i
        for j, (h, v) in enumerate(zip(cab_i, row)):
            cc = ws.cell(r, 2 + j, f_inv[h] if h in f_inv else v)
            if h in fm:
                cc.number_format = fm[h]
                cc.alignment = ALIGN_R if fm[h] != FMT_DATA else ALIGN_C
            if h in ("Dias", "Isento", "Resgatado"):
                cc.alignment = ALIGN_C
            if h in ("Base (app)", "Data da base"):
                cc.font = fnt(9, color=P["suave"])
    for r in range(HI + len(inv) + 1, HI + len(inv) + 300):
        for j, h in enumerate(cab_i):
            if h in fm:
                ws.cell(r, 2 + j).number_format = fm[h]
    nova_tabela(ws, "tbInvestimentos", f"B{HI}:{CL(ni + 1)}{HI + len(inv)}", cab_i, formulas=f_inv)
    ci_i = {h: CL(2 + j) for j, h in enumerate(cab_i)}
    fx = lambda h: f"{ci_i[h]}{HI + 1}:{ci_i[h]}{HI + 5000}"
    dv_lista(ws, '"sim,não"', fx("Isento"), msg="sim ou não (LCI/LCA/poupança = sim).")
    dv_lista(ws, '"sim,não"', fx("Resgatado"), msg="sim quando resgatar tudo.")
    bancos = sorted({str(r[1]) for r in inv if r[1]})
    if bancos:
        dvb = DataValidation(type="list", formula1='"' + ",".join(bancos)[:250] + '"', allow_blank=True, showErrorMessage=False)
        ws.add_data_validation(dvb)
        dvb.add(fx("Banco"))
    dvp = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True, showInputMessage=True,
                         promptTitle="% do CDI", prompt="110% do CDI: digite 110% (ou 110).")
    ws.add_data_validation(dvp)
    dvp.add(fx("% do CDI"))
    dva = DataValidation(type="list", formula1="=DatasRapidas", allow_blank=True, showErrorMessage=False,
                         showInputMessage=True, promptTitle="Data", prompt="Alt+↓ e Enter = hoje. Ou digite a data da aplicação.")
    ws.add_data_validation(dva)
    dva.add(fx("Data"))
    ws.conditional_formatting.add(f"B{HI + 1}:{CL(ni + 1)}{HI + 5000}", FormulaRule(
        formula=[f'${ci_i["Resgatado"]}{HI + 1}="sim"'], font=Font(color="FF" + P["suave"])))
    cx = ni + 3   # coluna do resumo
    ws.column_dimensions[CL(cx - 1)].width = 3
    ws.column_dimensions[CL(cx)].width = 26
    ws.column_dimensions[CL(cx + 1)].width = 19
    put(ws, f"{CL(cx)}4", "CARTEIRA", font=fnt(10, True, P["suave"]))
    resumo = [("Carteira hoje (bruto)", "=SUM(tbInvestimentos[Saldo hoje])", FMT_MOEDA),
              ("IR se resgatar hoje", "=SUM(tbInvestimentos[IR])", FMT_MOEDA),
              ("Líquido", "=SUM(tbInvestimentos[Líquido])", FMT_MOEDA),
              ("Total aplicado (ativos)", '=SUMIFS(tbInvestimentos[Aplicado],tbInvestimentos[Resgatado],"<>sim")', FMT_MOEDA),
              ("Rendimento bruto", f"={CL(cx + 1)}6-{CL(cx + 1)}9", FMT_MOEDA),
              ("Rentabilidade sobre o aplicado", f"=IF({CL(cx + 1)}9>0,{CL(cx + 1)}10/{CL(cx + 1)}9,0)", FMT_PCT),
              ("CDI hoje (aba Planos)", "=CDI", "0.00%")]
    for i, (rot, f, fmt_) in enumerate(resumo):
        r = 6 + i
        ws.row_dimensions[r].height = 22
        put(ws, f"{CL(cx)}{r}", rot, font=fnt(10, color=P["suave"]), fill_=CARD, align=ALIGN_L)
        put(ws, f"{CL(cx + 1)}{r}", f, font=fnt(11, True, P["positivo"] if i in (0, 2, 4) else P["texto"]),
            fill_=CARD, fmt=fmt_, align=ALIGN_R)
    area(ws, f"{CL(cx)}5:{CL(cx + 1)}5", fill_=CARD, border=Border(top=side(P["positivo"], "thick")))
    D["_ref_invest"] = f"Investimentos!{CL(cx + 1)}8"
    caixa_ajuda(ws, CL(cx), CL(cx + 1), 15, [
        "1ª linha vazia da tabela, digite:",
        "Data (Alt+↓ = hoje) · Banco",
        "% do CDI (110%) · Aplicado",
        "Isento = sim em LCI/LCA/poupança.",
        "O resto se calcula sozinho, todo dia.",
        "Reaplicação: escreva na Observação.",
        "Resgate total: Resgatado = sim.",
        "Resgate parcial: linha negativa.",
        "Aporte do mês vem daqui (não lance",
        "de novo em Lançamentos)."],
        titulo="COMO LANÇAR UM INVESTIMENTO")
    ws.freeze_panes = f"C{HI + 1}"

    # ---------------------------------------------------------------- Plano 1 e Plano 2 (como na Finanças Família 3.0)
    PL = D["planos"]
    ws = ws_p1
    pintar_fundo(ws, 20, 3)
    titulo_aba(ws, "Planos 1 e 2 — independência e renda durável",
               f'Plano 1: quando o juro líquido dos investimentos cobrir sozinho "{PL["rotulo_renda"]}". '
               "Plano 2: salário + aluguéis + juros chegando à meta. Latão = você edita. A carteira vem da aba Investimentos.", 18)
    larguras(ws, {"A": 2, "B": 34, "C": 19, "D": 36, "E": 17, "F": 38, "G": 22, "H": 16, "I": 3,
                  "J": 30, "K": 19, "L": 16, "M": 16, "N": 18, "O": 17, "P": 16, "Q": 17, "R": 13})
    put(ws, "B4", "PREMISSAS (você edita)", font=fnt(10, True, P["suave"]))
    premissas = [
        ("CDI", "CDI ao ano (hoje)", "=INDEX(tbCDI[CDI ao ano],MATCH(TODAY(),tbCDI[A partir de],1))", "0.00%",
         "Vem do Histórico do CDI (abaixo): mudou a Selic, acrescente uma linha lá."),
        ("RendaCobrir", PL["rotulo_renda"], PL["renda"], FMT_MOEDA, "O que os juros precisam pagar por mês."),
        ("AportePlanejado", "Aporte mensal planejado", PL["aporte"], FMT_MOEDA, "Quanto pretende aplicar por mês."),
        ("AporteModo", "Aporte usado na projeção", PL.get("aporte_modo") or "Planejado", None,
         "Planejado ou Média real (3 meses, sem depósitos > R$ 30 mil)."),
        ("MetaRenda", "Meta de renda durável (Plano 2)", PL["meta"], FMT_MOEDA, "Renda por mês que não depende de trabalhar mais."),
        ("Alugueis", PL["rotulo_alugueis"], PL["alugueis"], FMT_MOEDA, "Entra na renda durável do Plano 2."),
    ]
    HPm = 5
    for j, h in enumerate(["Premissa", "Valor", "Como usar"]):
        ws.cell(HPm, 2 + j, h)
    estilo_cabecalho(ws, HPm, 2, 4)
    estilo_corpo(ws, HPm + 1, HPm + len(premissas), 2, 4)
    for i, (nome, rot, v, fmt_, dica) in enumerate(premissas):
        r = HPm + 1 + i
        ws.row_dimensions[r].height = 30
        ws.cell(r, 2, rot)
        put(ws, f"C{r}", v, font=fnt(11, True, P["destaque"]), fill_=CARD2, fmt=fmt_, align=ALIGN_R,
            border=Border(bottom=side(P["destaque"])))
        put(ws, f"D{r}", dica, font=fnt(9, color=P["suave"], italic=True), align=Alignment(vertical="center", wrap_text=True, indent=1))
        wb.defined_names[nome] = DefinedName(nome, attr_text=f"'Planos'!$C${r}")
    nova_tabela(ws, "tbPremissas", f"B{HPm}:D{HPm + len(premissas)}", ["Premissa", "Valor", "Como usar"])
    dv_lista(ws, '"Planejado,Média real"', f"C{HPm + 4}", msg="Planejado ou Média real.")
    # situação de hoje
    put(ws, "F4", "PLANO 1 — HOJE", font=fnt(10, True, P["suave"]))
    ativo = 'tbInvestimentos[Saldo hoje]'
    pct_norm = "tbInvestimentos[% do CDI]/(1+99*(tbInvestimentos[% do CDI]>5))"
    ini3 = "EOMONTH(TODAY(),-4)+1"
    fim3 = "EOMONTH(TODAY(),-1)"
    media3 = (f'SUMPRODUCT((tbInvestimentos[Data]>={ini3})*(tbInvestimentos[Data]<={fim3})*(tbInvestimentos[Aplicado]>0)'
              f'*(tbInvestimentos[Aplicado]<=30000)*ISERROR(SEARCH("reaplica",tbInvestimentos[Observação]))'
              f'*ISERROR(SEARCH("correção",tbInvestimentos[Observação]))*tbInvestimentos[Aplicado])/3')
    hoje = [
        ("P1Total", "Investido hoje (bruto)", f"=SUM({ativo})", FMT_MOEDA),
        ("P1Pct", "% do CDI médio da carteira", f"=IFERROR(SUMPRODUCT({ativo},{pct_norm})/SUM({ativo}),1)", FMT_PCT),
        ("P1IR", "IR médio real da carteira", f"=IFERROR(SUMPRODUCT({ativo},tbInvestimentos[Alíquota IR])/SUM({ativo}),0.15)", FMT_PCT),
        ("P1TaxaBruta", "Rendimento bruto ao mês", "=((1+CDI)^(1/12)-1)*P1Pct", "0.000%"),
        ("P1Taxa", "Rendimento líquido ao mês", "=P1TaxaBruta*(1-P1IR)", "0.000%"),
        ("P1Juro", "Juro líquido por mês hoje", "=P1Total*P1Taxa", FMT_MOEDA),
        ("P1Media", "Aporte médio real (3 meses)", "=" + media3, FMT_MOEDA),
        ("P1Aporte", "Aporte usado na projeção", '=IF(AporteModo="Média real",P1Media,AportePlanejado)', FMT_MOEDA),
        ("P1Alvo", "Patrimônio-alvo (juro = renda a cobrir)", "=IFERROR(RendaCobrir/P1Taxa,0)", FMT_MOEDA),
        ("P1Falta", "Falta juntar", "=MAX(0,P1Alvo-P1Total)", FMT_MOEDA),
        ("Plano1Pct", "% do caminho", "=IF(P1Alvo>0,MIN(1,P1Total/P1Alvo),0)", FMT_PCT),
    ]
    for i, (nome, rot, f, fmt_) in enumerate(hoje):
        r = 6 + i
        ws.row_dimensions[r].height = 22
        put(ws, f"F{r}", rot, font=fnt(10, color=P["suave"]), fill_=CARD, align=ALIGN_L, border=Border(bottom=side(P["borda"])))
        put(ws, f"G{r}", f, font=fnt(11, True), fill_=CARD, fmt=fmt_, align=ALIGN_R, border=Border(bottom=side(P["borda"])))
        wb.defined_names[nome] = DefinedName(nome, attr_text=f"'Planos'!$G${r}")
    area(ws, "F5:G5", fill_=CARD, border=Border(top=side(P["destaque"], "thick")))
    r_pct = 6 + len(hoje) - 1
    ws.conditional_formatting.add(f"G{r_pct}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                            color="FF" + BARRA_META, showValue=True))
    NP1 = 48
    # Histórico do CDI: cada linha vale a partir da sua data; o rendimento já conquistado usa a taxa de cada período
    # (Correção IA: "se mudar a Selic, só muda o futuro"). Acumulado = soma de ln(1+CDI diário) por dia útil.
    HCD = 21
    cdi_hist = PL.get("cdi_hist") or [(dt.date(2020, 1, 1), PL["cdi"])]
    put(ws, f"B{HCD - 1}", "HISTÓRICO DO CDI — mudou a Selic? Acrescente uma linha com a data e a taxa nova (o que já rendeu não muda)",
        font=fnt(10, True, P["suave"]))
    cab_cdi = ["A partir de", "CDI ao ano", "Dias úteis", "Log diário", "Acumulado"]
    for j, h in enumerate(cab_cdi):
        ws.cell(HCD, 2 + j, h)
    estilo_cabecalho(ws, HCD, 2, 6)
    estilo_corpo(ws, HCD + 1, HCD + len(cdi_hist), 2, 6)
    tc_ = lambda c: f"tbCDI[[#This Row],[{c}]]"
    f_cdi = {"Dias úteis": f'=NETWORKDAYS(DATE(2000,1,1),{tc_("A partir de")})',
             "Log diário": f'=LN(1+{tc_("CDI ao ano")})/252',
             "Acumulado": (f'=IF(ISNUMBER(OFFSET({tc_("Acumulado")},-1,0)),OFFSET({tc_("Acumulado")},-1,0)+'
                           f'({tc_("Dias úteis")}-OFFSET({tc_("Dias úteis")},-1,0))*OFFSET({tc_("Log diário")},-1,0),0)')}
    for i, (desde, taxa) in enumerate(cdi_hist):
        r = HCD + 1 + i
        put(ws, f"B{r}", desde, font=fnt(11, True, P["destaque"]), fill_=CARD2, fmt=FMT_DATA, align=ALIGN_C)
        put(ws, f"C{r}", taxa, font=fnt(11, True, P["destaque"]), fill_=CARD2, fmt="0.00%", align=ALIGN_R)
        for j, h in enumerate(cab_cdi[2:]):
            c_ = ws.cell(r, 4 + j, f_cdi[h])
            c_.font = fnt(9, color=P["suave"])
    for r in range(HCD + 1, HCD + len(cdi_hist) + 30):
        ws.cell(r, 2).number_format = FMT_DATA
        ws.cell(r, 3).number_format = "0.00%"
    nova_tabela(ws, "tbCDI", f"B{HCD}:F{HCD + len(cdi_hist)}", cab_cdi, formulas=f_cdi)
    HP = HCD + len(cdi_hist) + 5
    r_chega = r_pct + 2
    put(ws, f"F{r_chega}", "INDEPENDENTE EM", font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_L)
    put(ws, f"G{r_chega}", (f'=IF(P1Total>=P1Alvo,"JÁ INDEPENDENTE",IF(COUNTIF($H${HP + 1}:$H${HP + NP1},"INDEPENDENTE")=0,'
                            f'"além de {NP1} meses",INDEX($B${HP + 1}:$B${HP + NP1},COUNTIF($H${HP + 1}:$H${HP + NP1},"<>INDEPENDENTE")+1)))'),
        font=fnt(12, True, P["positivo"]), fill_=CARD2, fmt=FMT_COMPETENCIA, align=ALIGN_R)
    ws.row_dimensions[r_chega].height = 28
    put(ws, f"B{HP - 1}", f"PROJEÇÃO — {NP1} MESES (juro líquido reinvestido + aporte todo mês)", font=fnt(10, True, P["suave"]))
    cab_p1 = ["Mês", "Patrimônio no início", "Aporte", "Juro líquido", "Patrimônio no fim", "Juro ÷ renda a cobrir", "Situação"]
    for j, h in enumerate(cab_p1):
        ws.cell(HP, 2 + j, h)
    estilo_cabecalho(ws, HP, 2, 8)
    estilo_corpo(ws, HP + 1, HP + NP1, 2, 8)
    for k in range(1, NP1 + 1):
        r = HP + k
        ws.cell(r, 2, f"=DATE(YEAR(TODAY()),MONTH(TODAY())+{k},1)").number_format = FMT_COMPETENCIA
        ws.cell(r, 3, "=P1Total" if k == 1 else f"=F{r - 1}")
        ws.cell(r, 4, "=P1Aporte")
        ws.cell(r, 5, f"=C{r}*P1Taxa")
        ws.cell(r, 6, f"=C{r}+D{r}+E{r}")
        ws.cell(r, 7, f"=IF(RendaCobrir>0,E{r}/RendaCobrir,0)").number_format = FMT_PCT
        ws.cell(r, 8, f'=IF(E{r}>=RendaCobrir,"INDEPENDENTE","")')
        for colx in "CDEF":
            ws[f"{colx}{r}"].number_format = FMT_MOEDA
            ws[f"{colx}{r}"].alignment = ALIGN_R
        ws.cell(r, 2).alignment = ALIGN_C
    ws.conditional_formatting.add(f"G{HP + 1}:G{HP + NP1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                                        color="FF" + BARRA_META, showValue=True))
    ws.conditional_formatting.add(f"B{HP + 1}:H{HP + NP1}", FormulaRule(formula=[f'$H{HP + 1}="INDEPENDENTE"'],
                                                                        font=Font(color="FF" + P["positivo"], bold=True)))
    put(ws, "D14", '=HYPERLINK("#Investimentos!A1","◉  Investimentos ➜")', font=fnt(11, True, P["cabecalho_txt"]), fill_=HEAD, align=ALIGN_C)
    put(ws, "D15", '=HYPERLINK("#Inicio!A1","⌂  Início")', font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_C)

    # ---- Plano 2 (mesma aba, à direita)
    put(ws, "M4", "CARREIRA (você edita — em ordem de data)", font=fnt(10, True, P["suave"]))
    carreira = PL["carreira"] or [(None, "Hoje", 0)]
    HC2 = 5
    for j, h in enumerate(["A partir de", "Posto / etapa", "Salário"]):
        ws.cell(HC2, 13 + j, h)
    estilo_cabecalho(ws, HC2, 13, 15)
    estilo_corpo(ws, HC2 + 1, HC2 + len(carreira), 13, 15)
    for i, (desde, posto, sal) in enumerate(carreira):
        r = HC2 + 1 + i
        c1 = ws.cell(r, 13, dt.date(int(desde), 1, 1) if isinstance(desde, (int, float)) else _data(desde) if desde else None)
        c1.number_format, c1.alignment = FMT_COMPETENCIA, ALIGN_C
        ws.cell(r, 14, posto)
        put(ws, f"O{r}", sal, font=fnt(11, True, P["destaque"]), fill_=CARD2, fmt=FMT_MOEDA, align=ALIGN_R)
    nova_tabela(ws, "tbCarreira", f"M{HC2}:O{HC2 + len(carreira)}", ["A partir de", "Posto / etapa", "Salário"])
    put(ws, f"M{HC2 + len(carreira) + 2}", "Linha sem data = hoje. Promoção nova? Acrescente uma linha com a data e o salário.",
        font=fnt(9, color=P["suave"], italic=True))
    sal_em = lambda m: f"IFERROR(LOOKUP(2,1/(tbCarreira[A partir de]<={m}),tbCarreira[Salário]),INDEX(tbCarreira[Salário],1))"
    put(ws, "J4", "PLANO 2 — HOJE", font=fnt(10, True, P["suave"]))
    NP2 = 96
    hoje2 = [
        ("Salário hoje (carreira)", "=" + sal_em("TODAY()"), FMT_MOEDA),
        (PL["rotulo_alugueis"], "=Alugueis", FMT_MOEDA),
        ("Juro líquido por mês hoje", "=P1Juro", FMT_MOEDA),
        ("Renda durável hoje", "=K6+K7+K8", FMT_MOEDA),
        ("Meta", "=MetaRenda", FMT_MOEDA),
        ("% da meta", "=IF(MetaRenda>0,MIN(1,K9/MetaRenda),0)", FMT_PCT),
    ]
    for i, (rot, f, fmt_) in enumerate(hoje2):
        r = 6 + i
        put(ws, f"J{r}", rot, font=fnt(10, color=P["suave"]), fill_=CARD, align=ALIGN_L, border=Border(bottom=side(P["borda"])))
        put(ws, f"K{r}", f, font=fnt(11, True), fill_=CARD, fmt=fmt_, align=ALIGN_R, border=Border(bottom=side(P["borda"])))
    area(ws, "J5:K5", fill_=CARD, border=Border(top=side(P["destaque"], "thick")))
    ws.conditional_formatting.add("K11", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                      color="FF" + BARRA_META, showValue=True))
    put(ws, "J13", "META ATINGIDA EM", font=fnt(11, True, P["destaque"]), fill_=CARD2, align=ALIGN_L)
    put(ws, "K13", (f'=IF(K9>=MetaRenda,"JÁ ATINGIDA",IF(COUNTIF($R${HP + 1}:$R${HP + NP2},"META ✔")=0,"além de {NP2 // 12} anos",'
                    f'INDEX($J${HP + 1}:$J${HP + NP2},COUNTIF($R${HP + 1}:$R${HP + NP2},"<>META ✔")+1)))'),
        font=fnt(12, True, P["positivo"]), fill_=CARD2, fmt=FMT_COMPETENCIA, align=ALIGN_R)
    ws.row_dimensions[13].height = 28
    put(ws, f"J{HP - 1}", f"PLANO 2 — PROJEÇÃO DE {NP2} MESES", font=fnt(10, True, P["suave"]))
    cab_p2 = ["Mês", "Salário", "Aluguéis", "Patrimônio no início", "Aporte", "Juro líquido", "Renda durável", "% da meta", "Situação"]
    for j, h in enumerate(cab_p2):
        ws.cell(HP, 10 + j, h)
    estilo_cabecalho(ws, HP, 10, 18)
    estilo_corpo(ws, HP + 1, HP + NP2, 10, 18)
    for k in range(1, NP2 + 1):
        r = HP + k
        ws.cell(r, 10, f"=DATE(YEAR(TODAY()),MONTH(TODAY())+{k},1)").number_format = FMT_COMPETENCIA
        ws.cell(r, 11, "=" + sal_em(f"J{r}"))
        ws.cell(r, 12, "=Alugueis")
        ws.cell(r, 13, "=P1Total" if k == 1 else f"=M{r - 1}+N{r - 1}+O{r - 1}")
        ws.cell(r, 14, "=P1Aporte")
        ws.cell(r, 15, f"=M{r}*P1Taxa")
        ws.cell(r, 16, f"=K{r}+L{r}+O{r}")
        ws.cell(r, 17, f"=IF(MetaRenda>0,P{r}/MetaRenda,0)").number_format = FMT_PCT
        ws.cell(r, 18, f'=IF(P{r}>=MetaRenda,"META ✔","")')
        for colx in "KLMNOP":
            ws[f"{colx}{r}"].number_format = FMT_MOEDA
            ws[f"{colx}{r}"].alignment = ALIGN_R
        ws.cell(r, 10).alignment = ALIGN_C
    ws.conditional_formatting.add(f"Q{HP + 1}:Q{HP + NP2}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                                        color="FF" + BARRA_META, showValue=True))
    ws.conditional_formatting.add(f"J{HP + 1}:R{HP + NP2}", FormulaRule(formula=[f'$R{HP + 1}="META ✔"'],
                                                                        font=Font(color="FF" + P["positivo"], bold=True)))
    put(ws, f"B{HP - 1}", f"PLANO 1 — PROJEÇÃO DE {NP1} MESES (juro líquido reinvestido + aporte todo mês)", font=fnt(10, True, P["suave"]))

    # ---------------------------------------------------------------- Acertos (Descontar Depois, por pessoa)
    # Correção IA: "quando coloco Descontar Depois, preciso colocar o nome da pessoa, pra fazer o cálculo automático".
    # Tudo aqui vem dos lançamentos do mês de trabalho (Caderno) com Situação = Descontar Depois e o nome em "Descontar de".
    ws = ws_ac
    lc_ac = D["_col_lanc"]
    pintar_fundo(ws, 16, 3)
    titulo_aba(ws, "Acertos — Descontar Depois", "Tudo o que você lançou como Descontar Depois no mês de trabalho do Caderno, somado por pessoa "
               "(coluna \"Descontar de\" em Lançamentos). Quando a pessoa pagar, registre em Recebidos.", 14)
    larguras(ws, {"A": 2, "B": 18, "C": 26, "D": 17, "E": 15, "F": 16, "G": 3, "H": 13, "I": 16, "J": 14, "K": 13, "L": 13,
                  "M": 32, "N": 16, "O": 2, "P": 12})
    ws.column_dimensions["P"].hidden = True
    put(ws, "B4", '="POR PESSOA — "&UPPER(MesSelecionado)&" / "&AnoSelecionado', font=fnt(10, True, P["suave"]))
    cab_a = ["Pessoa", "Descontar Depois no mês", "A cobrar no mês", "Já recebido", "Falta"]
    pessoas_ac = D.get("acertos_pessoas") or [None]
    rateios = {k.lower(): (k, v) for k, v in (D.get("rateios") or {}).items()}
    HAc = 5
    for j, h in enumerate(cab_a):
        ws.cell(HAc, 2 + j, h)
    estilo_cabecalho(ws, HAc, 2, 6)
    ws.row_dimensions[HAc].height = 36
    estilo_corpo(ws, HAc + 1, HAc + len(pessoas_ac), 2, 6)
    ta = lambda c: f"tbAcertos[[#This Row],[{c}]]"
    mes_ini, mes_fim = "DATE(AnoSelecionado,MesNum,1)", "EOMONTH(DATE(AnoSelecionado,MesNum,1),0)"
    f_dd = (f'=IF({ta("Pessoa")}="",0,SUMIFS({L}[Valor],{L}[Situação],"Descontar Depois",{L}[Descontar de],{ta("Pessoa")},'
            f'{L}[Mês],MesNum,{L}[Ano],AnoSelecionado))')
    f_rec = (f'=IF({ta("Pessoa")}="",0,SUMIFS(tbRecebidos[Valor],tbRecebidos[Pessoa],{ta("Pessoa")},'
             f'tbRecebidos[Mês de referência],">="&{mes_ini},tbRecebidos[Mês de referência],"<="&{mes_fim}))')
    f_falta = f'=N({ta("A cobrar no mês")})-N({ta("Já recebido")})'
    rateio_de = {}
    for i, pessoa in enumerate(pessoas_ac):
        r = HAc + 1 + i
        put(ws, f"B{r}", pessoa, font=fnt(10, True, P["destaque"]), fill_=CARD2)
        ws.cell(r, 3, f_dd)
        ws.cell(r, 4, f"={ta('Descontar Depois no mês')}")
        if pessoa and pessoa.lower() in rateios:
            rateio_de[pessoa] = r
        ws.cell(r, 5, f_rec)
        ws.cell(r, 6, f_falta)
        for colx in range(3, 7):
            ws.cell(r, colx).number_format = FMT_MOEDA
            ws.cell(r, colx).alignment = ALIGN_R
        ws.cell(r, 6).font = fnt(11, True, P["alerta"])
    fim_a = HAc + len(pessoas_ac)
    nova_tabela(ws, "tbAcertos", f"B{HAc}:F{fim_a}", cab_a,
                formulas={"Descontar Depois no mês": f_dd, "A cobrar no mês": f"={ta('Descontar Depois no mês')}",
                          "Já recebido": f_rec, "Falta": f_falta})
    wb.defined_names["PessoasDescontar"] = DefinedName(
        "PessoasDescontar", attr_text=f"OFFSET(Acertos!$B${HAc + 1},0,0,MAX(1,SUMPRODUCT(--(LEN(Acertos!$B${HAc + 1}:$B${HAc + 200})>0))),1)")
    r = fim_a + 1
    put(ws, f"B{r}", "Descontar Depois sem nome no mês", font=fnt(9, color=P["suave"]), align=ALIGN_L)
    ws.merge_cells(f"B{r}:C{r}")
    put(ws, f"D{r}", f'=SUMIFS({L}[Valor],{L}[Situação],"Descontar Depois",{L}[Descontar de],"",{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)',
        font=fnt(10, True), fmt=FMT_MOEDA, align=ALIGN_R)
    ws.conditional_formatting.add(f"D{r}", FormulaRule(formula=[f"D{r}>0"], font=Font(color="FF" + P["negativo"], bold=True)))
    put(ws, f"B{r + 1}", "Pessoa nova? Escreva o nome na linha vazia logo abaixo da tabela e use o mesmo nome em \"Descontar de\".",
        font=fnt(9, color=P["suave"], italic=True))
    D["_ref_receber"] = "SUM(tbAcertos[Falta])"

    # detalhe do mês (mesma técnica do Caderno)
    DT = r + 4
    put(ws, f"B{DT - 1}", "DESCONTAR DEPOIS DO MÊS — lançamento a lançamento", font=fnt(10, True, P["suave"]))
    for j, h in enumerate(["Data", "Descrição", "Valor", "Descontar de", "Parcela"]):
        ws.cell(DT, 2 + j, h)
    estilo_cabecalho(ws, DT, 2, 6)
    NDT = 60
    for k in range(1, NDT + 1):
        rr = DT + k
        ws[f"P{rr}"] = ArrayFormula(f"P{rr}", f'=IFERROR(SMALL(IF(({L}[Situação]="Descontar Depois")*({L}[Mês]=MesNum)*({L}[Ano]=AnoSelecionado),'
                                               f'IF({L}[Data]="",0,{L}[Data])*100000+ROW({L}[Data])),{k}),0)')
        lin = f"MOD($P{rr},100000)"
        for j, rot in enumerate(["Data", "Descrição", "Valor", "Descontar de", "Parcela"]):
            fonte = f"INDEX(Lancamentos!${lc_ac[rot]}:${lc_ac[rot]},{lin})"
            c_ = ws.cell(rr, 2 + j, f'=IF($P{rr}=0,"",IF({fonte}="","",{fonte}))' if rot in ("Data", "Valor") else f'=IF($P{rr}=0,"",{fonte}&"")')
            c_.fill, c_.font, c_.border = CARD, fnt(10), Border(bottom=side(P["borda"]))
            c_.alignment = ALIGN_R if rot == "Valor" else ALIGN_L
            if rot == "Data":
                c_.number_format = FMT_DATA
            if rot == "Valor":
                c_.number_format = FMT_MOEDA
        ws.row_dimensions[rr].height = 18
    ws.conditional_formatting.add(f"E{DT + 1}:E{DT + NDT}", FormulaRule(formula=[f'AND($C{DT + 1}<>"",$E{DT + 1}="")'],
                                                                       fill=fill(misturar(P["negativo"], P["card"], 0.55))))

    # Recebidos
    put(ws, "H4", "RECEBIDOS (quando a pessoa pagar)", font=fnt(10, True, P["suave"]))
    cab_r = ["Data", "Pessoa", "Valor", "Referente a (mês)", "Mês de referência"]
    recs = D.get("recebidos") or [[None] * 4]
    for j, h in enumerate(cab_r):
        ws.cell(HAc, 8 + j, h)
    estilo_cabecalho(ws, HAc, 8, 12)
    estilo_corpo(ws, HAc + 1, HAc + len(recs), 8, 12)
    tr_ = lambda c: f"tbRecebidos[[#This Row],[{c}]]"
    f_ref = (f'=IF(N({tr_("Referente a (mês)")})>0,DATE(YEAR({tr_("Referente a (mês)")}),MONTH({tr_("Referente a (mês)")}),1),'
             f'IF(N({tr_("Data")})>0,DATE(YEAR({tr_("Data")}),MONTH({tr_("Data")}),1),""))')
    for i, rec in enumerate(recs):
        rr = HAc + 1 + i
        for j, v in enumerate(rec[:4]):
            ws.cell(rr, 8 + j, v)
        ws.cell(rr, 12, f_ref)
    for rr in range(HAc + 1, HAc + len(recs) + 300):
        ws.cell(rr, 8).number_format = FMT_DATA
        ws.cell(rr, 10).number_format = FMT_MOEDA
        ws.cell(rr, 11).number_format = FMT_COMPETENCIA
        ws.cell(rr, 12).number_format = FMT_COMPETENCIA
    nova_tabela(ws, "tbRecebidos", f"H{HAc}:L{HAc + len(recs)}", cab_r, formulas={"Mês de referência": f_ref})
    dv_lista(ws, "=PessoasDescontar", f"I{HAc + 1}:I{HAc + 400}", msg="Escolha uma pessoa da tabela ao lado.")
    dvr = DataValidation(type="list", formula1="=DatasRapidas", allow_blank=True, showErrorMessage=False,
                         showInputMessage=True, promptTitle="Data", prompt="Alt+↓ e Enter = hoje.")
    ws.add_data_validation(dvr)
    dvr.add(f"H{HAc + 1}:H{HAc + 400}")
    put(ws, f"H{HAc + len(recs) + 2}", "Referente a (mês) vazio = o mês da data do recebimento.", font=fnt(9, color=P["suave"], italic=True))

    # Rateio (a mesma conta do app, com os gastos do mês de quem divide a casa)
    linha = 5
    for pessoa, rr_p in rateio_de.items():
        _, vals = rateios[pessoa.lower()]
        put(ws, f"M{linha - 1}", f"RATEIO — {pessoa.upper()}", font=fnt(10, True, P["suave"]))
        ini = linha
        passos = [("Aluguel dele (digite)", vals.get("aluguel dele")),
                  ("Aluguel dela (digite)", vals.get("aluguel dela")),
                  ("Gastos da casa (Descontar Depois)", f"=C{rr_p}"),
                  ("Lucro", f"=N{ini}+N{ini + 1}-N{ini + 2}"),
                  ("Metade do lucro", f"=N{ini + 3}/2"),
                  ("Total a cobrar", f"=N{ini + 1}-N{ini + 4}")]
        for j, (rot, v) in enumerate(passos):
            rr = ini + j
            final = j == len(passos) - 1
            put(ws, f"M{rr}", rot, font=fnt(10, final, P["texto"] if final else P["suave"]), fill_=CARD2 if final else CARD, align=ALIGN_L)
            put(ws, f"N{rr}", v, font=fnt(11 if final else 10, final, P["alerta"] if final else (P["destaque"] if j < 2 else P["texto"])),
                fill_=CARD2 if final else CARD, fmt=FMT_MOEDA, align=ALIGN_R)
        ws.cell(rr_p, 4).value = f"=N{ini + len(passos) - 1}"
        put(ws, f"M{ini + len(passos)}", f"A cobrar de {pessoa} = este total (em vez da soma simples).", font=fnt(9, color=P["suave"], italic=True))
        linha = ini + len(passos) + 3
    ws.freeze_panes = f"A{HAc + 1}"

    # Dashboard retirado (Correção IA: "com a aba Caderno não precisa a aba Dashboard").
    # A grade de 6 cartões continua servindo à aba Início.
    card_cols = []
    colx = 2
    for k in range(6):
        card_cols.append((colx, colx + 2))
        colx += 4
    ult_col = card_cols[-1][1] + 1
    c5, c6 = card_cols[4], card_cols[5]

    # ---------------------------------------------------------------- Início (central da família)
    ws = ws_ini
    pintar_fundo(ws, ult_col + 6, 45)
    ws.column_dimensions["A"].width = 3
    for c0, c1 in card_cols:
        for cc in range(c0, c1 + 1):
            ws.column_dimensions[CL(cc)].width = LARG_CARD
        ws.column_dimensions[CL(c1 + 1)].width = 2.5
    ws.column_dimensions[CL(ult_col)].width = 3
    ws.sheet_view.zoomScale = 90
    for rrow, h in {1: 12, 2: 26, 3: 28, 4: 14}.items():
        ws.row_dimensions[rrow].height = h
    put(ws, "B2", "Central da Família", font=fnt(20, True), align=Alignment(vertical="bottom"))
    put(ws, "B3", "Tudo do dia a dia num lugar só: lance, confira as contas do mês e acompanhe o que vence.",
        font=fnt(9, color=P["suave"]), align=Alignment(vertical="center"))
    put(ws, f"{CL(c5[0])}2", "MÊS DE TRABALHO", font=fnt(9, True, P["suave"]), align=Alignment(vertical="bottom", indent=1))
    mesclar(ws, f"{CL(c5[0])}3:{CL(c5[1])}3", '=MesSelecionado&" / "&AnoSelecionado', fill_=CARD2, font=fnt(13, True, P["destaque"]),
            align=ALIGN_C, border=Border(bottom=side(P["destaque"], "medium")))
    mesclar(ws, f"{CL(c6[0])}3:{CL(c6[1])}3", '=HYPERLINK("#Caderno!J3","trocar no Caderno ➜")', fill_=CARD,
            font=fnt(10, True, P["suave"]), align=ALIGN_C)
    ws.freeze_panes = "A5"
    # Botões
    prox = f'"#Lancamentos!B"&(ROWS({L}[Data])+{D["_linhas_importadas"][0]})'
    botoes = [(f'=HYPERLINK({prox},"✚  Novo lançamento")', P["destaque"]),
              ('=HYPERLINK("#Caderno!A1","📒  Caderno do mês")', P["positivo"]),
              ('=HYPERLINK("#ContasMes!A1","✔  Contas do mês")', P["azul"]),
              ('=HYPERLINK("#Parcelar!A1","⟳  Parcelar / repetir")', P["roxo"]),
              ('=HYPERLINK("#Potes!A1","◉  Potes")', misturar(P["positivo"], P["card"], 0.35)),
              ("=HYPERLINK(\"#'Planos'!A1\",\"◎  Plano 1 e 2\")", misturar(P["azul"], P["card"], 0.35))]
    for rrow in (5, 6, 7):
        ws.row_dimensions[rrow].height = 16
    for (c0, c1), (f, cor) in zip(card_cols, botoes):
        mesclar(ws, f"{CL(c0)}5:{CL(c1)}7", f, fill_=fill(cor), font=fnt(12, True, P["botao_txt"]), align=ALIGN_C)
    ws.row_dimensions[8].height = 14
    # Status do mês
    receber = D.get("_ref_receber") or "0"
    invest = D.get("_ref_invest") or "0"
    resumo_c = D["_contas_resumo"]
    status = [
        ("SALDO LIVRE DO MÊS", "=Calc!B32", FMT_MOEDA, "Resultado", "=Calc!B13", FMT_MOEDA, P["destaque"]),
        ("CONTAS FIXAS LANÇADAS", f"=ContasMes!{resumo_c[0]}", None, "Falta", f"=ContasMes!{resumo_c[1]}", FMT_MOEDA, P["positivo"]),
        ("A PAGAR NO MÊS", "=Calc!B28", FMT_MOEDA, "Atrasadas", "=Calc!B29", "0", P["alerta"]),
        ("% COMPROMETIDO DA RENDA", "=Calc!B33", FMT_PCT, "Orçam.", "=Calc!B18", FMT_PCT, P["roxo"]),
        ("A RECEBER (ACERTOS)", f"={receber}", FMT_MOEDA, "Parcelas", "=SUM(tbDividas[Saldo devedor])", FMT_MOEDA, P["alerta"]),
        ("INVESTIMENTOS (LÍQUIDO)", f"={invest}", FMT_MOEDA, "Plano 1", "=Plano1Pct", FMT_PCT, P["positivo"]),
    ]
    for rrow, h in {9: 8, 10: 20, 11: 34, 12: 18, 13: 10, 14: 18}.items():
        ws.row_dimensions[rrow].height = h
    st_cells = []
    for (c0, c1), (rot, f, fmt, sub, fsub, fmtsub, cor) in zip(card_cols, status):
        a, b = CL(c0), CL(c1)
        area(ws, f"{a}9:{b}13", fill_=CARD)
        area(ws, f"{a}9:{b}9", border=Border(top=side(cor, "thick")))
        mesclar(ws, f"{a}10:{b}10", rot, fill_=CARD, font=fnt(9, True, P["suave"]), align=Alignment(horizontal="left", vertical="center", indent=1))
        mesclar(ws, f"{a}11:{b}11", f, fill_=CARD, font=fnt(KPI_PT - 2, True), fmt=fmt, align=Alignment(horizontal="left", vertical="center", indent=1))
        put(ws, f"{a}12", sub, font=fnt(9, color=P["suave"]), fill_=CARD, align=Alignment(horizontal="left", vertical="center", indent=1))
        mesclar(ws, f"{CL(c0 + 1)}12:{b}12", fsub, fill_=CARD, font=fnt(9, True, P["suave"]), fmt=fmtsub, align=Alignment(horizontal="right", vertical="center", indent=1))
        st_cells.append(f"{a}11")
    ws.conditional_formatting.add(st_cells[0], FormulaRule(formula=[f"{st_cells[0]}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(st_cells[0], FormulaRule(formula=[f"{st_cells[0]}>=0"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(st_cells[2], FormulaRule(formula=[f"{st_cells[2]}>0"], font=Font(color="FF" + P["alerta"], bold=True)))
    o = st_cells[3]
    ws.conditional_formatting.add(o, FormulaRule(formula=[f"{o}>1"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(o, FormulaRule(formula=[f"AND({o}>0.8,{o}<=1)"], font=Font(color="FF" + P["alerta"], bold=True)))
    ws.conditional_formatting.add(o, FormulaRule(formula=[f"{o}<=0.8"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add("G12:H12", FormulaRule(formula=["G12>0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add("K12:L12", FormulaRule(formula=["K12>0"], font=Font(color="FF" + P["negativo"], bold=True)))
    # Listas: próximos vencimentos / últimos lançamentos
    T = 15
    ws.row_dimensions[T].height = 26
    ws.row_dimensions[T + 1].height = 24
    put(ws, f"B{T}", "Próximos vencimentos e atrasados", font=fnt(12, True), align=Alignment(vertical="center"))
    put(ws, f"N{T}", "Últimos lançamentos", font=fnt(12, True), align=Alignment(vertical="center"))
    for rot, a, b in [("Vence", "B", "C"), ("Descrição", "D", "H"), ("Valor", "I", "J"), ("Status", "K", "L")]:
        mesclar(ws, f"{a}{T + 1}:{b}{T + 1}", rot, fill_=HEAD, font=fnt(10, True, P["cabecalho_txt"]), align=ALIGN_R if rot == "Valor" else ALIGN_L)
    for rot, a, b in [("Data", "N", "O"), ("Descrição", "P", "T"), ("Categoria", "U", "V"), ("Valor", "W", "X")]:
        mesclar(ws, f"{a}{T + 1}:{b}{T + 1}", rot, fill_=HEAD, font=fnt(10, True, P["cabecalho_txt"]), align=ALIGN_R if rot == "Valor" else ALIGN_L)
    bb = Border(bottom=side(P["borda"]))
    for k in range(8):
        r = T + 2 + k
        cr = 2 + k
        ws.row_dimensions[r].height = 21
        mesclar(ws, f"B{r}:C{r}", f"=Calc!AW{cr}", fill_=CARD, font=fnt(10, color=P["suave"]), fmt=FMT_DATA, align=ALIGN_L, border=bb)
        mesclar(ws, f"D{r}:H{r}", f"=Calc!AX{cr}", fill_=CARD, font=fnt(10), align=ALIGN_L, border=bb)
        mesclar(ws, f"I{r}:J{r}", f"=Calc!AY{cr}", fill_=CARD, font=fnt(10, True), fmt=FMT_MOEDA, align=ALIGN_R, border=bb)
        mesclar(ws, f"K{r}:L{r}", f"=Calc!AZ{cr}", fill_=CARD, font=fnt(10, True, P["suave"]), align=ALIGN_L, border=bb)
        mesclar(ws, f"N{r}:O{r}", f"=Calc!BC{cr}", fill_=CARD, font=fnt(10, color=P["suave"]), fmt=FMT_DATA, align=ALIGN_L, border=bb)
        mesclar(ws, f"P{r}:T{r}", f"=Calc!BD{cr}", fill_=CARD, font=fnt(10), align=ALIGN_L, border=bb)
        mesclar(ws, f"U{r}:V{r}", f"=Calc!BF{cr}", fill_=CARD, font=fnt(10, color=P["suave"]), align=ALIGN_L, border=bb)
        mesclar(ws, f"W{r}:X{r}", f"=Calc!BE{cr}", fill_=CARD, font=fnt(10, True), fmt=FMT_MOEDA, align=ALIGN_R, border=bb)
    fim_l = T + 9
    ws.conditional_formatting.add(f"K{T + 2}:L{fim_l}", FormulaRule(formula=[f'$K{T + 2}="Atrasada"'], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(f"K{T + 2}:L{fim_l}", FormulaRule(formula=[f'$K{T + 2}="Esta semana"'], font=Font(color="FF" + P["alerta"], bold=True)))
    ws.conditional_formatting.add(f"W{T + 2}:X{fim_l}", FormulaRule(formula=[f'Calc!$BG{2}="Receita"'], font=Font(color="FF" + P["positivo"], bold=True)))
    # Guia rápido
    G0 = fim_l + 2
    put(ws, f"B{G0}", "Lançar em 10 segundos", font=fnt(12, True))
    passos = ["1.  ✚ Novo lançamento (botão acima) leva à primeira linha vazia.",
              "2.  Data: Alt+↓ e Enter põe HOJE (ou Ctrl+;)  →  Tab.",
              "3.  Descrição: escolha um favorito (Alt+↓) ou digite  →  Tab.",
              "4.  Valor  →  Enter. Tipo, Categoria, Situação e Pote se preenchem sozinhos.",
              "5.  Conta futura? Troque a Situação para Pendente; ela aparece em Próximos vencimentos.",
              "Compra parcelada ou lançamento que se repete? Use  ⟳ Parcelar / repetir."]
    for i, t in enumerate(passos):
        put(ws, f"B{G0 + 1 + i}", t, font=fnt(10, color=P["texto"] if i < 5 else P["suave"], italic=i == 5))
    put(ws, f"N{G0}", "Rotina da família", font=fnt(12, True))
    rotina = ["Toda semana: lançar os gastos (5 min) e olhar o 📒 Caderno do mês.",
              "Quando pagar uma conta: trocar Pendente ➜ Pago em Lançamentos.",
              "Início do mês: trocar o Mês de trabalho no Caderno e colar o bloco de ✔ Contas do mês.",
              "Fim do mês: Conferência do mês no Caderno ➜ data em \"Conferido em\" no Histórico.",
              "Pedidos pra IA: aba 🤖 Correções IA. Potes: atualize o Guardado quando aplicar.",
              "Atalhos úteis: Ctrl+D copia a célula de cima • Alt+↓ abre a lista • Ctrl+↓ vai ao fim."]
    for i, t in enumerate(rotina):
        put(ws, f"N{G0 + 1 + i}", t, font=fnt(10, color=P["texto"] if i < 5 else P["suave"], italic=i == 5))
    ws.print_area = f"A1:{CL(ult_col)}{G0 + 7}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = ws.page_margins.top = ws.page_margins.bottom = 0.3

    for w in (ws_cad, ws_lanc, ws_rev, ws_contas, ws_parc, ws_orc, ws_potes, ws_anual, ws_hist, ws_p1, ws_p2, ws_inv, ws_ac, ws_at, ws_corr, ws_cfg):
        if w is None:
            continue
        w.page_setup.orientation = "landscape"
        w.page_setup.paperSize = w.PAPERSIZE_A4
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.page_margins.left = w.page_margins.right = 0.3

    cores_guia = [(ws_ini, P["destaque"]), (ws_cad, P["positivo"]), (ws_contas, P["azul"]), (ws_parc, P["roxo"]), (ws_at, P["roxo"]),
                  (ws_potes, P["positivo"]), (ws_corr, P["alerta"]),
                  (ws_rev, P["negativo"]), (ws_p1, P["destaque"]), (ws_lanc, P["positivo"]), (ws_orc, P["alerta"]), (ws_anual, P["roxo"]),
                  (ws_hist, P["roxo"]), (ws_inv, P["positivo"]), (ws_ac, P["alerta"]),
                  (ws_cfg, P["suave"]), (ws_calc, P["borda"])]
    for w, cor in cores_guia:
        if w is not None:
            w.sheet_properties.tabColor = cor

    # Varredura final: toda célula sem preenchimento/cor recebe o fundo e a cor de texto do tema
    for w in wb.worksheets:
        if w is ws_calc:
            continue
        for row in w.iter_rows():
            for cell in row:
                if cell.fill is None or cell.fill.fill_type is None:
                    cell.fill = BG
                if cell.font is None or cell.font.color is None:
                    f = cell.font
                    cell.font = Font(name=FONTE, size=f.sz or 10, bold=f.b, italic=f.i, color="FF" + P["texto"])

    # Correção IA: abas que não são de preenchimento nem de visualização ficam ocultas (Config e Calc).
    # Para ver: botão direito numa aba → Reexibir.
    ws_cfg.sheet_state = "hidden"
    wb.active = 0
    for w in wb.worksheets:
        w.sheet_view.tabSelected = w is ws_ini
    wb.calculation.fullCalcOnLoad = True
    wb.save(caminho)
    ini_i, fim_i = D["_linhas_importadas"]
    ignorar = " ".join(f"{D['_col_lanc'][h]}{ini_i}:{D['_col_lanc'][h]}{fim_i}" for h in COLS_AUTO)
    pos_processar(caminho, {"Anual": ("C", "N", "Q", linhas_spark)}, {"Lancamentos": ignorar})


# =============================================================================
# PÓS-PROCESSAMENTO DO XML (recursos que o openpyxl não expõe)
#  - barras de dados no estilo Excel 2010+ (preenchimento sólido, sem degradê)
#  - minigráficos (sparklines) na aba Anual
# =============================================================================
NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
NS_XM = "http://schemas.microsoft.com/office/excel/2006/main"


def _xml_sparklines(aba, col_ini, col_fim, col_dest, linhas):
    sp = "".join(f"<x14:sparkline><xm:f>{aba}!{col_ini}{r}:{col_fim}{r}</xm:f><xm:sqref>{col_dest}{r}</xm:sqref></x14:sparkline>"
                 for r in linhas)
    cor = lambda h: f'rgb="FF{h}"'
    return (f'<ext uri="{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}" xmlns:x14="{NS_X14}">'
            f'<x14:sparklineGroups xmlns:xm="{NS_XM}">'
            '<x14:sparklineGroup displayEmptyCellsAs="gap" markers="1" high="1" lineWeight="1.25">'
            f'<x14:colorSeries {cor(P["destaque"])}/><x14:colorNegative {cor(P["negativo"])}/>'
            f'<x14:colorAxis {cor(P["suave"])}/><x14:colorMarkers {cor(P["destaque"])}/>'
            f'<x14:colorFirst {cor(P["destaque"])}/><x14:colorLast {cor(P["destaque"])}/>'
            f'<x14:colorHigh {cor(P["alerta"])}/><x14:colorLow {cor(P["destaque"])}/>'
            f'<x14:sparklines>{sp}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>')


def _databars_2010(xml, seq):
    """Liga cada <cfRule type="dataBar"> a uma definição x14 com preenchimento sólido."""
    defs = []

    def por_bloco(m):
        sqref, corpo = m.group(1), m.group(2)

        def por_regra(mr):
            regra = mr.group(0)
            vals = re.findall(r'<cfvo type="num" val="([^"]+)"', regra)
            guid = "{%08X-0000-4000-8000-%012X}" % (0xDA7ABA00 + seq[0], seq[0])
            seq[0] += 1
            defs.append(
                f'<x14:conditionalFormatting xmlns:xm="{NS_XM}"><x14:cfRule type="dataBar" id="{guid}">'
                '<x14:dataBar minLength="0" maxLength="100" border="0" gradient="0">'
                f'<x14:cfvo type="num"><xm:f>{vals[0]}</xm:f></x14:cfvo>'
                f'<x14:cfvo type="num"><xm:f>{vals[1]}</xm:f></x14:cfvo>'
                f'<x14:negativeFillColor rgb="FF{P["negativo"]}"/><x14:axisColor rgb="FF{P["suave"]}"/>'
                f'</x14:dataBar></x14:cfRule><xm:sqref>{sqref}</xm:sqref></x14:conditionalFormatting>')
            ext = (f'<extLst><ext uri="{{B025F937-C7B1-47D3-B67F-A62EFF666E3E}}" xmlns:x14="{NS_X14}">'
                   f'<x14:id>{guid}</x14:id></ext></extLst>')
            return regra.replace("</cfRule>", ext + "</cfRule>")

        corpo = re.sub(r'<cfRule type="dataBar".*?</cfRule>', por_regra, corpo, flags=re.S)
        return f'<conditionalFormatting sqref="{sqref}">{corpo}</conditionalFormatting>'

    xml = re.sub(r'<conditionalFormatting sqref="([^"]+)">(.*?)</conditionalFormatting>', por_bloco, xml, flags=re.S)
    if not defs:
        return xml, ""
    return xml, (f'<ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" xmlns:x14="{NS_X14}">'
                 f'<x14:conditionalFormattings>{"".join(defs)}</x14:conditionalFormattings></ext>')


def pos_processar(caminho, sparklines, ignorar_calc=None):
    """ignorar_calc = {aba: sqref}: esconde o aviso "fórmula inconsistente com a coluna" nas linhas importadas."""
    ignorar_calc = ignorar_calc or {}
    with zipfile.ZipFile(caminho) as z:
        itens = z.infolist()
        conteudo = {i.filename: z.read(i.filename) for i in itens}
    wbxml = conteudo["xl/workbook.xml"].decode("utf8")
    rels = conteudo["xl/_rels/workbook.xml.rels"].decode("utf8")
    rel_alvo = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship[^>]*?Id="(rId\d+)"[^>]*?Target="([^"]+)"', rels)}
    rel_alvo.update({m.group(2): m.group(1) for m in re.finditer(r'<Relationship[^>]*?Target="([^"]+)"[^>]*?Id="(rId\d+)"', rels)})
    seq = [1]
    for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="(rId\d+)"', wbxml):
        aba, rid = m.group(1), m.group(2)
        alvo = rel_alvo[rid].lstrip("/")
        alvo = alvo if alvo.startswith("xl/") else "xl/" + alvo
        xml = conteudo[alvo].decode("utf8")
        assert "<extLst>" not in xml
        xml, ext_db = _databars_2010(xml, seq)
        if aba in ignorar_calc:
            tag = f'<ignoredErrors><ignoredError sqref="{ignorar_calc[aba]}" calculatedColumn="1"/></ignoredErrors>'
            pos = min(i for i in (xml.find(t) for t in ("<drawing", "<legacyDrawing", "<tableParts", "<extLst", "</worksheet>")) if i >= 0)
            xml = xml[:pos] + tag + xml[pos:]
        ext_sp = _xml_sparklines(aba, *sparklines[aba]) if aba in sparklines else ""
        if ext_db or ext_sp:
            xml = xml.replace("</worksheet>", f"<extLst>{ext_db}{ext_sp}</extLst></worksheet>")
        conteudo[alvo] = xml.encode("utf8")
    tmp = caminho + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for i in itens:
            z.writestr(i, conteudo[i.filename])
    os.replace(tmp, caminho)


if __name__ == "__main__":
    pasta = os.path.dirname(os.path.abspath(__file__))
    if ARGS.dados:
        dados = importar(ARGS.dados)
        info = dados["info_importacao"]
        print(f"Importados {info['linhas']} lançamentos ({info['sugeridas']} com categoria sugerida pelo histórico, "
              f"{info['sem_categoria']} 'Sem categoria').")
    else:
        dados = dados_exemplo()
        preparar_revisao(dados, {})
        preparar_acertos(dados)
    destino = ARGS.saida or os.path.join(pasta, NOME_ARQUIVO if ARGS.dados else "Financas_Pessoais_Dashboard_EXEMPLO.xlsx")
    construir(dados, destino)
    print(f"Planilha gerada: {destino} (tema {TEMA})")
