# -*- coding: utf-8 -*-
"""
Gera a planilha "Financas_Pessoais_Dashboard.xlsx" do zero.

    python build_planilha.py            # usa o TEMA definido abaixo
    python build_planilha.py claro      # força o tema claro
    python build_planilha.py escuro     # força o tema escuro

Todos os números do Dashboard, Orçamento e Anual são fórmulas vivas do Excel:
o script só escreve os lançamentos de exemplo, as metas e as listas de configuração.
Fórmulas em inglês com vírgula como separador (o Excel pt-BR traduz ao abrir).
"""
import datetime as dt
import os
import re
import shutil
import sys
import tempfile
import zipfile

from openpyxl import Workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.drawing.text import (CharacterProperties, Font as DFont, Paragraph,
                                   ParagraphProperties, RegularTextRun, RichTextProperties)
from openpyxl.formatting.rule import (ColorScaleRule, DataBarRule, FormulaRule)
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
TEMA = "escuro"  # "escuro" ou "claro"
if len(sys.argv) > 1 and sys.argv[1] in ("escuro", "claro"):
    TEMA = sys.argv[1]

NOME_ARQUIVO = "Financas_Pessoais_Dashboard.xlsx"
FONTE = "Calibri"
HOJE = dt.date(2026, 9, 28)          # referência para os dados de exemplo
MES_PADRAO, ANO_PADRAO = 9, 2026      # mês/ano que o Dashboard abre selecionado

PALETAS = {
    "escuro": dict(
        fundo="121417", card="1E2228", card2="262B33", texto="E8EAED", suave="9AA0A6",
        borda="2C323B", grade="2A2F37", destaque="4F8CFF", positivo="2ECC71",
        alerta="F5B942", negativo="FF5C5C", cabecalho_txt="FFFFFF", roxo="9B7BFF",
        heat_max="6B2F36",
    ),
    "claro": dict(
        fundo="F7F8FA", card="FFFFFF", card2="EEF1F5", texto="1F2937", suave="6B7280",
        borda="E5E7EB", grade="E5E7EB", destaque="2563EB", positivo="16A34A",
        alerta="D97706", negativo="DC2626", cabecalho_txt="FFFFFF", roxo="7C3AED",
        heat_max="F8C4C4",
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
CORES_CATEGORIAS = [P["destaque"], P["positivo"], P["alerta"], P["negativo"], P["roxo"],
                    "2EC4B6", "FF9F43", "8A94A6"]

FMT_MOEDA = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
FMT_MOEDA_ZERO_TRACO = '"R$" #,##0.00;[Red]-"R$" #,##0.00;"–"'
FMT_DATA = "dd/mm/yyyy"
FMT_PCT = "0.0%"
FMT_VAR = '"▲ "0.0%;"▼ "0.0%;0.0%'

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto",
         "Setembro", "Outubro", "Novembro", "Dezembro"]
MESES_ABREV = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

# =============================================================================
# LISTAS DA ABA CONFIG
# =============================================================================
CAT_RECEITA = ["Salário", "Freelance / Extras", "Rendimentos", "Reembolsos", "Outras receitas"]
CAT_DESPESA = [
    ("Moradia", "Aluguel/Financiamento, Condomínio, IPTU, Manutenção"),
    ("Contas da casa", "Energia, Água, Gás, Internet, Celular"),
    ("Alimentação", "Mercado, Feira, Padaria, Açougue"),
    ("Restaurantes e delivery", "Restaurante, Delivery, Lanches"),
    ("Transporte", "Combustível, Estacionamento, App de transporte, Manutenção do carro"),
    ("Educação", "Escola, Material escolar, Cursos, Livros"),
    ("Saúde", "Plano de saúde, Farmácia, Consultas, Exames"),
    ("Lazer", "Passeios, Cinema, Viagens, Hobbies"),
    ("Assinaturas", "Streaming, Academia, Aplicativos, Música"),
    ("Compras", "Roupas, Casa, Eletrônicos, Presentes"),
    ("Impostos e taxas", "IPVA, Tarifas bancárias, Anuidades"),
    ("Outros", "Diversos"),
]
CONTAS = [("Conta corrente Itaú", "Conta corrente", "Itaú"),
          ("Conta Nubank", "Conta digital", "Nubank"),
          ("Poupança Caixa", "Poupança", "Caixa"),
          ("Carteira", "Dinheiro", "—")]
CARTOES = [("Cartão Nubank", "Mastercard", 3, 10, 8000),
           ("Cartão Itaú", "Visa", 25, 5, 12000)]
FORMAS = ["PIX", "Cartão de crédito", "Cartão de débito", "Boleto", "Débito automático",
          "Transferência", "Dinheiro"]
TIPOS = ["Receita", "Despesa"]
ANOS = list(range(2024, 2033))

ORCAMENTO = {"Moradia": 3100, "Contas da casa": 520, "Alimentação": 1700,
             "Restaurantes e delivery": 450, "Transporte": 700, "Educação": 1500,
             "Saúde": 1100, "Lazer": 600, "Assinaturas": 260, "Compras": 500,
             "Impostos e taxas": 200, "Outros": 300}

METAS = [("Reserva de emergência", 36000, 14200, dt.date(2027, 12, 31)),
         ("Viagem de férias (família)", 12000, 4300, dt.date(2027, 1, 31)),
         ("Troca do carro", 45000, 9800, dt.date(2028, 12, 31)),
         ("Notebook novo", 6500, 5200, dt.date(2026, 12, 15))]
DIVIDAS = [("Financiamento do carro", 1180.00, 22, 48, dt.date(2028, 11, 10)),
           ("Geladeira (cartão Itaú)", 349.90, 7, 10, dt.date(2026, 12, 5)),
           ("Curso de inglês", 420.00, 3, 12, dt.date(2027, 6, 10))]


# =============================================================================
# DADOS DE EXEMPLO (Abr–Set/2026) — marcados como "EXEMPLO" na Observação
# (data, descrição, tipo, categoria, subcategoria, conta, forma, valor)
# =============================================================================
def dados_exemplo():
    R, D = "Receita", "Despesa"
    itau, nu, cx, cart = "Conta corrente Itaú", "Conta Nubank", "Poupança Caixa", "Carteira"
    cnu, citau = "Cartão Nubank", "Cartão Itaú"
    linhas = []
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
        linhas += [
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
            linhas.append((d(day), *rest))
    linhas.sort(key=lambda x: x[0])
    return linhas


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
HEAD = fill(P["destaque"])
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
def construir(caminho):
    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_lanc = wb.create_sheet("Lancamentos")
    ws_orc = wb.create_sheet("Orcamento")
    ws_anual = wb.create_sheet("Anual")
    ws_metas = wb.create_sheet("Metas")
    ws_cfg = wb.create_sheet("Config")
    ws_calc = wb.create_sheet("Calc")

    # ---------------------------------------------------------------- Config
    ws = ws_cfg
    pintar_fundo(ws, 30, 3)
    titulo_aba(ws, "Configurações", "Listas editáveis. Para adicionar um item, digite na primeira linha vazia logo abaixo da tabela: ela cresce sozinha e os menus suspensos acompanham.", 26)
    larguras(ws, {"A": 2, "B": 24, "C": 3, "D": 26, "E": 58, "F": 3, "G": 24, "H": 16, "I": 12,
                  "J": 3, "K": 18, "L": 13, "M": 13, "N": 13, "O": 14, "P": 3, "Q": 22,
                  "R": 3, "S": 12, "T": 3, "U": 7, "V": 13, "W": 9, "X": 3, "Y": 9})
    L0 = 5  # linha de cabeçalho das tabelas de config

    def tabela_cfg(nome, col_ini, cabec, linhas, secao):
        ci = col_ini
        put(ws, f"{get_column_letter(ci)}{L0 - 1}", secao, font=fnt(10, True, P["suave"]))
        for j, h in enumerate(cabec):
            ws.cell(L0, ci + j, h)
        for i, lin in enumerate(linhas):
            for j, v in enumerate(lin if isinstance(lin, (list, tuple)) else [lin]):
                ws.cell(L0 + 1 + i, ci + j, v)
        cf = ci + len(cabec) - 1
        estilo_cabecalho(ws, L0, ci, cf)
        estilo_corpo(ws, L0 + 1, L0 + len(linhas), ci, cf)
        ref = f"{get_column_letter(ci)}{L0}:{get_column_letter(cf)}{L0 + len(linhas)}"
        nova_tabela(ws, nome, ref, cabec)

    tabela_cfg("tbCatReceita", 2, ["Categoria de receita"], CAT_RECEITA, "RECEITAS")
    tabela_cfg("tbCatDespesa", 4, ["Categoria de despesa", "Subcategorias (sugestões)"], CAT_DESPESA, "DESPESAS")
    tabela_cfg("tbContas", 7, ["Conta", "Tipo de conta", "Banco"], CONTAS, "CONTAS / BANCOS")
    tabela_cfg("tbCartoes", 11, ["Cartão", "Bandeira", "Dia fechamento", "Dia vencimento", "Limite"], CARTOES, "CARTÕES")
    tabela_cfg("tbFormas", 17, ["Forma de pagamento"], FORMAS, "FORMAS DE PAGAMENTO")
    tabela_cfg("tbTipos", 19, ["Tipo"], TIPOS, "TIPOS")
    tabela_cfg("tbMeses", 21, ["Nº", "Mês", "Abrev."], [(i + 1, MESES[i], MESES_ABREV[i]) for i in range(12)], "MESES")
    tabela_cfg("tbAnos", 25, ["Ano"], ANOS, "ANOS")
    for r in range(L0 + 1, L0 + 1 + len(CARTOES)):
        ws[f"O{r}"].number_format = FMT_MOEDA
    ws.freeze_panes = "A6"

    def nome_dinamico(nome, col, max_linha=205):
        rng = f"Config!${col}${L0 + 1}:${col}${max_linha}"
        wb.defined_names[nome] = DefinedName(
            nome, attr_text=f"OFFSET(Config!${col}${L0 + 1},0,0,MAX(1,SUMPRODUCT(--(LEN({rng})>0))),1)")

    nome_dinamico("CatReceita", "B")
    nome_dinamico("CatDespesa", "D")
    nome_dinamico("ListaContas", "G")
    nome_dinamico("ListaCartoes", "K")
    nome_dinamico("FormasPagamento", "Q")
    nome_dinamico("Tipos", "S")
    nome_dinamico("Anos", "Y")
    wb.defined_names["Meses"] = DefinedName("Meses", attr_text=f"Config!$V${L0 + 1}:$V${L0 + 12}")
    wb.defined_names["MesesAbrev"] = DefinedName("MesesAbrev", attr_text=f"Config!$W${L0 + 1}:$W${L0 + 12}")

    # ---------------------------------------------------------------- Lançamentos
    ws = ws_lanc
    dados = dados_exemplo()
    n = len(dados)
    H = 4                      # linha do cabeçalho
    ult = H + n                # última linha de dados
    pintar_fundo(ws, 14, 3)
    titulo_aba(ws, "Lançamentos", "Lance cada receita ou despesa numa linha (valor sempre positivo — o Tipo define se entra ou sai). Linhas com \"EXEMPLO\" na Observação são dados fictícios: filtre e apague quando quiser.", 12)
    larguras(ws, {"A": 2, "B": 13, "C": 34, "D": 11, "E": 24, "F": 22, "G": 22, "H": 20,
                  "I": 15, "J": 7, "K": 8, "L": 16, "M": 2})
    cab_lanc = ["Data", "Descrição", "Tipo", "Categoria", "Subcategoria", "Conta",
                "Forma de pagamento", "Valor", "Mês", "Ano", "Observação"]
    f_mes = '=IF(tbLancamentos[[#This Row],[Data]]="","",MONTH(tbLancamentos[[#This Row],[Data]]))'
    f_ano = '=IF(tbLancamentos[[#This Row],[Data]]="","",YEAR(tbLancamentos[[#This Row],[Data]]))'
    for j, h in enumerate(cab_lanc):
        ws.cell(H, 2 + j, h)
    estilo_cabecalho(ws, H, 2, 12)
    estilo_corpo(ws, H + 1, ult, 2, 12)
    for i, lin in enumerate(dados):
        r = H + 1 + i
        for j, v in enumerate(lin):
            ws.cell(r, 2 + j, v)
        ws.cell(r, 10, f_mes)
        ws.cell(r, 11, f_ano)
        ws.cell(r, 12, "EXEMPLO").font = fnt(9, color=P["suave"], italic=True)
        ws.cell(r, 2).number_format = FMT_DATA
        ws.cell(r, 9).number_format = FMT_MOEDA
        ws.cell(r, 9).alignment = ALIGN_R
        for c in (10, 11):
            ws.cell(r, c).alignment = ALIGN_C
            ws.cell(r, c).font = fnt(10, color=P["suave"])
    nova_tabela(ws, "tbLancamentos", f"B{H}:L{ult}", cab_lanc,
                formulas={"Mês": f_mes, "Ano": f_ano})
    # Formatos para as próximas linhas digitadas (a tabela cresce por baixo)
    for r in range(ult + 1, 1001):
        ws.cell(r, 2).number_format = FMT_DATA
        ws.cell(r, 9).number_format = FMT_MOEDA
    ws.freeze_panes = f"C{H + 1}"
    LIM = 5000
    dv_lista(ws, "=Tipos", f"D{H + 1}:D{LIM}", msg="Use Receita ou Despesa.")
    dv_lista(ws, "=Categorias", f"E{H + 1}:E{LIM}", msg="Escolha uma categoria cadastrada na aba Config.")
    dv_lista(ws, "=Contas", f"G{H + 1}:G{LIM}", msg="Escolha uma conta ou cartão cadastrado na aba Config.")
    dv_lista(ws, "=FormasPagamento", f"H{H + 1}:H{LIM}", msg="Escolha uma forma de pagamento cadastrada na aba Config.")
    dv_val = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv_val.error, dv_val.errorTitle, dv_val.showErrorMessage = "Digite um valor positivo: o Tipo define se é entrada ou saída.", "Valor inválido", True
    ws.add_data_validation(dv_val)
    dv_val.add(f"I{H + 1}:I{LIM}")
    dv_dt = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
    dv_dt.error, dv_dt.errorTitle, dv_dt.showErrorMessage = "Digite uma data válida (dd/mm/aaaa).", "Data inválida", True
    ws.add_data_validation(dv_dt)
    dv_dt.add(f"B{H + 1}:B{LIM}")
    ws.conditional_formatting.add(f"I{H + 1}:I{LIM}", FormulaRule(formula=[f'$D{H + 1}="Receita"'], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(f"D{H + 1}:D{LIM}", FormulaRule(formula=[f'$D{H + 1}="Receita"'], font=Font(color="FF" + P["positivo"])))
    ws.conditional_formatting.add(f"D{H + 1}:D{LIM}", FormulaRule(formula=[f'$D{H + 1}="Despesa"'], font=Font(color="FF" + P["negativo"])))

    # ---------------------------------------------------------------- Calc (oculta)
    ws = ws_calc
    ws.sheet_state = "hidden"
    larguras(ws, {"A": 26, "B": 14})
    c = ws
    L = "tbLancamentos"

    def sif(tipo, mes, ano, extra=""):
        return f'SUMIFS({L}[Valor],{L}[Tipo],"{tipo}",{L}[Mês],{mes},{L}[Ano],{ano}{extra})'

    param = [
        ("Mês selecionado (nº)", "=MATCH(MesSelecionado,Meses,0)"),                   # B2
        ("Ano selecionado", "=AnoSelecionado"),                                        # B3
        ("Data de referência", "=DATE(B3,B2,1)"),                                      # B4
        ("Mês anterior (data)", "=DATE(B3,B2-1,1)"),                                   # B5
        ("Mês anterior (nº)", "=MONTH(B5)"),                                           # B6
        ("Ano do mês anterior", "=YEAR(B5)"),                                          # B7
        ("", None),                                                                    # B8
        ("Receitas do mês", "=" + sif("Receita", "$B$2", "$B$3")),                     # B9
        ("Despesas do mês", "=" + sif("Despesa", "$B$2", "$B$3")),                     # B10
        ("Receitas mês anterior", "=" + sif("Receita", "$B$6", "$B$7")),               # B11
        ("Despesas mês anterior", "=" + sif("Despesa", "$B$6", "$B$7")),               # B12
        ("Saldo do mês", "=B9-B10"),                                                   # B13
        ("Saldo mês anterior", "=B11-B12"),                                            # B14
        ("Taxa de poupança", "=IF(B9>0,B13/B9,0)"),                                    # B15
        ("Variação despesas vs. anterior", "=IF(B12>0,B10/B12-1,0)"),                  # B16
        ("Orçamento total (metas)", "=SUM(tbOrcamento[Meta mensal])"),                 # B17
        ("% do orçamento utilizado", "=IF(B17>0,B10/B17,0)"),                          # B18
        ("Taxa de poupança mês anterior", "=IF(B11>0,B14/B11,0)"),                     # B19
        ("Qtde categorias receita", f"=SUMPRODUCT(--(LEN(Config!$B$6:$B$205)>0))"),   # B20
        ("Qtde categorias despesa", f"=SUMPRODUCT(--(LEN(Config!$D$6:$D$205)>0))"),   # B21
        ("Qtde contas", f"=SUMPRODUCT(--(LEN(Config!$G$6:$G$205)>0))"),               # B22
        ("Qtde cartões", f"=SUMPRODUCT(--(LEN(Config!$K$6:$K$205)>0))"),              # B23
        ("Qtde categorias no orçamento", "=COUNTA(tbOrcamento[Categoria])"),           # B24
        ("Diferença despesas vs. anterior", "=B10-B12"),                               # B25
        ("Meses com dados no ano", "=SUMPRODUCT(--(L2:L13+M2:M13>0))"),                # B26
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
    for j, h in enumerate(["Data", "Rótulo", "Receitas", "Despesas", "Saldo"]):
        c.cell(1, 4 + j, h)
    for i in range(12):
        r = 2 + i
        c.cell(r, 4, f"=DATE($B$3,$B$2-{11 - i},1)").number_format = FMT_DATA
        c.cell(r, 5, f'=INDEX(MesesAbrev,MONTH(D{r}))&"/"&RIGHT(YEAR(D{r}),2)')
        c.cell(r, 6, "=" + sif("Receita", f"MONTH(D{r})", f"YEAR(D{r})"))
        c.cell(r, 7, "=" + sif("Despesa", f"MONTH(D{r})", f"YEAR(D{r})"))
        c.cell(r, 8, f"=F{r}-G{r}")

    # Ano selecionado mês a mês (J:O)
    for j, h in enumerate(["Mês", "Rótulo", "Receitas", "Despesas", "Saldo", "Saldo acumulado"]):
        c.cell(1, 10 + j, h)
    for i in range(12):
        r = 2 + i
        c.cell(r, 10, i + 1)
        c.cell(r, 11, f"=INDEX(MesesAbrev,J{r})")
        c.cell(r, 12, "=" + sif("Receita", f"J{r}", "$B$3"))
        c.cell(r, 13, "=" + sif("Despesa", f"J{r}", "$B$3"))
        c.cell(r, 14, f"=L{r}-M{r}")
        c.cell(r, 15, f"=SUM($N$2:N{r})")

    # Despesas por categoria no mês (Q:S) + ranking (U:X) p/ a rosca
    NSLOT = 40
    for j, h in enumerate(["Categoria", "Realizado", "Chave ordenação"]):
        c.cell(1, 17 + j, h)
    for i in range(NSLOT):
        r = 2 + i
        c.cell(r, 17, f'=IF({i + 1}<=$B$21,INDEX(Config!$D$6:$D$205,{i + 1}),"")')
        c.cell(r, 18, f'=IF(Q{r}="",0,SUMIFS({L}[Valor],{L}[Categoria],Q{r},{L}[Tipo],"Despesa",{L}[Mês],$B$2,{L}[Ano],$B$3))')
        c.cell(r, 19, f'=IF(Q{r}="",-1,ROUND(R{r},2)+({NSLOT + 1 - (i + 1)})/10000000)')
    for j, h in enumerate(["Posição", "Chave", "Linha", "Categoria (rosca)", "Valor (rosca)"]):
        c.cell(1, 20 + j, h)
    for k in range(1, 8):
        r = 1 + k
        c.cell(r, 20, k)
        c.cell(r, 21, f"=LARGE($S$2:$S${NSLOT + 1},T{r})")
        c.cell(r, 22, f"=MATCH(U{r},$S$2:$S${NSLOT + 1},0)")
        c.cell(r, 23, f'=IF(U{r}<0,"",INDEX($Q$2:$Q${NSLOT + 1},V{r}))')
        c.cell(r, 24, f"=IF(U{r}<0,0,INDEX($R$2:$R${NSLOT + 1},V{r}))")
    c.cell(9, 20, 8)
    c.cell(9, 23, "Outras")
    c.cell(9, 24, "=MAX(0,$B$10-SUM(X2:X8))")

    # Orçado x realizado (Z:AC) — espelha tbOrcamento
    N_ORC = len(ORCAMENTO)
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
    cond = f'({L}[Tipo]="Despesa")*({L}[Mês]=$B$2)*({L}[Ano]=$B$3)'
    for k in range(1, 6):
        r = 1 + k
        c.cell(r, 31, k)
        ref = f"AF{r}"
        c[ref] = ArrayFormula(ref, f"=IFERROR(LARGE(IF({cond},ROUND({L}[Valor],2)+ROW({L}[Valor])/1000000000),AE{r}),0)")
        c.cell(r, 34, f'=IF(AG{r}=0,"",INDEX({L}[Data],AG{r}))')
        c.cell(r, 35, f'=IF(AG{r}=0,"",INDEX({L}[Descrição],AG{r}))')
        c.cell(r, 36, f'=IF(AG{r}=0,"",INDEX({L}[Categoria],AG{r}))')
        c.cell(r, 37, f'=IF(AG{r}=0,"",INDEX({L}[Valor],AG{r}))')
    # AG usa MATCH(linha, ROW(coluna)) → precisa ser matricial também
    for k in range(1, 6):
        r = 1 + k
        ref = f"AG{r}"
        c[ref] = ArrayFormula(ref, f"=IF(AF{r}<=0,0,MATCH(ROUND((AF{r}-ROUND(AF{r},2))*1000000000,0),ROW({L}[Valor]),0))")

    # Listas combinadas para os menus suspensos (AM: categorias, AO: contas+cartões)
    c.cell(1, 39, "Categorias (receitas + despesas)")
    for i in range(80):
        r = 2 + i
        k = i + 1
        c.cell(r, 39, f'=IF({k}<=$B$20,INDEX(Config!$B$6:$B$205,{k}),IF({k}-$B$20<=$B$21,INDEX(Config!$D$6:$D$205,MAX(1,{k}-$B$20)),""))')
    c.cell(1, 41, "Contas + cartões")
    for i in range(40):
        r = 2 + i
        k = i + 1
        c.cell(r, 41, f'=IF({k}<=$B$22,INDEX(Config!$G$6:$G$205,{k}),IF({k}-$B$22<=$B$23,INDEX(Config!$K$6:$K$205,MAX(1,{k}-$B$22)),""))')
    wb.defined_names["Categorias"] = DefinedName("Categorias", attr_text="OFFSET(Calc!$AM$2,0,0,MAX(1,Calc!$B$20+Calc!$B$21),1)")
    wb.defined_names["Contas"] = DefinedName("Contas", attr_text="OFFSET(Calc!$AO$2,0,0,MAX(1,Calc!$B$22+Calc!$B$23),1)")

    # ---------------------------------------------------------------- Orçamento
    ws = ws_orc
    pintar_fundo(ws, 10, 3)
    titulo_aba(ws, "Orçamento", "Defina a meta mensal de cada categoria de despesa. O realizado acompanha o mês/ano escolhido no Dashboard.", 7)
    larguras(ws, {"A": 2, "B": 28, "C": 16, "D": 20, "E": 16, "F": 14, "G": 3, "H": 40})
    put(ws, "H2", '="Mês de referência: "&MesSelecionado&" / "&AnoSelecionado', font=fnt(11, True, P["destaque"]), align=Alignment(vertical="center"))
    HO = 5
    cab_orc = ["Categoria", "Meta mensal", "Realizado no mês selecionado", "Diferença", "% usado"]
    for j, h in enumerate(cab_orc):
        ws.cell(HO, 2 + j, h)
    estilo_cabecalho(ws, HO, 2, 6)
    ws.row_dimensions[HO].height = 32
    f_real = (f'=SUMIFS({L}[Valor],{L}[Categoria],tbOrcamento[[#This Row],[Categoria]],{L}[Tipo],"Despesa",'
              f'{L}[Mês],MesNum,{L}[Ano],AnoSelecionado)')
    f_dif = "=tbOrcamento[[#This Row],[Meta mensal]]-tbOrcamento[[#This Row],[Realizado no mês selecionado]]"
    f_pct = ("=IF(tbOrcamento[[#This Row],[Meta mensal]]>0,"
             "tbOrcamento[[#This Row],[Realizado no mês selecionado]]/tbOrcamento[[#This Row],[Meta mensal]],0)")
    ult_o = HO + N_ORC
    estilo_corpo(ws, HO + 1, ult_o + 1, 2, 6)
    for i, (cat, meta) in enumerate(ORCAMENTO.items()):
        r = HO + 1 + i
        ws.cell(r, 2, cat)
        ws.cell(r, 3, meta)
        ws.cell(r, 4, f_real)
        ws.cell(r, 5, f_dif)
        ws.cell(r, 6, f_pct)
    tot = ult_o + 1
    ws.cell(tot, 2, "Total")
    ws.cell(tot, 3, "=SUBTOTAL(109,tbOrcamento[Meta mensal])")
    ws.cell(tot, 4, "=SUBTOTAL(109,tbOrcamento[Realizado no mês selecionado])")
    ws.cell(tot, 5, "=SUBTOTAL(109,tbOrcamento[Diferença])")
    ws.cell(tot, 6, f"=IF(C{tot}>0,D{tot}/C{tot},0)")
    area(ws, f"B{tot}:F{tot}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for r in range(HO + 1, tot + 1):
        for col in "CDE":
            ws[f"{col}{r}"].number_format = FMT_MOEDA
            ws[f"{col}{r}"].alignment = ALIGN_R
        ws[f"F{r}"].number_format = FMT_PCT
        ws[f"F{r}"].alignment = ALIGN_R
    nova_tabela(ws, "tbOrcamento", f"B{HO}:F{tot}", cab_orc,
                formulas={"Realizado no mês selecionado": f_real, "Diferença": f_dif, "% usado": f_pct},
                totais={"Categoria": "Total", "Meta mensal": "sum",
                        "Realizado no mês selecionado": "sum", "Diferença": "sum"})
    ws.tables["tbOrcamento"].tableColumns[4].totalsRowFunction = "custom"
    ws.tables["tbOrcamento"].tableColumns[4].totalsRowFormula = TableFormula(
        attr_text="IF(SUM(tbOrcamento[Meta mensal])>0,SUM(tbOrcamento[Realizado no mês selecionado])/SUM(tbOrcamento[Meta mensal]),0)")
    ws.cell(tot, 6, "=IF(SUM(tbOrcamento[Meta mensal])>0,SUM(tbOrcamento[Realizado no mês selecionado])/SUM(tbOrcamento[Meta mensal]),0)")
    regras_percentual(ws, f"F{HO + 1}:F{tot}", f"F{HO + 1}")
    ws.conditional_formatting.add(f"E{HO + 1}:E{tot}", FormulaRule(formula=[f"E{HO + 1}<0"], font=Font(color="FF" + P["negativo"])))
    dv_lista(ws, "=CatDespesa", f"B{HO + 1}:B{tot + 40}", msg="Use uma categoria de despesa cadastrada na aba Config.")
    put(ws, "H5", "Legenda do % usado", font=fnt(10, True))
    for i, (txt, cor) in enumerate([("até 80% — dentro do planejado", P["positivo"]),
                                    ("80% a 100% — atenção", P["alerta"]),
                                    ("acima de 100% — estourou a meta", P["negativo"])]):
        put(ws, f"H{6 + i}", "●  " + txt, font=fnt(10, color=cor))
    ws.freeze_panes = f"A{HO + 1}"

    # ---------------------------------------------------------------- Anual
    ws = ws_anual
    pintar_fundo(ws, 20, 3)
    titulo_aba(ws, "Visão anual", "Categorias × meses do ano selecionado no Dashboard. Linhas vazias estão reservadas para categorias novas cadastradas na aba Config.", 17)
    put(ws, "O2", '="Ano: "&AnoSelecionado', font=fnt(14, True, P["destaque"]), align=Alignment(horizontal="right", vertical="center"))
    ws.merge_cells("O2:Q2")
    larguras(ws, {"A": 2, "B": 26, **{get_column_letter(3 + i): 11.5 for i in range(12)},
                  "O": 14, "P": 13, "Q": 16, "R": 2})
    HA = 5
    cab = ["Categoria"] + MESES_ABREV + ["Total", "Média mensal", "Tendência"]
    for j, h in enumerate(cab):
        ws.cell(HA, 2 + j, h)
    estilo_cabecalho(ws, HA, 2, 17)
    for col in range(3, 17):
        ws.cell(HA, col).alignment = Alignment(horizontal="right", vertical="center", indent=1)
    ws.cell(HA, 17).alignment = ALIGN_C

    def linha_anual(r, cat_formula, tipo, secao_fill=CARD):
        ws.row_dimensions[r].height = 20
        ws.cell(r, 2, cat_formula)
        for m in range(12):
            col = get_column_letter(3 + m)
            ws.cell(r, 3 + m, f'=IF($B{r}="",0,SUMIFS({L}[Valor],{L}[Categoria],$B{r},{L}[Tipo],"{tipo}",{L}[Mês],{m + 1},{L}[Ano],AnoSelecionado))')
        ws.cell(r, 15, f"=SUM(C{r}:N{r})")
        ws.cell(r, 16, f"=IF(Calc!$B$26>0,O{r}/Calc!$B$26,0)")
        area(ws, f"B{r}:Q{r}", fill_=secao_fill, font=fnt(10), border=Border(bottom=side(P["borda"])))
        area(ws, f"C{r}:P{r}", fmt=FMT_MOEDA_ZERO_TRACO, align=ALIGN_R)
        ws.cell(r, 2).alignment = ALIGN_L

    def secao(r, texto, cor):
        ws.row_dimensions[r].height = 22
        area(ws, f"B{r}:Q{r}", fill_=CARD2, font=fnt(10, True, cor))
        ws.cell(r, 2, texto)
        ws.cell(r, 2).alignment = ALIGN_L

    r = HA + 1
    secao(r, "RECEITAS", P["positivo"])
    rec_ini = r + 1
    for i in range(len(CAT_RECEITA) + 2):
        r += 1
        linha_anual(r, f'=IF({i + 1}<=Calc!$B$20,INDEX(Config!$B$6:$B$205,{i + 1}),"")', "Receita")
    rec_fim = r
    r += 1
    tot_rec = r
    ws.cell(r, 2, "Total de receitas")
    for col in range(3, 17):
        cl = get_column_letter(col)
        ws.cell(r, col, f"=SUM({cl}{rec_ini}:{cl}{rec_fim})")
    area(ws, f"B{r}:Q{r}", fill_=CARD2, font=fnt(10, True, P["positivo"]), fmt=FMT_MOEDA_ZERO_TRACO,
         border=Border(top=side(P["positivo"])))
    area(ws, f"C{r}:P{r}", align=ALIGN_R)
    ws.cell(r, 2).alignment = ALIGN_L
    ws.cell(r, 16, f"=IF(Calc!$B$26>0,O{r}/Calc!$B$26,0)")
    r += 2
    secao(r, "DESPESAS", P["negativo"])
    des_ini = r + 1
    for i in range(len(CAT_DESPESA) + 3):
        r += 1
        linha_anual(r, f'=IF({i + 1}<=Calc!$B$21,INDEX(Config!$D$6:$D$205,{i + 1}),"")', "Despesa")
    des_fim = r
    r += 1
    tot_des = r
    ws.cell(r, 2, "Total de despesas")
    for col in range(3, 17):
        cl = get_column_letter(col)
        ws.cell(r, col, f"=SUM({cl}{des_ini}:{cl}{des_fim})")
    area(ws, f"B{r}:Q{r}", fill_=CARD2, font=fnt(10, True, P["negativo"]), fmt=FMT_MOEDA_ZERO_TRACO,
         border=Border(top=side(P["negativo"])))
    area(ws, f"C{r}:P{r}", align=ALIGN_R)
    ws.cell(r, 2).alignment = ALIGN_L
    ws.cell(r, 16, f"=IF(Calc!$B$26>0,O{r}/Calc!$B$26,0)")
    r += 2
    saldo_l = r
    ws.cell(r, 2, "Saldo do mês")
    for col in range(3, 17):
        cl = get_column_letter(col)
        ws.cell(r, col, f"={cl}{tot_rec}-{cl}{tot_des}")
    area(ws, f"B{r}:Q{r}", fill_=CARD2, font=fnt(11, True), fmt=FMT_MOEDA_ZERO_TRACO,
         border=Border(top=side(P["destaque"])))
    area(ws, f"C{r}:P{r}", align=ALIGN_R)
    ws.cell(r, 2).alignment = ALIGN_L
    ws.row_dimensions[r].height = 22
    r += 1
    ws.cell(r, 2, "Taxa de poupança")
    for col in range(3, 17):
        cl = get_column_letter(col)
        ws.cell(r, col, f"=IF({cl}{tot_rec}>0,{cl}{saldo_l}/{cl}{tot_rec},0)")
    area(ws, f"B{r}:Q{r}", fill_=CARD, font=fnt(10, color=P["suave"]), fmt=FMT_PCT)
    area(ws, f"C{r}:P{r}", align=ALIGN_R)
    ws.cell(r, 2).alignment = ALIGN_L
    ws.cell(r, 16).value = None
    ws.freeze_panes = f"C{HA + 1}"
    # Mapa de calor discreto nas despesas
    ws.conditional_formatting.add(f"C{des_ini}:N{des_fim}", ColorScaleRule(
        start_type="num", start_value=0, start_color="FF" + P["card"],
        end_type="max", end_color="FF" + P["heat_max"]))
    ws.conditional_formatting.add(f"C{saldo_l}:O{saldo_l}", FormulaRule(formula=[f"C{saldo_l}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    linhas_spark = list(range(rec_ini, rec_fim + 1)) + [tot_rec] + list(range(des_ini, des_fim + 1)) + [tot_des, saldo_l]

    # ---------------------------------------------------------------- Metas & Dívidas
    ws = ws_metas
    pintar_fundo(ws, 18, 3)
    titulo_aba(ws, "Metas & Dívidas", "Acompanhe objetivos de poupança e parcelamentos. Atualize o \"Valor atual\" e as \"Parcelas pagas\" quando fizer aportes/pagamentos.", 15)
    larguras(ws, {"A": 2, "B": 28, "C": 14, "D": 14, "E": 14, "F": 12, "G": 16, "H": 4,
                  "I": 28, "J": 14, "K": 11, "L": 11, "M": 15, "N": 14, "O": 2})
    HM = 5
    put(ws, "B4", "METAS", font=fnt(10, True, P["suave"]))
    put(ws, "I4", "DÍVIDAS E PARCELAMENTOS", font=fnt(10, True, P["suave"]))
    cab_m = ["Nome", "Valor alvo", "Valor atual", "% atingido", "Prazo", "Aporte mensal necessário"]
    f_pa = ("=IF(tbMetas[[#This Row],[Valor alvo]]>0,MIN(1,tbMetas[[#This Row],[Valor atual]]/tbMetas[[#This Row],[Valor alvo]]),0)")
    f_ap = ("=IF(tbMetas[[#This Row],[Valor atual]]>=tbMetas[[#This Row],[Valor alvo]],0,"
            "(tbMetas[[#This Row],[Valor alvo]]-tbMetas[[#This Row],[Valor atual]])/"
            "MAX(1,(YEAR(tbMetas[[#This Row],[Prazo]])-YEAR(TODAY()))*12+MONTH(tbMetas[[#This Row],[Prazo]])-MONTH(TODAY())))")
    for j, h in enumerate(cab_m):
        ws.cell(HM, 2 + j, h)
    estilo_cabecalho(ws, HM, 2, 7)
    ws.row_dimensions[HM].height = 32
    estilo_corpo(ws, HM + 1, HM + len(METAS) + 1, 2, 7)
    for i, (nome, alvo, atual, prazo) in enumerate(METAS):
        r = HM + 1 + i
        ws.cell(r, 2, nome)
        ws.cell(r, 3, alvo)
        ws.cell(r, 4, atual)
        ws.cell(r, 5, f_pa)
        ws.cell(r, 6, prazo)
        ws.cell(r, 7, f_ap)
    tm = HM + len(METAS) + 1
    ws.cell(tm, 2, "Total")
    ws.cell(tm, 3, "=SUBTOTAL(109,tbMetas[Valor alvo])")
    ws.cell(tm, 4, "=SUBTOTAL(109,tbMetas[Valor atual])")
    ws.cell(tm, 5, "=IF(SUM(tbMetas[Valor alvo])>0,SUM(tbMetas[Valor atual])/SUM(tbMetas[Valor alvo]),0)")
    ws.cell(tm, 7, "=SUBTOTAL(109,tbMetas[Aporte mensal necessário])")
    area(ws, f"B{tm}:G{tm}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for r in range(HM + 1, tm + 1):
        for col in "CDG":
            ws[f"{col}{r}"].number_format = FMT_MOEDA
            ws[f"{col}{r}"].alignment = ALIGN_R
        ws[f"E{r}"].number_format = FMT_PCT
        ws[f"E{r}"].alignment = ALIGN_R
        ws[f"F{r}"].number_format = FMT_DATA
        ws[f"F{r}"].alignment = ALIGN_C
    nova_tabela(ws, "tbMetas", f"B{HM}:G{tm}", cab_m, formulas={"% atingido": f_pa, "Aporte mensal necessário": f_ap},
                totais={"Nome": "Total", "Valor alvo": "sum", "Valor atual": "sum", "Aporte mensal necessário": "sum"})
    tcm = ws.tables["tbMetas"].tableColumns[3]
    tcm.totalsRowFunction = "custom"
    tcm.totalsRowFormula = TableFormula(attr_text="IF(SUM(tbMetas[Valor alvo])>0,SUM(tbMetas[Valor atual])/SUM(tbMetas[Valor alvo]),0)")
    ws.conditional_formatting.add(f"E{HM + 1}:E{tm}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                                   color="FF" + BARRA_META, showValue=True))

    cab_d = ["Descrição", "Valor da parcela", "Parcelas pagas", "Total de parcelas", "Saldo devedor", "Data de término"]
    f_sd = ("=tbDividas[[#This Row],[Valor da parcela]]*MAX(0,tbDividas[[#This Row],[Total de parcelas]]-tbDividas[[#This Row],[Parcelas pagas]])")
    for j, h in enumerate(cab_d):
        ws.cell(HM, 9 + j, h)
    estilo_cabecalho(ws, HM, 9, 14)
    estilo_corpo(ws, HM + 1, HM + len(DIVIDAS) + 1, 9, 14)
    for i, (desc, parc, pagas, total, fim) in enumerate(DIVIDAS):
        r = HM + 1 + i
        ws.cell(r, 9, desc)
        ws.cell(r, 10, parc)
        ws.cell(r, 11, pagas)
        ws.cell(r, 12, total)
        ws.cell(r, 13, f_sd)
        ws.cell(r, 14, fim)
    td = HM + len(DIVIDAS) + 1
    ws.cell(td, 9, "Total")
    ws.cell(td, 10, "=SUBTOTAL(109,tbDividas[Valor da parcela])")
    ws.cell(td, 13, "=SUBTOTAL(109,tbDividas[Saldo devedor])")
    area(ws, f"I{td}:N{td}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for r in range(HM + 1, td + 1):
        for col in "JM":
            ws[f"{col}{r}"].number_format = FMT_MOEDA
            ws[f"{col}{r}"].alignment = ALIGN_R
        for col in "KL":
            ws[f"{col}{r}"].alignment = ALIGN_C
        ws[f"N{r}"].number_format = FMT_DATA
        ws[f"N{r}"].alignment = ALIGN_C
    nova_tabela(ws, "tbDividas", f"I{HM}:N{td}", cab_d, formulas={"Saldo devedor": f_sd},
                totais={"Descrição": "Total", "Valor da parcela": "sum", "Saldo devedor": "sum"})
    put(ws, f"I{td + 2}", "Dica: a parcela total do mês também deve ser lançada na aba Lançamentos quando for paga.",
        font=fnt(9, color=P["suave"], italic=True))
    put(ws, f"B{tm + 2}", "Aporte mensal necessário = (valor alvo − valor atual) ÷ meses restantes até o prazo.",
        font=fnt(9, color=P["suave"], italic=True))

    # ---------------------------------------------------------------- Dashboard
    ws = ws_dash
    # Grade: A = margem; 6 cards de 3 colunas (largura 10) separados por colunas de respiro (2.5)
    card_cols = []  # (col_ini, col_fim) índices base 1
    col = 2
    for k in range(6):
        card_cols.append((col, col + 2))
        col += 4
    ult_col = card_cols[-1][1] + 1   # margem direita
    pintar_fundo(ws, ult_col + 6, 75)
    ws.column_dimensions["A"].width = 3
    for ci, cf in card_cols:
        for cc in range(ci, cf + 1):
            ws.column_dimensions[get_column_letter(cc)].width = 10
        if cf + 1 <= ult_col:
            ws.column_dimensions[get_column_letter(cf + 1)].width = 2.5
    ws.column_dimensions[get_column_letter(ult_col)].width = 3
    ws.sheet_view.zoomScale = 90
    CL = get_column_letter

    # Cabeçalho
    for rr, h in {1: 12, 2: 26, 3: 28, 4: 16}.items():
        ws.row_dimensions[rr].height = h
    put(ws, "B2", "Finanças da Família", font=fnt(20, True), align=Alignment(vertical="bottom"))
    put(ws, "B3", "Visão mensal — escolha o mês e o ano nos seletores à direita; tudo se atualiza sozinho.",
        font=fnt(9, color=P["suave"]), align=Alignment(vertical="center"))
    # Seletores
    c5, c6 = card_cols[4], card_cols[5]
    mes_rng = f"{CL(c5[0])}3:{CL(c5[1])}3"
    ano_rng = f"{CL(c6[0])}3:{CL(c6[1])}3"
    put(ws, f"{CL(c5[0])}2", "MÊS", font=fnt(9, True, P["suave"]), align=Alignment(vertical="bottom", indent=1))
    put(ws, f"{CL(c6[0])}2", "ANO", font=fnt(9, True, P["suave"]), align=Alignment(vertical="bottom", indent=1))
    borda_sel = Border(bottom=side(P["destaque"], "medium"))
    mesclar(ws, mes_rng, MESES[MES_PADRAO - 1], fill_=CARD2, font=fnt(14, True, P["destaque"]),
            align=Alignment(horizontal="center", vertical="center"), border=borda_sel)
    mesclar(ws, ano_rng, ANO_PADRAO, fill_=CARD2, font=fnt(14, True, P["destaque"]),
            align=Alignment(horizontal="center", vertical="center"), border=borda_sel)
    mes_cell = f"{CL(c5[0])}3"
    ano_cell = f"{CL(c6[0])}3"
    wb.defined_names["MesSelecionado"] = DefinedName("MesSelecionado", attr_text=f"Dashboard!${CL(c5[0])}$3")
    wb.defined_names["AnoSelecionado"] = DefinedName("AnoSelecionado", attr_text=f"Dashboard!${CL(c6[0])}$3")
    dv_lista(ws, "=Meses", mes_cell, "Mês inválido", "Escolha um mês da lista.")
    dv_lista(ws, "=Anos", ano_cell, "Ano inválido", "Escolha um ano da lista (cadastre novos anos na aba Config).")
    ws.freeze_panes = "A5"

    # KPI cards (linhas 5–9)
    for rr, h in {5: 8, 6: 20, 7: 38, 8: 18, 9: 10, 10: 18}.items():
        ws.row_dimensions[rr].height = h
    kpis = [
        ("RECEITAS DO MÊS", "=Calc!B9", FMT_MOEDA, "Mês anterior", "=Calc!B11", FMT_MOEDA, P["positivo"]),
        ("DESPESAS DO MÊS", "=Calc!B10", FMT_MOEDA, "Mês anterior", "=Calc!B12", FMT_MOEDA, P["negativo"]),
        ("SALDO DO MÊS", "=Calc!B13", FMT_MOEDA, "Mês anterior", "=Calc!B14", FMT_MOEDA, P["destaque"]),
        ("TAXA DE POUPANÇA", "=Calc!B15", FMT_PCT, "Mês anterior", "=Calc!B19", FMT_PCT, P["roxo"]),
        ("DESPESAS VS. MÊS ANTERIOR", "=Calc!B16", FMT_VAR, "Diferença", "=Calc!B25", FMT_MOEDA, P["alerta"]),
        ("% DO ORÇAMENTO UTILIZADO", "=Calc!B18", FMT_PCT, "Meta total", "=Calc!B17", FMT_MOEDA, P["destaque"]),
    ]
    kpi_cells = []
    for (ci, cf), (rot, f, fmt, sub, fsub, fmtsub, cor) in zip(card_cols, kpis):
        a, b = CL(ci), CL(cf)
        area(ws, f"{a}5:{b}9", fill_=CARD)
        area(ws, f"{a}5:{b}5", border=Border(top=side(cor, "thick")))
        mesclar(ws, f"{a}6:{b}6", rot, fill_=CARD, font=fnt(9, True, P["suave"]), align=Alignment(horizontal="left", vertical="center", indent=1))
        mesclar(ws, f"{a}7:{b}7", f, fill_=CARD, font=fnt(22, True), fmt=fmt, align=Alignment(horizontal="left", vertical="center", indent=1))
        put(ws, f"{a}8", sub, font=fnt(9, color=P["suave"]), fill_=CARD, align=Alignment(horizontal="left", vertical="center", indent=1))
        mesclar(ws, f"{CL(ci + 1)}8:{b}8", fsub, fill_=CARD, font=fnt(9, True, P["suave"]), fmt=fmtsub, align=Alignment(horizontal="right", vertical="center", indent=1))
        kpi_cells.append(f"{a}7")
    # Cores condicionais nos KPIs
    sal, var, orc = kpi_cells[2], kpi_cells[4], kpi_cells[5]
    ws.conditional_formatting.add(sal, FormulaRule(formula=[f"{sal}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(sal, FormulaRule(formula=[f"{sal}>=0"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(kpi_cells[3], FormulaRule(formula=[f"{kpi_cells[3]}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(var, FormulaRule(formula=[f"{var}>0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(var, FormulaRule(formula=[f"{var}<0"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(orc, FormulaRule(formula=[f"{orc}>1"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(orc, FormulaRule(formula=[f"AND({orc}>0.8,{orc}<=1)"], font=Font(color="FF" + P["alerta"], bold=True)))
    ws.conditional_formatting.add(orc, FormulaRule(formula=[f"{orc}<=0.8"], font=Font(color="FF" + P["positivo"], bold=True)))

    # Gráficos — duas linhas de dois gráficos
    esq = (card_cols[0][0] - 1, card_cols[2][1])        # col base 0 início, col base 0 fim (exclusivo)
    dir_ = (card_cols[3][0] - 1, card_cols[5][1])
    G1_INI, G1_FIM = 10, 29          # linhas base 0 (linha 11 a 29)
    G2_INI, G2_FIM = 30, 49
    for rr in range(11, 50):
        ws.row_dimensions[rr].height = 18
    ws.row_dimensions[30].height = 14

    # 1) Rosca — despesas por categoria no mês
    ch = DoughnutChart(holeSize=58)
    ch.add_data(Reference(ws_calc, min_col=24, min_row=2, max_row=9), titles_from_data=False)
    ch.set_categories(Reference(ws_calc, min_col=23, min_row=2, max_row=9))
    s = ch.series[0]
    for i in range(8):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=CORES_CATEGORIAS[i % len(CORES_CATEGORIAS)],
                                                     ln=LineProperties(solidFill=P["card"], w=19050))
        s.dPt.append(pt)
    estilizar_grafico(ch, "Despesas por categoria no mês")
    ch.legend.position = "r"
    ch.firstSliceAng = 0
    ancorar(ws, ch, esq[0], G1_INI, esq[1], G1_FIM)

    # 2) Colunas — receitas x despesas últimos 12 meses
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.gapWidth = 70
    ch.overlap = -10
    ch.add_data(Reference(ws_calc, min_col=6, max_col=7, min_row=1, max_row=13), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=5, min_row=2, max_row=13))
    for s, cor in zip(ch.series, (P["positivo"], P["negativo"])):
        s.graphicalProperties = GraphicalProperties(solidFill=cor, ln=LineProperties(noFill=True))
    estilizar_grafico(ch, "Receitas × Despesas — últimos 12 meses")
    estilizar_eixos(ch)
    ancorar(ws, ch, dir_[0], G1_INI, dir_[1], G1_FIM)

    # 3) Linha — saldo acumulado no ano
    ch = LineChart()
    ch.add_data(Reference(ws_calc, min_col=15, min_row=1, max_row=13), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=11, min_row=2, max_row=13))
    s = ch.series[0]
    s.smooth = True
    s.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=P["destaque"], w=31750))
    s.marker = Marker(symbol="circle", size=6)
    s.marker.graphicalProperties = GraphicalProperties(solidFill=P["destaque"], ln=LineProperties(solidFill=P["card"]))
    estilizar_grafico(ch, "Saldo acumulado no ano")
    estilizar_eixos(ch)
    ch.legend = None
    ancorar(ws, ch, esq[0], G2_INI, esq[1], G2_FIM)

    # 4) Barras horizontais — orçado x realizado
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "clustered"
    ch.gapWidth = 50
    ch.overlap = -5
    ch.add_data(Reference(ws_calc, min_col=27, max_col=28, min_row=1, max_row=1 + N_ORC), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=26, min_row=2, max_row=1 + N_ORC))
    for s, cor in zip(ch.series, (P["suave"], P["destaque"])):
        s.graphicalProperties = GraphicalProperties(solidFill=cor, ln=LineProperties(noFill=True))
    estilizar_grafico(ch, "Orçado × Realizado por categoria")
    estilizar_eixos(ch)
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.crosses = "max"
    ch.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill=P["grade"], w=6350)))
    ancorar(ws, ch, dir_[0], G2_INI, dir_[1], G2_FIM)

    # Tabelas do Dashboard
    T0 = 51
    ws.row_dimensions[50].height = 14
    ws.row_dimensions[T0].height = 26
    ws.row_dimensions[T0 + 1].height = 24
    # Top 5 (lado esquerdo: B..L)
    put(ws, f"B{T0}", "Top 5 maiores gastos do mês", font=fnt(12, True), align=Alignment(vertical="center"))
    cols_top = [("Data", "B", "C"), ("Descrição", "D", "H"), ("Categoria", "I", "J"), ("Valor", "K", "L")]
    for rot, a, b in cols_top:
        mesclar(ws, f"{a}{T0 + 1}:{b}{T0 + 1}", rot, fill_=HEAD, font=fnt(10, True, P["cabecalho_txt"]),
                align=ALIGN_R if rot == "Valor" else ALIGN_L)
    for k in range(5):
        r = T0 + 2 + k
        ws.row_dimensions[r].height = 22
        cr = 2 + k
        b = Border(bottom=side(P["borda"]))
        mesclar(ws, f"B{r}:C{r}", f"=Calc!AH{cr}", fill_=CARD, font=fnt(10, color=P["suave"]), fmt=FMT_DATA, align=ALIGN_L, border=b)
        mesclar(ws, f"D{r}:H{r}", f"=Calc!AI{cr}", fill_=CARD, font=fnt(10), align=ALIGN_L, border=b)
        mesclar(ws, f"I{r}:J{r}", f"=Calc!AJ{cr}", fill_=CARD, font=fnt(10, color=P["suave"]), align=ALIGN_L, border=b)
        mesclar(ws, f"K{r}:L{r}", f"=Calc!AK{cr}", fill_=CARD, font=fnt(10, True, P["negativo"]), fmt=FMT_MOEDA, align=ALIGN_R, border=b)

    # Orçado x realizado (lado direito: N..X)
    put(ws, f"N{T0}", "Orçado × Realizado por categoria", font=fnt(12, True), align=Alignment(vertical="center"))
    cols_orc = [("Categoria", "N", "P"), ("Meta", "Q", "R"), ("Realizado", "S", "T"), ("Diferença", "U", "V"), ("% usado", "W", "X")]
    for rot, a, b in cols_orc:
        mesclar(ws, f"{a}{T0 + 1}:{b}{T0 + 1}", rot, fill_=HEAD, font=fnt(10, True, P["cabecalho_txt"]),
                align=ALIGN_L if rot == "Categoria" else ALIGN_R)
    for i in range(N_ORC):
        r = T0 + 2 + i
        cr = 2 + i
        ws.row_dimensions[r].height = 22
        b = Border(bottom=side(P["borda"]))
        mesclar(ws, f"N{r}:P{r}", f"=Calc!Z{cr}", fill_=CARD, font=fnt(10), align=ALIGN_L, border=b)
        mesclar(ws, f"Q{r}:R{r}", f"=Calc!AA{cr}", fill_=CARD, font=fnt(10, color=P["suave"]), fmt=FMT_MOEDA, align=ALIGN_R, border=b)
        mesclar(ws, f"S{r}:T{r}", f"=Calc!AB{cr}", fill_=CARD, font=fnt(10), fmt=FMT_MOEDA, align=ALIGN_R, border=b)
        mesclar(ws, f"U{r}:V{r}", f"=Q{r}-S{r}", fill_=CARD, font=fnt(10), fmt=FMT_MOEDA, align=ALIGN_R, border=b)
        mesclar(ws, f"W{r}:X{r}", f"=Calc!AC{cr}", fill_=CARD, font=fnt(10, True), fmt=FMT_PCT, align=ALIGN_R, border=b)
    fim_orc = T0 + 1 + N_ORC
    regras_percentual(ws, f"W{T0 + 2}:X{fim_orc}", f"W{T0 + 2}")
    ws.conditional_formatting.add(f"U{T0 + 2}:V{fim_orc}", FormulaRule(formula=[f"U{T0 + 2}<0"], font=Font(color="FF" + P["negativo"])))
    # Legenda / dica abaixo do Top 5
    put(ws, f"B{T0 + 8}", "Como usar", font=fnt(10, True, P["suave"]))
    dicas = ["1. Lance receitas e despesas na aba Lançamentos (valor sempre positivo).",
             "2. Troque o mês/ano no topo: tudo se atualiza sozinho.",
             "3. Ajuste metas na aba Orçamento e listas na aba Config.",
             "% usado: verde até 80% • amarelo 80–100% • vermelho acima de 100%."]
    for i, d in enumerate(dicas):
        put(ws, f"B{T0 + 9 + i}", d, font=fnt(9, color=P["suave"]))

    # Impressão / PDF do Dashboard
    ws.print_area = f"A1:{CL(ult_col)}{fim_orc + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.3

    for w in (ws_lanc, ws_orc, ws_anual, ws_metas, ws_cfg):
        w.page_setup.orientation = "landscape"
        w.page_setup.paperSize = w.PAPERSIZE_A4
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.page_margins.left = w.page_margins.right = 0.3

    # Cores das guias
    ws_dash.sheet_properties.tabColor = P["destaque"]
    ws_lanc.sheet_properties.tabColor = P["positivo"]
    ws_orc.sheet_properties.tabColor = P["alerta"]
    ws_anual.sheet_properties.tabColor = P["roxo"]
    ws_metas.sheet_properties.tabColor = "2EC4B6"
    ws_cfg.sheet_properties.tabColor = P["suave"]
    ws_calc.sheet_properties.tabColor = P["borda"]

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

    wb.active = 0
    for w in wb.worksheets:
        w.sheet_view.tabSelected = w is ws_dash
    wb.calculation.fullCalcOnLoad = True
    wb.save(caminho)
    pos_processar(caminho, {"Anual": ("C", "N", "Q", linhas_spark)})


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


def pos_processar(caminho, sparklines):
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
    destino = os.path.join(os.path.dirname(os.path.abspath(__file__)), NOME_ARQUIVO)
    construir(destino)
    print(f"Planilha gerada: {destino} (tema {TEMA})")
