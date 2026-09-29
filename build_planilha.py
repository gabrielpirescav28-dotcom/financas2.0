# -*- coding: utf-8 -*-
"""
Gera a planilha de finanças pessoais com dashboard do zero.

    python build_planilha.py                                   # dados de exemplo
    python build_planilha.py --dados Financas_Plinio.xlsx      # importa os dados reais
    python build_planilha.py --tema claro                      # tema claro
    python build_planilha.py --dados X.xlsx --saida Minha.xlsx # escolhe o nome do arquivo

Todos os números do Dashboard, Orçamento, Anual e Histórico são fórmulas vivas do Excel:
o script só escreve os lançamentos, as metas e as listas de configuração.
Fórmulas em inglês com vírgula como separador (o Excel pt-BR traduz ao abrir).
"""
import argparse
import collections
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
_ap.add_argument("tema_pos", nargs="?", choices=["escuro", "claro"], help=argparse.SUPPRESS)
_ap.add_argument("--tema", choices=["escuro", "claro"], help="tema visual (padrão: escuro)")
_ap.add_argument("--dados", help="planilha exportada do app (Financas_Plinio_*.xlsx) para importar")
_ap.add_argument("--saida", help="nome do arquivo gerado")
ARGS = _ap.parse_args() if __name__ == "__main__" else _ap.parse_args([])

TEMA = ARGS.tema or ARGS.tema_pos or "escuro"   # "escuro" ou "claro"
NOME_ARQUIVO = "Financas_Pessoais_Dashboard.xlsx"
FONTE = "Calibri"

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


FMT_COMPETENCIA = "mm/yyyy"

# Colunas da tabela tbLancamentos (ordem na aba)
COLS_LANC = ["Data", "Descrição", "Tipo", "Categoria", "Subcategoria", "Pote", "Conta",
             "Forma de pagamento", "Parcela", "Valor", "Situação", "Vencimento", "Descontar de",
             "Competência", "Mês", "Ano", "Conta no mês?", "Observação"]
LARG_LANC = [12, 30, 12, 22, 16, 12, 17, 18, 8, 14, 16, 12, 12, 12, 6, 7, 9, 32]
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
    for cat, _ in cat_despesa:
        media = soma.get(cat, 0) / meses
        if media >= 1:
            orc[cat] = int(math.ceil(media / 50.0) * 50)
    return dict(sorted(orc.items(), key=lambda kv: -kv[1]))


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
        investimentos=None, acertos=None,
        subtitulo_lanc="Linhas com \"EXEMPLO\" na Observação são dados fictícios: filtre e apague quando quiser.",
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
    cat_despesa = [(c, ", ".join(k for k, _ in exemplos[c].most_common(5))) for c in cats_des]

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
        investimentos=investimentos, cab_investimentos=cab_inv, data_posicao=data_pos,
        acertos=acertos, rateios=rateios, totais_acerto=totais_acerto,
        subtitulo_lanc=(f"Importado do app em {data_pos:%d/%m/%Y}. " if data_pos else "") +
                       "Só entra nos totais o que tem 'Conta no mês?' = sim (receita recebida; gasto pago ou a descontar; nunca investimento, pendente ou projetado).",
        info_importacao=dict(linhas=len(lanc), sugeridas=sugeridas, sem_categoria=sem_cat),
    )


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
def construir(D, caminho):
    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard"
    ws_lanc = wb.create_sheet("Lancamentos")
    ws_orc = wb.create_sheet("Orcamento")
    ws_anual = wb.create_sheet("Anual")
    ws_hist = wb.create_sheet("Historico")
    ws_metas = wb.create_sheet("Metas")
    ws_inv = wb.create_sheet("Investimentos") if D.get("investimentos") else None
    ws_ac = wb.create_sheet("Acertos") if D.get("acertos") else None
    ws_cfg = wb.create_sheet("Config")
    ws_calc = wb.create_sheet("Calc")
    L = "tbLancamentos"
    CL = get_column_letter
    CONTA = f'{L}[Conta no mês?],"sim"'

    # ---------------------------------------------------------------- Config
    ws = ws_cfg
    pintar_fundo(ws, 45, 3)
    titulo_aba(ws, "Configurações", "Listas editáveis. Para adicionar um item, digite na primeira linha vazia logo abaixo da tabela: ela cresce sozinha e os menus suspensos acompanham.", 26)
    ws.column_dimensions["A"].width = 2
    L0 = 5  # linha de cabeçalho das tabelas de config
    cfg = [
        ("tbCatReceita", "RECEITAS", ["Categoria de receita"], [[c] for c in D["cat_receita"]], [24]),
        ("tbCatDespesa", "GASTOS", ["Categoria de gasto", "Subcategorias / exemplos"], D["cat_despesa"], [26, 56]),
        ("tbCatOutras", "INVESTIMENTOS / OUTRAS", ["Categoria (fora dos gastos)"], [[c] for c in D["cat_outras"]], [26]),
        ("tbPotes", "POTES", ["Pote"], [[p] for p in D["potes"]], [16]),
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
                      ("Potes", "tbPotes"), ("ListaContas", "tbContas"), ("ListaCartoes", "tbCartoes"),
                      ("FormasPagamento", "tbFormas"), ("Tipos", "tbTipos"), ("Situacoes", "tbSituacoes"),
                      ("Anos", "tbAnos")]:
        nome_dinamico(nome, tab)
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
    titulo_aba(ws, "Lançamentos", "Lance cada receita ou gasto numa linha (valor sempre positivo — o Tipo define se entra ou sai). " + D["subtitulo_lanc"], ncol)
    ws.column_dimensions["A"].width = 2
    for h, w in zip(COLS_LANC, LARG_LANC):
        ws.column_dimensions[cl[h]].width = w
    tr = lambda c: f"{L}[[#This Row],[{c}]]"
    f_mes = f'=IF(AND({tr("Data")}="",{tr("Competência")}=""),"",MONTH(IF({tr("Competência")}<>"",{tr("Competência")},{tr("Data")})))'
    f_ano = f'=IF(AND({tr("Data")}="",{tr("Competência")}=""),"",YEAR(IF({tr("Competência")}<>"",{tr("Competência")},{tr("Data")})))'
    f_conta = (f'=IF({tr("Valor")}="","",IF(AND({tr("Tipo")}<>"Investimento",OR({tr("Situação")}="",{tr("Situação")}="Pago",'
               f'{tr("Situação")}="Recebido",{tr("Situação")}="Descontar Depois")),"sim","não"))')
    formulas_lanc = {"Mês": f_mes, "Ano": f_ano, "Conta no mês?": f_conta}
    for h in COLS_LANC:
        ws.cell(H, ci[h], h)
    estilo_cabecalho(ws, H, 2, ncol + 1)
    estilo_corpo(ws, H + 1, ult, 2, ncol + 1)
    fmts = {"Data": FMT_DATA, "Vencimento": FMT_DATA, "Competência": FMT_COMPETENCIA, "Valor": FMT_MOEDA}
    suaves = ("Mês", "Ano", "Conta no mês?", "Parcela", "Competência")
    for i, d in enumerate(lanc):
        r = H + 1 + i
        for h in COLS_LANC:
            c = ws.cell(r, ci[h])
            if h in formulas_lanc:
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
    nova_tabela(ws, L, f"B{H}:{CL(ncol + 1)}{ult}", COLS_LANC, formulas=formulas_lanc)
    for r in range(ult + 1, ult + 1001):           # formatos prontos para as próximas linhas
        for h, f in fmts.items():
            ws.cell(r, ci[h]).number_format = f
    ws.freeze_panes = f"{cl['Tipo']}{H + 1}"
    LIM = ult + 20000
    faixa = lambda h: f"{cl[h]}{H + 1}:{cl[h]}{LIM}"
    dv_lista(ws, "=Tipos", faixa("Tipo"), msg="Escolha um tipo cadastrado na aba Config.")
    dv_lista(ws, "=Categorias", faixa("Categoria"), msg="Escolha uma categoria cadastrada na aba Config.")
    dv_lista(ws, "=Potes", faixa("Pote"), msg="Escolha um pote cadastrado na aba Config.")
    dv_lista(ws, "=Contas", faixa("Conta"), msg="Escolha uma conta ou cartão cadastrado na aba Config.")
    dv_lista(ws, "=FormasPagamento", faixa("Forma de pagamento"), msg="Escolha uma forma de pagamento cadastrada na aba Config.")
    dv_lista(ws, "=Situacoes", faixa("Situação"), msg="Escolha uma situação cadastrada na aba Config.")
    dv_val = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv_val.error, dv_val.errorTitle, dv_val.showErrorMessage = "Digite um valor positivo: o Tipo define se é entrada ou saída.", "Valor inválido", True
    ws.add_data_validation(dv_val)
    dv_val.add(faixa("Valor"))
    dv_dt = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True)
    dv_dt.error, dv_dt.errorTitle, dv_dt.showErrorMessage = "Digite uma data válida (dd/mm/aaaa).", "Data inválida", True
    ws.add_data_validation(dv_dt)
    for h in ("Data", "Vencimento", "Competência"):
        dv_dt.add(faixa(h))
    t1 = f"${cl['Tipo']}{H + 1}"
    ws.conditional_formatting.add(faixa("Valor"), FormulaRule(formula=[f'{t1}="Receita"'], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(faixa("Tipo"), FormulaRule(formula=[f'{t1}="Receita"'], font=Font(color="FF" + P["positivo"])))
    ws.conditional_formatting.add(faixa("Tipo"), FormulaRule(formula=[f'{t1}="Investimento"'], font=Font(color="FF" + P["destaque"])))
    ws.conditional_formatting.add(faixa("Tipo"), FormulaRule(formula=[f'AND({t1}<>"",{t1}<>"Receita",{t1}<>"Investimento")'], font=Font(color="FF" + P["negativo"])))
    s1 = f"${cl['Situação']}{H + 1}"
    ws.conditional_formatting.add(faixa("Situação"), FormulaRule(formula=[f'{s1}="Pendente"'], font=Font(color="FF" + P["alerta"], bold=True)))

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

    # ---------------------------------------------------------------- Orçamento
    ws = ws_orc
    pintar_fundo(ws, 10, 3)
    titulo_aba(ws, "Orçamento", "Defina a meta mensal de cada categoria de gasto. O realizado acompanha o mês/ano escolhido no Dashboard.", 7)
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
    put(ws, "H10", D["nota_orcamento"], font=fnt(9, color=P["suave"], italic=True),
        align=Alignment(wrap_text=True, vertical="top"))
    ws.merge_cells("H10:H13")
    ws.freeze_panes = f"A{HO + 1}"

    # ---------------------------------------------------------------- Anual
    ws = ws_anual
    pintar_fundo(ws, 20, 3)
    titulo_aba(ws, "Visão anual", "Categorias × meses do ano selecionado no Dashboard. Linhas vazias estão reservadas para categorias novas cadastradas na aba Config.", 17)
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
    tipos_gasto = [t for t in D["tipos"] if t not in ("Receita", "Investimento")]
    cab_h = ["Mês", "Receitas"] + tipos_gasto + (["Total de gastos"] if len(tipos_gasto) > 1 else []) + ["Resultado", "Taxa de poupança"]
    nh = len(cab_h)
    pintar_fundo(ws, nh + 3, 3)
    titulo_aba(ws, "Histórico mês a mês", "Todos os meses desde o primeiro lançamento. Mesma regra do Dashboard: só entra o que tem 'Conta no mês?' = sim.", nh)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 12
    for j in range(1, nh):
        ws.column_dimensions[CL(2 + j)].width = 17
    HH = 5
    for j, h in enumerate(cab_h):
        ws.cell(HH, 2 + j, h)
    estilo_cabecalho(ws, HH, 2, nh + 1)
    for j in range(1, nh):
        ws.cell(HH, 2 + j).alignment = Alignment(horizontal="right", vertical="center", indent=1, wrap_text=True)
    comps = [_competencia(d) for d in lanc]
    (a0, m0), (a1, m1) = min(comps), max(comps)
    meses_hist = []
    a, m = a0, m0
    while (a, m) <= (a1, m1):
        meses_hist.append((a, m))
        a, m = _mais_meses(a, m, 1)
    estilo_corpo(ws, HH + 1, HH + len(meses_hist), 2, nh + 1)
    col_rec = CL(3)
    col_tot = CL(3 + len(tipos_gasto)) if len(tipos_gasto) > 1 else CL(3)
    for i, (a, m) in enumerate(meses_hist):
        r = HH + 1 + i
        ws.cell(r, 2, dt.date(a, m, 1)).number_format = FMT_COMPETENCIA
        base = f'{CONTA},{L}[Mês],MONTH($B{r}),{L}[Ano],YEAR($B{r})'
        ws.cell(r, 3, f'=SUMIFS({L}[Valor],{L}[Tipo],"Receita",{base})')
        for j, t in enumerate(tipos_gasto):
            ws.cell(r, 4 + j, f'=SUMIFS({L}[Valor],{L}[Tipo],"{t}",{base})')
        k = 4 + len(tipos_gasto)
        if len(tipos_gasto) > 1:
            ws.cell(r, k, f"=SUM({CL(4)}{r}:{CL(3 + len(tipos_gasto))}{r})")
            col_tot = CL(k)
            k += 1
        else:
            col_tot = CL(4)
        ws.cell(r, k, f"={col_rec}{r}-{col_tot}{r}")
        ws.cell(r, k + 1, f"=IF({col_rec}{r}>0,{CL(k)}{r}/{col_rec}{r},0)")
        for colx in range(3, k + 1):
            ws.cell(r, colx).number_format = FMT_MOEDA_ZERO_TRACO
            ws.cell(r, colx).alignment = ALIGN_R
        ws.cell(r, k + 1).number_format = FMT_PCT
        ws.cell(r, k + 1).alignment = ALIGN_R
        ws.cell(r, k).font = fnt(10, True)
        if m == 12:
            for colx in range(2, nh + 2):
                ws.cell(r, colx).border = Border(bottom=side(P["destaque"]))
    col_res = CL(4 + len(tipos_gasto) + (1 if len(tipos_gasto) > 1 else 0))
    fim_h = HH + len(meses_hist)
    ws.conditional_formatting.add(f"{col_res}{HH + 1}:{col_res}{fim_h}", FormulaRule(formula=[f"{col_res}{HH + 1}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(f"{col_res}{HH + 1}:{col_res}{fim_h}", FormulaRule(formula=[f"{col_res}{HH + 1}>0"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.freeze_panes = f"C{HH + 1}"

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
    f_pa = "=IF(N(tbMetas[[#This Row],[Valor alvo]])>0,MIN(1,N(tbMetas[[#This Row],[Valor atual]])/tbMetas[[#This Row],[Valor alvo]]),0)"
    f_ap = ("=IFERROR(IF(N(tbMetas[[#This Row],[Valor atual]])>=N(tbMetas[[#This Row],[Valor alvo]]),0,"
            "(tbMetas[[#This Row],[Valor alvo]]-N(tbMetas[[#This Row],[Valor atual]]))/"
            "MAX(1,(YEAR(tbMetas[[#This Row],[Prazo]])-YEAR(TODAY()))*12+MONTH(tbMetas[[#This Row],[Prazo]])-MONTH(TODAY()))),0)")
    metas = D["metas"] or [(None, None, None, None)]
    for j, h in enumerate(cab_m):
        ws.cell(HM, 2 + j, h)
    estilo_cabecalho(ws, HM, 2, 7)
    ws.row_dimensions[HM].height = 32
    estilo_corpo(ws, HM + 1, HM + len(metas) + 1, 2, 7)
    for i, (nome, alvo, atual, prazo) in enumerate(metas):
        r = HM + 1 + i
        ws.cell(r, 2, nome)
        ws.cell(r, 3, alvo)
        ws.cell(r, 4, atual)
        ws.cell(r, 5, f_pa)
        ws.cell(r, 6, prazo)
        ws.cell(r, 7, f_ap)
    tm = HM + len(metas) + 1
    f_pa_tot = "=IF(SUM(tbMetas[Valor alvo])>0,SUM(tbMetas[Valor atual])/SUM(tbMetas[Valor alvo]),0)"
    ws.cell(tm, 2, "Total")
    ws.cell(tm, 3, "=SUBTOTAL(109,tbMetas[Valor alvo])")
    ws.cell(tm, 4, "=SUBTOTAL(109,tbMetas[Valor atual])")
    ws.cell(tm, 5, f_pa_tot)
    ws.cell(tm, 7, "=SUBTOTAL(109,tbMetas[Aporte mensal necessário])")
    area(ws, f"B{tm}:G{tm}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for r in range(HM + 1, tm + 1):
        for colx in "CDG":
            ws[f"{colx}{r}"].number_format = FMT_MOEDA
            ws[f"{colx}{r}"].alignment = ALIGN_R
        ws[f"E{r}"].number_format = FMT_PCT
        ws[f"E{r}"].alignment = ALIGN_R
        ws[f"F{r}"].number_format = FMT_DATA
        ws[f"F{r}"].alignment = ALIGN_C
    nova_tabela(ws, "tbMetas", f"B{HM}:G{tm}", cab_m, formulas={"% atingido": f_pa, "Aporte mensal necessário": f_ap},
                totais={"Nome": "Total", "Valor alvo": "sum", "Valor atual": "sum", "Aporte mensal necessário": "sum"})
    tcm = ws.tables["tbMetas"].tableColumns[3]
    tcm.totalsRowFunction = "custom"
    tcm.totalsRowFormula = TableFormula(attr_text=f_pa_tot.lstrip("="))
    ws.conditional_formatting.add(f"E{HM + 1}:E{tm}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                                                   color="FF" + BARRA_META, showValue=True))

    cab_d = ["Descrição", "Valor da parcela", "Parcelas pagas", "Total de parcelas", "Saldo devedor", "Data de término"]
    f_sd = "=N(tbDividas[[#This Row],[Valor da parcela]])*MAX(0,N(tbDividas[[#This Row],[Total de parcelas]])-N(tbDividas[[#This Row],[Parcelas pagas]]))"
    dividas = D["dividas"] or [(None, None, None, None, None)]
    for j, h in enumerate(cab_d):
        ws.cell(HM, 9 + j, h)
    estilo_cabecalho(ws, HM, 9, 14)
    estilo_corpo(ws, HM + 1, HM + len(dividas) + 1, 9, 14)
    for i, (desc, parc, pagas, total, fim) in enumerate(dividas):
        r = HM + 1 + i
        for j, v in enumerate((desc, parc, pagas, total)):
            ws.cell(r, 9 + j, v)
        ws.cell(r, 13, f_sd)
        ws.cell(r, 14, fim)
    td = HM + len(dividas) + 1
    ws.cell(td, 9, "Total")
    ws.cell(td, 10, "=SUBTOTAL(109,tbDividas[Valor da parcela])")
    ws.cell(td, 13, "=SUBTOTAL(109,tbDividas[Saldo devedor])")
    area(ws, f"I{td}:N{td}", fill_=CARD2, font=fnt(10, True), border=Border(top=side(P["destaque"])))
    for r in range(HM + 1, td + 1):
        for colx in "JM":
            ws[f"{colx}{r}"].number_format = FMT_MOEDA
            ws[f"{colx}{r}"].alignment = ALIGN_R
        for colx in "KL":
            ws[f"{colx}{r}"].alignment = ALIGN_C
        ws[f"N{r}"].number_format = FMT_DATA
        ws[f"N{r}"].alignment = ALIGN_C
    nova_tabela(ws, "tbDividas", f"I{HM}:N{td}", cab_d, formulas={"Saldo devedor": f_sd},
                totais={"Descrição": "Total", "Valor da parcela": "sum", "Saldo devedor": "sum"})
    nota_div = ("Parcelamentos em andamento encontrados nos lançamentos de "
                f"{MESES[D['mes_padrao'] - 1].lower()}/{D['ano_padrao']}. " if not D["exemplo"] and D["dividas"] else "")
    put(ws, f"I{td + 2}", nota_div + "As parcelas continuam sendo lançadas na aba Lançamentos quando forem pagas.",
        font=fnt(9, color=P["suave"], italic=True))
    put(ws, f"B{tm + 2}", "Aporte mensal necessário = (valor alvo − valor atual) ÷ meses restantes até o prazo."
        + ("" if D["metas"] else " Cadastre suas metas na linha vazia da tabela."),
        font=fnt(9, color=P["suave"], italic=True))

    # ---------------------------------------------------------------- Investimentos
    if ws_inv is not None:
        ws = ws_inv
        cab_i = D["cab_investimentos"]
        inv = D["investimentos"]
        ni = len(cab_i)
        pintar_fundo(ws, ni + 8, 3)
        pos = D.get("data_posicao")
        titulo_aba(ws, "Investimentos", ("Saldo hoje, IR e Líquido: posição calculada pelo app em "
                                         f"{pos:%d/%m/%Y}" if pos else "Saldo hoje, IR e Líquido: posição calculada pelo app")
                   + " (valores fixos, não se atualizam sozinhos). Totais ao lado são fórmulas.", ni)
        ws.column_dimensions["A"].width = 2
        larg_i = {"Data": 12, "Banco": 11, "% do CDI": 10, "Aplicado": 15, "Dias": 7, "Isento": 8, "Resgatado": 10,
                  "Alíquota IR": 10, "Saldo hoje": 15, "IR": 12, "Líquido": 15, "Vencimento": 12, "Observação": 40}
        HI = 5
        for j, h in enumerate(cab_i):
            ws.cell(HI, 2 + j, h)
            ws.column_dimensions[CL(2 + j)].width = larg_i.get(h, 12)
        estilo_cabecalho(ws, HI, 2, ni + 1)
        estilo_corpo(ws, HI + 1, HI + len(inv), 2, ni + 1)
        fm = {"Data": FMT_DATA, "Vencimento": FMT_DATA, "% do CDI": FMT_PCT, "Alíquota IR": FMT_PCT,
              "Aplicado": FMT_MOEDA, "Saldo hoje": FMT_MOEDA, "IR": FMT_MOEDA, "Líquido": FMT_MOEDA}
        for i, row in enumerate(inv):
            r = HI + 1 + i
            for j, (h, v) in enumerate(zip(cab_i, row)):
                cc = ws.cell(r, 2 + j, v)
                if h in fm:
                    cc.number_format = fm[h]
                    cc.alignment = ALIGN_R if fm[h] != FMT_DATA else ALIGN_C
                if h in ("Dias", "Isento", "Resgatado"):
                    cc.alignment = ALIGN_C
            if str(row[cab_i.index("Resgatado")] or "").lower() == "sim":
                for j in range(ni):
                    ws.cell(r, 2 + j).font = fnt(10, color=P["suave"])
        nova_tabela(ws, "tbInvestimentos", f"B{HI}:{CL(ni + 1)}{HI + len(inv)}", cab_i)
        cx = ni + 3   # coluna do resumo
        ws.column_dimensions[CL(cx - 1)].width = 3
        ws.column_dimensions[CL(cx)].width = 24
        ws.column_dimensions[CL(cx + 1)].width = 17
        put(ws, f"{CL(cx)}4", "CARTEIRA", font=fnt(10, True, P["suave"]))
        resumo = [("Carteira hoje (bruto)", "=SUM(tbInvestimentos[Saldo hoje])", FMT_MOEDA),
                  ("IR se resgatar hoje", "=SUM(tbInvestimentos[IR])", FMT_MOEDA),
                  ("Líquido", "=SUM(tbInvestimentos[Líquido])", FMT_MOEDA),
                  ("Total aplicado (ativos)", '=SUMIFS(tbInvestimentos[Aplicado],tbInvestimentos[Resgatado],"<>sim")', FMT_MOEDA),
                  ("Rendimento bruto", f"={CL(cx + 1)}6-{CL(cx + 1)}9", FMT_MOEDA),
                  ("Rentabilidade sobre o aplicado", f"=IF({CL(cx + 1)}9>0,{CL(cx + 1)}10/{CL(cx + 1)}9,0)", FMT_PCT)]
        for i, (rot, f, fmt_) in enumerate(resumo):
            r = 6 + i
            ws.row_dimensions[r].height = 22
            put(ws, f"{CL(cx)}{r}", rot, font=fnt(10, color=P["suave"]), fill_=CARD, align=ALIGN_L)
            put(ws, f"{CL(cx + 1)}{r}", f, font=fnt(11, True, P["positivo"] if i in (0, 2, 4) else P["texto"]),
                fill_=CARD, fmt=fmt_, align=ALIGN_R)
        area(ws, f"{CL(cx)}5:{CL(cx + 1)}5", fill_=CARD, border=Border(top=side(P["positivo"], "thick")))
        ws.freeze_panes = f"C{HI + 1}"

    # ---------------------------------------------------------------- Acertos a cobrar
    if ws_ac is not None:
        ws = ws_ac
        cab_a = ["Pessoa", "Mês", "Data", "Descrição", "Parcela", "Valor cheio", "Já pago", "Falta"]
        itens = D["acertos"]
        pintar_fundo(ws, 16, 3)
        titulo_aba(ws, "Acertos a cobrar", "Valores a receber de outras pessoas. Atualize \"Já pago\" quando receberem; o resumo ao lado recalcula.", 10)
        larguras(ws, {"A": 2, "B": 14, "C": 10, "D": 12, "E": 28, "F": 9, "G": 14, "H": 13, "I": 14, "J": 3, "K": 26, "L": 16})
        HAc = 5
        for j, h in enumerate(cab_a):
            ws.cell(HAc, 2 + j, h)
        estilo_cabecalho(ws, HAc, 2, 9)
        estilo_corpo(ws, HAc + 1, HAc + len(itens), 2, 9)
        f_falta = "=N(tbAcertos[[#This Row],[Valor cheio]])-N(tbAcertos[[#This Row],[Já pago]])"
        for i, row in enumerate(itens):
            r = HAc + 1 + i
            for j, v in enumerate(row):
                ws.cell(r, 2 + j, v)
            ws.cell(r, 9, f_falta)
            ws.cell(r, 4).number_format = FMT_DATA
            for colx in (7, 8, 9):
                ws.cell(r, colx).number_format = FMT_MOEDA
                ws.cell(r, colx).alignment = ALIGN_R
        nova_tabela(ws, "tbAcertos", f"B{HAc}:I{HAc + len(itens)}", cab_a, formulas={"Falta": f_falta})
        # Resumo por pessoa (+ rateio quando existir)
        pessoas = list(dict.fromkeys(row[0] for row in itens))
        put(ws, "K4", "TOTAL A COBRAR POR PESSOA", font=fnt(10, True, P["suave"]))
        linha = 6
        area(ws, "K5:L5", fill_=CARD, border=Border(top=side(P["alerta"], "thick")))
        ref_total = {}
        for p in pessoas:
            put(ws, f"K{linha}", p, font=fnt(10), fill_=CARD, align=ALIGN_L)
            ref_total[p] = f"L{linha}"
            put(ws, f"L{linha}", f'=SUMIFS(tbAcertos[Falta],tbAcertos[Pessoa],K{linha})', font=fnt(11, True, P["alerta"]),
                fill_=CARD, fmt=FMT_MOEDA, align=ALIGN_R)
            linha += 1
        linha += 1
        for p, vals in D.get("rateios", {}).items():
            put(ws, f"K{linha}", f"RATEIO — {p.upper()}", font=fnt(10, True, P["suave"]))
            linha += 1
            ini = linha
            passos = [("Aluguel dele (digite)", vals.get("aluguel dele")),
                      ("Aluguel dela (digite)", vals.get("aluguel dela")),
                      ("Gastos da casa", f'=SUMIFS(tbAcertos[Falta],tbAcertos[Pessoa],"{p}")'),
                      ("Lucro", f"=L{ini}+L{ini + 1}-L{ini + 2}"),
                      ("Metade do lucro", f"=L{ini + 3}/2"),
                      ("Total a cobrar", f"=L{ini + 1}-L{ini + 4}")]
            for j, (rot, v) in enumerate(passos):
                r = ini + j
                final = j == len(passos) - 1
                put(ws, f"K{r}", rot, font=fnt(10, final, P["texto"] if final else P["suave"]), fill_=CARD2 if final else CARD, align=ALIGN_L)
                put(ws, f"L{r}", v, font=fnt(11 if final else 10, final, P["alerta"] if final else (P["destaque"] if j < 2 else P["texto"])),
                    fill_=CARD2 if final else CARD, fmt=FMT_MOEDA, align=ALIGN_R)
            ws[ref_total[p]].value = f"=L{ini + len(passos) - 1}"
            put(ws, f"K{ini + len(passos) + 1}", "Mesma conta do arquivo original, passo a passo.", font=fnt(9, color=P["suave"], italic=True))
            linha = ini + len(passos) + 3
        ws.freeze_panes = f"A{HAc + 1}"

    # ---------------------------------------------------------------- Dashboard
    ws = ws_dash
    card_cols = []
    colx = 2
    for k in range(6):
        card_cols.append((colx, colx + 2))
        colx += 4
    ult_col = card_cols[-1][1] + 1
    pintar_fundo(ws, ult_col + 6, 80)
    ws.column_dimensions["A"].width = 3
    for c0, c1 in card_cols:
        for cc in range(c0, c1 + 1):
            ws.column_dimensions[CL(cc)].width = 10
        if c1 + 1 <= ult_col:
            ws.column_dimensions[CL(c1 + 1)].width = 2.5
    ws.column_dimensions[CL(ult_col)].width = 3
    ws.sheet_view.zoomScale = 90

    for rrow, h in {1: 12, 2: 26, 3: 28, 4: 16}.items():
        ws.row_dimensions[rrow].height = h
    put(ws, "B2", "Finanças da Família", font=fnt(20, True), align=Alignment(vertical="bottom"))
    put(ws, "B3", "Visão mensal — escolha o mês e o ano nos seletores à direita; tudo se atualiza sozinho.",
        font=fnt(9, color=P["suave"]), align=Alignment(vertical="center"))
    c5, c6 = card_cols[4], card_cols[5]
    put(ws, f"{CL(c5[0])}2", "MÊS", font=fnt(9, True, P["suave"]), align=Alignment(vertical="bottom", indent=1))
    put(ws, f"{CL(c6[0])}2", "ANO", font=fnt(9, True, P["suave"]), align=Alignment(vertical="bottom", indent=1))
    borda_sel = Border(bottom=side(P["destaque"], "medium"))
    mesclar(ws, f"{CL(c5[0])}3:{CL(c5[1])}3", MESES[D["mes_padrao"] - 1], fill_=CARD2, font=fnt(14, True, P["destaque"]),
            align=Alignment(horizontal="center", vertical="center"), border=borda_sel)
    mesclar(ws, f"{CL(c6[0])}3:{CL(c6[1])}3", D["ano_padrao"], fill_=CARD2, font=fnt(14, True, P["destaque"]),
            align=Alignment(horizontal="center", vertical="center"), border=borda_sel)
    wb.defined_names["MesSelecionado"] = DefinedName("MesSelecionado", attr_text=f"Dashboard!${CL(c5[0])}$3")
    wb.defined_names["AnoSelecionado"] = DefinedName("AnoSelecionado", attr_text=f"Dashboard!${CL(c6[0])}$3")
    dv_lista(ws, "=Meses", f"{CL(c5[0])}3", "Mês inválido", "Escolha um mês da lista.")
    dv_lista(ws, "=Anos", f"{CL(c6[0])}3", "Ano inválido", "Escolha um ano da lista (cadastre novos anos na aba Config).")
    ws.freeze_panes = "A5"

    for rrow, h in {5: 8, 6: 20, 7: 38, 8: 18, 9: 10, 10: 18}.items():
        ws.row_dimensions[rrow].height = h
    kpis = [
        ("RECEITAS DO MÊS", "=Calc!B9", FMT_MOEDA, "Mês anterior", "=Calc!B11", FMT_MOEDA, P["positivo"]),
        ("GASTOS DO MÊS", "=Calc!B10", FMT_MOEDA, "Mês anterior", "=Calc!B12", FMT_MOEDA, P["negativo"]),
        ("SALDO DO MÊS", "=Calc!B13", FMT_MOEDA, "Mês anterior", "=Calc!B14", FMT_MOEDA, P["destaque"]),
        ("TAXA DE POUPANÇA", "=Calc!B15", FMT_PCT, "Mês anterior", "=Calc!B19", FMT_PCT, P["roxo"]),
        ("GASTOS VS. MÊS ANTERIOR", "=Calc!B16", FMT_VAR, "Diferença", "=Calc!B25", FMT_MOEDA, P["alerta"]),
        ("% DO ORÇAMENTO UTILIZADO", "=Calc!B18", FMT_PCT, "Meta total", "=Calc!B17", FMT_MOEDA, P["destaque"]),
    ]
    kpi_cells = []
    for (c0, c1), (rot, f, fmt, sub, fsub, fmtsub, cor) in zip(card_cols, kpis):
        a, b = CL(c0), CL(c1)
        area(ws, f"{a}5:{b}9", fill_=CARD)
        area(ws, f"{a}5:{b}5", border=Border(top=side(cor, "thick")))
        mesclar(ws, f"{a}6:{b}6", rot, fill_=CARD, font=fnt(9, True, P["suave"]), align=Alignment(horizontal="left", vertical="center", indent=1))
        mesclar(ws, f"{a}7:{b}7", f, fill_=CARD, font=fnt(22, True), fmt=fmt, align=Alignment(horizontal="left", vertical="center", indent=1))
        put(ws, f"{a}8", sub, font=fnt(9, color=P["suave"]), fill_=CARD, align=Alignment(horizontal="left", vertical="center", indent=1))
        mesclar(ws, f"{CL(c0 + 1)}8:{b}8", fsub, fill_=CARD, font=fnt(9, True, P["suave"]), fmt=fmtsub, align=Alignment(horizontal="right", vertical="center", indent=1))
        kpi_cells.append(f"{a}7")
    sal, var, orc = kpi_cells[2], kpi_cells[4], kpi_cells[5]
    ws.conditional_formatting.add(sal, FormulaRule(formula=[f"{sal}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(sal, FormulaRule(formula=[f"{sal}>=0"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(kpi_cells[3], FormulaRule(formula=[f"{kpi_cells[3]}<0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(var, FormulaRule(formula=[f"{var}>0"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(var, FormulaRule(formula=[f"{var}<0"], font=Font(color="FF" + P["positivo"], bold=True)))
    ws.conditional_formatting.add(orc, FormulaRule(formula=[f"{orc}>1"], font=Font(color="FF" + P["negativo"], bold=True)))
    ws.conditional_formatting.add(orc, FormulaRule(formula=[f"AND({orc}>0.8,{orc}<=1)"], font=Font(color="FF" + P["alerta"], bold=True)))
    ws.conditional_formatting.add(orc, FormulaRule(formula=[f"{orc}<=0.8"], font=Font(color="FF" + P["positivo"], bold=True)))

    esq = (card_cols[0][0] - 1, card_cols[2][1])
    dir_ = (card_cols[3][0] - 1, card_cols[5][1])
    G1_INI, G1_FIM = 10, 29
    G2_INI, G2_FIM = 30, 49
    for rrow in range(11, 50):
        ws.row_dimensions[rrow].height = 18 if N_ORC <= 12 else 20
    ws.row_dimensions[30].height = 14

    ch = DoughnutChart(holeSize=58)
    ch.add_data(Reference(ws_calc, min_col=24, min_row=2, max_row=9), titles_from_data=False)
    ch.set_categories(Reference(ws_calc, min_col=23, min_row=2, max_row=9))
    s = ch.series[0]
    for i in range(8):
        pt = DataPoint(idx=i)
        pt.graphicalProperties = GraphicalProperties(solidFill=CORES_CATEGORIAS[i % len(CORES_CATEGORIAS)],
                                                     ln=LineProperties(solidFill=P["card"], w=19050))
        s.dPt.append(pt)
    estilizar_grafico(ch, "Gastos por categoria no mês")
    ch.legend.position = "r"
    ch.firstSliceAng = 0
    ancorar(ws, ch, esq[0], G1_INI, esq[1], G1_FIM)

    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.gapWidth = 70
    ch.overlap = -10
    ch.add_data(Reference(ws_calc, min_col=6, max_col=7, min_row=1, max_row=13), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=5, min_row=2, max_row=13))
    for s, cor in zip(ch.series, (P["positivo"], P["negativo"])):
        s.graphicalProperties = GraphicalProperties(solidFill=cor, ln=LineProperties(noFill=True))
    estilizar_grafico(ch, "Receitas × Gastos — últimos 12 meses")
    estilizar_eixos(ch)
    ancorar(ws, ch, dir_[0], G1_INI, dir_[1], G1_FIM)

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

    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "clustered"
    ch.gapWidth = 40
    ch.overlap = -5
    ch.add_data(Reference(ws_calc, min_col=27, max_col=28, min_row=1, max_row=1 + N_ORC), titles_from_data=True)
    ch.set_categories(Reference(ws_calc, min_col=26, min_row=2, max_row=1 + N_ORC))
    for s, cor in zip(ch.series, (P["suave"], P["destaque"])):
        s.graphicalProperties = GraphicalProperties(solidFill=cor, ln=LineProperties(noFill=True))
    estilizar_grafico(ch, "Orçado × Realizado por categoria")
    estilizar_eixos(ch)
    if N_ORC > 12:
        ch.x_axis.txPr = txpr(P["suave"], 750)
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.crosses = "max"
    ancorar(ws, ch, dir_[0], G2_INI, dir_[1], G2_FIM)

    T0 = 51
    ws.row_dimensions[50].height = 14
    ws.row_dimensions[T0].height = 26
    ws.row_dimensions[T0 + 1].height = 24
    put(ws, f"B{T0}", "Top 5 maiores gastos do mês", font=fnt(12, True), align=Alignment(vertical="center"))
    for rot, a, b in [("Data", "B", "C"), ("Descrição", "D", "H"), ("Categoria", "I", "J"), ("Valor", "K", "L")]:
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

    put(ws, f"N{T0}", "Orçado × Realizado por categoria", font=fnt(12, True), align=Alignment(vertical="center"))
    for rot, a, b in [("Categoria", "N", "P"), ("Meta", "Q", "R"), ("Realizado", "S", "T"), ("Diferença", "U", "V"), ("% usado", "W", "X")]:
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
    put(ws, f"B{T0 + 8}", "Como usar", font=fnt(10, True, P["suave"]))
    dicas = ["1. Lance receitas e gastos na aba Lançamentos (valor sempre positivo).",
             "2. Troque o mês/ano no topo: tudo se atualiza sozinho.",
             "3. Ajuste metas na aba Orçamento e listas na aba Config.",
             "4. Só entra nos totais o que tem 'Conta no mês?' = sim.",
             "% usado: verde até 80% • amarelo 80–100% • vermelho acima de 100%."]
    for i, dtxt in enumerate(dicas):
        put(ws, f"B{T0 + 9 + i}", dtxt, font=fnt(9, color=P["suave"]))

    ws.print_area = f"A1:{CL(ult_col)}{max(fim_orc, T0 + 14) + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.3

    for w in (ws_lanc, ws_orc, ws_anual, ws_hist, ws_metas, ws_inv, ws_ac, ws_cfg):
        if w is None:
            continue
        w.page_setup.orientation = "landscape"
        w.page_setup.paperSize = w.PAPERSIZE_A4
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
        w.page_margins.left = w.page_margins.right = 0.3

    cores_guia = [(ws_dash, P["destaque"]), (ws_lanc, P["positivo"]), (ws_orc, P["alerta"]), (ws_anual, P["roxo"]),
                  (ws_hist, P["roxo"]), (ws_metas, "2EC4B6"), (ws_inv, P["positivo"]), (ws_ac, P["alerta"]),
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
    pasta = os.path.dirname(os.path.abspath(__file__))
    if ARGS.dados:
        dados = importar_plinio(ARGS.dados)
        info = dados["info_importacao"]
        print(f"Importados {info['linhas']} lançamentos ({info['sugeridas']} com categoria sugerida pelo histórico, "
              f"{info['sem_categoria']} 'Sem categoria').")
    else:
        dados = dados_exemplo()
    destino = ARGS.saida or os.path.join(pasta, NOME_ARQUIVO if ARGS.dados else "Financas_Pessoais_Dashboard_EXEMPLO.xlsx")
    construir(dados, destino)
    print(f"Planilha gerada: {destino} (tema {TEMA})")
