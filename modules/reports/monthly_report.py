import os
import unicodedata
from datetime import datetime

import pandas as pd
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.legends import Legend
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle

from modules.processors.data_processor import get_weekly_data

NAVY = colors.HexColor("#1B3A5C")
GOLD = colors.HexColor("#F4A720")
BLUE_CARD = colors.HexColor("#1565C0")
GREEN_CARD = colors.HexColor("#2E7D32")
PURPLE_CARD = colors.HexColor("#6A1B9A")
ORANGE_CARD = colors.HexColor("#BF360C")
SEC_MACH = colors.HexColor("#0D3B6E")
SEC_VEH = colors.HexColor("#1B5E20")
ROW_EVEN = colors.HexColor("#EEF2F7")
TOTAL_BG = colors.HexColor("#263238")
WHITE = colors.white
GRAY_TEXT = colors.HexColor("#78909C")
LIGHT_BLUE = colors.HexColor("#B0BEC5")

LOCS_HORA = ["LOC 01", "LOC 02", "LOC 05", "LOC 08"]
LISTA_VERMELHA = ["MC 01", "MC 13"] + LOCS_HORA


def _fmt_hours(v):
    h = int(v)
    m = int(round((v - h) * 60))
    return f"{h}:{m:02d}h"


def _is_time(name):
    n = name.upper()
    return "MC" in n or any(loc in n for loc in LOCS_HORA)


def _is_alert(name, hours):
    return any(x in name.upper() for x in LISTA_VERMELHA) or hours == 0


def _norm_name(name):
    text = unicodedata.normalize("NFKD", str(name).upper())
    return "".join(c for c in text if not unicodedata.combining(c))


def _is_retro(name):
    return "RETRO" in _norm_name(name)


def _is_caminhao(name):
    n = _norm_name(name)
    return "CAMINHAO" in n or "TRUCK" in n


def _is_veiculo(name):
    n = _norm_name(name)
    keys = ["VEICULO", "CARRO", "PICKUP", "VIATURA", "UTILITARIO", "SUV", "MOTO"]
    return any(k in n for k in keys)


def _top3(items, matcher):
    filtered = [i for i in items if matcher(i["display"]) and i["hours"] > 0]
    return sorted(filtered, key=lambda x: x["hours"], reverse=True)[:3]


def _build_top3_category_chart(width, title, rows, as_hours, bar_color):
    chart_h = 52 * mm
    drawing = Drawing(width, chart_h)
    drawing.add(Rect(0, 0, width, chart_h, fillColor=colors.HexColor("#F5F7FA"), strokeColor=colors.HexColor("#CFD8E3")))
    drawing.add(String(width / 2, chart_h - 11, title, fontName="Helvetica-Bold", fontSize=9, fillColor=NAVY, textAnchor="middle"))

    if not rows:
        drawing.add(String(width / 2, chart_h / 2, "Sem dados para este periodo", fontName="Helvetica", fontSize=9, fillColor=GRAY_TEXT, textAnchor="middle"))
        return drawing

    labels = [f"{idx}. {str(item['display'])[:10]}" for idx, item in enumerate(rows, start=1)]
    values = [float(item["hours"]) for item in rows]

    bar = VerticalBarChart()
    bar.x = 10 * mm
    bar.y = 10 * mm
    bar.height = 30 * mm
    bar.width = width - 20 * mm
    bar.data = [values]
    bar.valueAxis.valueMin = 0
    bar.valueAxis.valueMax = _round_axis_max(max(values) * 1.2)
    bar.valueAxis.valueStep = max(1, int(bar.valueAxis.valueMax / 4))
    bar.valueAxis.labels.fontSize = 7
    bar.categoryAxis.categoryNames = labels
    bar.categoryAxis.labels.fontSize = 6.5
    bar.categoryAxis.labels.boxAnchor = "n"
    bar.bars[0].fillColor = bar_color
    bar.bars[0].strokeColor = bar_color
    bar.barSpacing = 8
    bar.groupSpacing = 12
    bar.barLabelFormat = "%0.0f"
    bar.barLabels.nudge = 6
    bar.barLabels.fontSize = 7
    bar.barLabels.fillColor = colors.HexColor("#37474F")
    bar.barLabels.boxAnchor = "s"
    bar.barLabels.visible = True
    drawing.add(bar)

    unit = "horas" if as_hours else "km"
    drawing.add(String(width - 8, chart_h - 11, unit, fontName="Helvetica", fontSize=7, fillColor=GRAY_TEXT, textAnchor="end"))
    return drawing


def _p(text, size=9, bold=False, color="#000000", align=TA_LEFT):
    return Paragraph(text, ParagraphStyle(
        "auto",
        fontName="Helvetica-Bold" if bold else "Helvetica",
        fontSize=size,
        textColor=colors.HexColor(color),
        alignment=align,
        leading=size * 1.5,
        leftPadding=2,
        rightPadding=2,
        spaceBefore=0,
        spaceAfter=0,
    ))


def _round_axis_max(max_value):
    if max_value <= 0:
        return 10
    for step in [10, 25, 50, 100, 250, 500, 1000]:
        if max_value <= step:
            return step
    return int(((max_value // 1000) + 1) * 1000)


def create_monthly_report(excel_path, obs_sheet_name, output_folder, selected_sheets=None, weekly_data_loader=None):
    try:
        xl = pd.ExcelFile(excel_path)
        available_sheets = xl.sheet_names
    except Exception as e:
        print(f"ERRO ao abrir Excel: {e}")
        return None

    if selected_sheets:
        valid_sheets = []
        for sheet in selected_sheets:
            sheet_name = str(sheet or "").strip()
            if sheet_name and sheet_name in available_sheets and sheet_name != obs_sheet_name and sheet_name not in valid_sheets:
                valid_sheets.append(sheet_name)
    else:
        valid_sheets = [s for s in available_sheets if s != obs_sheet_name]

    if not valid_sheets:
        print("ERRO: Nenhuma aba semanal encontrada.")
        return None

    print(f"[*] Abas semanais: {valid_sheets}")

    weekly_data_loader = weekly_data_loader or get_weekly_data
    machine_totals = {}
    for sheet in valid_sheets:
        print(f"    -> Lendo: {sheet}")
        for item in weekly_data_loader(excel_path, sheet):
            key = " ".join(item["machine"].upper().split())
            if key not in machine_totals:
                machine_totals[key] = {
                    "display": item["machine"],
                    "hours": 0.0,
                    "operator": item["operator"],
                }
            machine_totals[key]["hours"] += item["hours"]

    if not machine_totals:
        print("Sem dados.")
        return None

    aggregated = list(machine_totals.values())
    machines = [i for i in aggregated if _is_time(i["display"])]
    vehicles = [i for i in aggregated if not _is_time(i["display"])]
    total_h = sum(i["hours"] for i in machines)
    total_km = sum(i["hours"] for i in vehicles)
    weeks = len(valid_sheets)
    n_equip = len(aggregated)
    period = f"{valid_sheets[0]}  ->  {valid_sheets[-1]}" if weeks > 1 else valid_sheets[0]
    gen_date = datetime.now().strftime("%d/%m/%Y  %H:%M")
    return _build_pdf(machines, vehicles, total_h, total_km, weeks, n_equip, period, gen_date, output_folder, valid_sheets)


def _build_pdf(machines, vehicles, total_h, total_km, weeks, n_equip, period, gen_date, output_folder, valid_sheets):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    logo_path = os.path.join("assets", "company", "company_2.png")
    safe = valid_sheets[-1].replace(" ", "_").replace("/", "-")
    path = os.path.join(output_folder, f"Report_MENSAL_{safe}.pdf")

    page_w, page_h = A4
    header_h = 55 * mm
    footer_h = 14 * mm
    margin = 12 * mm
    deco = 3
    frame_w = page_w - 2 * margin

    frame = Frame(
        margin,
        footer_h + 4 * mm,
        frame_w,
        page_h - header_h - deco - footer_h - 12 * mm,
        id="main",
        leftPadding=0,
        rightPadding=0,
        topPadding=5 * mm,
        bottomPadding=0,
    )

    def on_page(canv, doc):
        canv.saveState()
        canv.setFillColor(NAVY)
        canv.rect(0, page_h - header_h, page_w, header_h, fill=1, stroke=0)
        canv.setFillColor(GOLD)
        canv.rect(0, page_h - header_h - deco, page_w, deco, fill=1, stroke=0)

        if os.path.exists(logo_path):
            canv.drawImage(logo_path, 12 * mm, page_h - header_h + 8 * mm, width=35 * mm, height=35 * mm, preserveAspectRatio=True, mask="auto")

        canv.setFillColor(WHITE)
        canv.setFont("Helvetica-Bold", 20)
        canv.drawCentredString(page_w / 2, page_h - 22 * mm, "RELATORIO MENSAL DE PRODUCAO")
        canv.setFont("Helvetica", 10)
        canv.setFillColor(LIGHT_BLUE)
        canv.drawCentredString(page_w / 2, page_h - 34 * mm, f"Periodo:  {period}")

        canv.setFont("Helvetica", 8)
        canv.setFillColor(GRAY_TEXT)
        canv.drawRightString(page_w - 12 * mm, page_h - 15 * mm, f"Gerado em  {gen_date}")
        canv.drawRightString(page_w - 12 * mm, page_h - 26 * mm, f"{weeks} semana(s)  .  {n_equip} equipamentos")

        canv.setFillColor(NAVY)
        canv.rect(0, 0, page_w, footer_h, fill=1, stroke=0)
        canv.setFillColor(GOLD)
        canv.rect(0, footer_h, page_w, 2, fill=1, stroke=0)

        cy = (footer_h - 8) / 2
        canv.setFillColor(LIGHT_BLUE)
        canv.setFont("Helvetica", 8)
        canv.drawString(12 * mm, cy, "InsightFlow  .  Relatorios Operacionais")
        canv.drawRightString(page_w - 12 * mm, cy, f"Pagina {doc.page}")
        canv.setFillColor(GOLD)
        canv.setFont("Helvetica-Bold", 8)
        canv.drawCentredString(page_w / 2, cy, f"Horas: {_fmt_hours(total_h)}   |   KM: {int(total_km)}")
        canv.restoreState()

    template = PageTemplate(id="main", frames=[frame], onPage=on_page)
    doc = BaseDocTemplate(path, pagesize=A4, pageTemplates=[template])

    card_style = ParagraphStyle("card", fontName="Helvetica", fontSize=9, textColor=WHITE, alignment=TA_CENTER, leading=14)

    def make_card(label, value):
        return Paragraph(
            f'<font size="7" color="#CCCCCC">{label}</font><br/><font size="13"><b>{value}</b></font>',
            card_style,
        )

    cards_table = Table(
        [[
            make_card("TOTAL DE HORAS", _fmt_hours(total_h)),
            make_card("TOTAL DE KM", f"{int(total_km)} km"),
            make_card("SEMANAS ANALISADAS", str(weeks)),
            make_card("EQUIPAMENTOS", str(n_equip)),
        ]],
        colWidths=[frame_w / 4] * 4,
        rowHeights=[16 * mm],
    )
    cards_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), BLUE_CARD),
        ("BACKGROUND", (1, 0), (1, 0), GREEN_CARD),
        ("BACKGROUND", (2, 0), (2, 0), PURPLE_CARD),
        ("BACKGROUND", (3, 0), (3, 0), ORANGE_CARD),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEAFTER", (0, 0), (2, 0), 1.5, WHITE),
    ]))

    top_retro = _top3(machines, _is_retro)
    top_caminhao = _top3(vehicles, _is_caminhao)
    top_veiculo = _top3(vehicles, lambda n: _is_veiculo(n) or not _is_caminhao(n))
    retro_chart = _build_top3_category_chart(frame_w, "TOP 3 RETRO (HORAS)", top_retro, True, colors.HexColor("#1E88E5"))
    caminhao_chart = _build_top3_category_chart(frame_w, "TOP 3 CAMINHAO (KM)", top_caminhao, False, colors.HexColor("#43A047"))
    veiculo_chart = _build_top3_category_chart(frame_w, "TOP 3 VEICULO (KM)", top_veiculo, False, colors.HexColor("#FB8C00"))

    doc.build([
        cards_table,
        Spacer(1, 4 * mm),
        retro_chart,
        Spacer(1, 3 * mm),
        caminhao_chart,
        Spacer(1, 3 * mm),
        veiculo_chart,
    ])
    return path
