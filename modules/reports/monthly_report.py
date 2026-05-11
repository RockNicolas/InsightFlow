import os
import re
import unicodedata
from datetime import datetime

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from modules.processors.data_processor import get_weekly_data

NAVY = colors.HexColor("#17324D")
BLUE = colors.HexColor("#2563EB")
GREEN = colors.HexColor("#2E7D32")
PURPLE = colors.HexColor("#6A1B9A")
ORANGE = colors.HexColor("#BF360C")
RED = colors.HexColor("#C62828")
LIGHT_BG = colors.HexColor("#F5F8FC")
ROW_ALT = colors.HexColor("#F8FAFC")
BORDER = colors.HexColor("#D7E0EA")
TEXT = colors.HexColor("#162334")
MUTED = colors.HexColor("#607285")
WHITE = colors.white
MONTH_LABELS = {
    "01": "JANEIRO",
    "02": "FEVEREIRO",
    "03": "MARCO",
    "04": "ABRIL",
    "05": "MAIO",
    "06": "JUNHO",
    "07": "JULHO",
    "08": "AGOSTO",
    "09": "SETEMBRO",
    "10": "OUTUBRO",
    "11": "NOVEMBRO",
    "12": "DEZEMBRO",
}

LOCS_HORA = ["LOC 01", "LOC 02", "LOC 05", "LOC 08"]
LISTA_VERMELHA = ["MC 01", "MC 13"] + LOCS_HORA
TRUCK_KEYWORDS = ["CAMINHAO", "CACAMBA", "PRANCHA", "VOLVO"]

_base_styles = getSampleStyleSheet()
STYLE_SECTION = ParagraphStyle(
    "section",
    parent=_base_styles["Heading2"],
    fontName="Helvetica-Bold",
    fontSize=14,
    leading=18,
    textColor=NAVY,
    spaceAfter=6,
)
STYLE_BODY = ParagraphStyle(
    "body",
    parent=_base_styles["BodyText"],
    fontName="Helvetica",
    fontSize=9.5,
    leading=14,
    textColor=TEXT,
)
STYLE_MUTED = ParagraphStyle(
    "muted",
    parent=STYLE_BODY,
    textColor=MUTED,
)
STYLE_CARD = ParagraphStyle(
    "card",
    parent=STYLE_BODY,
    alignment=TA_CENTER,
    textColor=WHITE,
    leading=16,
)
STYLE_TABLE_HEADER = ParagraphStyle(
    "table_header",
    parent=STYLE_BODY,
    fontName="Helvetica-Bold",
    fontSize=8.5,
    textColor=WHITE,
    alignment=TA_CENTER,
)
STYLE_TABLE_CELL = ParagraphStyle(
    "table_cell",
    parent=STYLE_BODY,
    fontSize=8.5,
    leading=11.5,
)


def _fmt_hours(value):
    hours = int(value)
    minutes = int(round((value - hours) * 60))
    return f"{hours}:{minutes:02d}h"


def _fmt_metric(value, as_hours):
    return _fmt_hours(value) if as_hours else f"{int(round(value))} km"


def _norm_name(name):
    text = unicodedata.normalize("NFKD", str(name).upper())
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _is_machine(name):
    normalized = _norm_name(name)
    return "MC" in normalized or any(loc in normalized for loc in LOCS_HORA)


def _is_truck(name):
    normalized = _norm_name(name)
    return any(keyword in normalized for keyword in TRUCK_KEYWORDS)


def _is_alert(name, hours):
    normalized = _norm_name(name)
    return any(item in normalized for item in LISTA_VERMELHA) or hours == 0


def _truncate(value, limit):
    text = str(value or "").strip()
    return text if len(text) <= limit else f"{text[:limit - 1]}..."


def _top_items(items, limit=5):
    return sorted(items, key=lambda item: item["hours"], reverse=True)[:limit]


def _rank_label(name):
    text = str(name or "").strip()
    normalized = _norm_name(text)

    if "RETROESCAVADEIRA" in normalized and "MC" in normalized:
        match = re.search(r"\b(MC\s*\d{1,2})\b", normalized)
        if match:
            return f"{match.group(1)} - RETROESCAVADEIRA"

    return text


def _extract_month_label(sheet_names):
    for sheet_name in reversed(sheet_names):
        match = re.search(r"(?:\.|/)(\d{2})(?:\D|$)", str(sheet_name or ""))
        if match:
            return MONTH_LABELS.get(match.group(1), match.group(1))

    current_month = datetime.now().strftime("%m")
    return MONTH_LABELS.get(current_month, current_month)


def _top_summary(items, as_hours, empty_label="Sem dados"):
    top_item = _top_items(items, 1)
    if not top_item:
        return f"{empty_label} - -"

    item = top_item[0]
    return f"{_truncate(_rank_label(item['display']), 32)} - {_fmt_metric(item['hours'], as_hours)}"


def _color_hex(color):
    return "#{:02X}{:02X}{:02X}".format(
        int(color.red * 255),
        int(color.green * 255),
        int(color.blue * 255),
    )


def _draw_page_chrome(canvas, doc, period, gen_date, sheet_count):
    canvas.saveState()
    page_w, page_h = A4
    header_h = 24 * mm
    footer_h = 11 * mm
    logo_path = os.path.join("assets", "company", "company_2.png")

    canvas.setFillColor(NAVY)
    canvas.rect(0, page_h - header_h, page_w, header_h, fill=1, stroke=0)

    if os.path.exists(logo_path):
        canvas.drawImage(
            logo_path,
            12 * mm,
            page_h - 18 * mm,
            width=22 * mm,
            height=10 * mm,
            preserveAspectRatio=True,
            mask="auto",
        )

    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica-Bold", 18)
    canvas.drawString(38 * mm, page_h - 12 * mm, "RELATÓRIO MENSAL DE PRODUÇÃO")
    canvas.setFont("Helvetica", 9)
    canvas.drawString(38 * mm, page_h - 18 * mm, f"Período: {period}")

    canvas.setFont("Helvetica", 8)
    canvas.drawRightString(page_w - 12 * mm, page_h - 12 * mm, f"Gerado em {gen_date}")
    canvas.drawRightString(page_w - 12 * mm, page_h - 18 * mm, f"{sheet_count} semana(s) selecionada(s)")

    canvas.setFillColor(NAVY)
    canvas.rect(0, 0, page_w, footer_h, fill=1, stroke=0)
    canvas.setFillColor(colors.HexColor("#D6E2EE"))
    canvas.setFont("Helvetica", 8)
    canvas.drawString(12 * mm, 4, "InsightFlow  |  Resumo operacional mensal")
    canvas.drawRightString(page_w - 12 * mm, 4, f"Página {doc.page}")
    canvas.restoreState()


def _build_kpi_cards(doc_width, cards):
    rows = []
    styles = []
    idx = 0
    for row_idx in range(2):
        row = []
        for col_idx in range(2):
            label, value, bg = cards[idx]
            row.append(Paragraph(
                f'<font size="8" color="#DDE7F1">{label}</font><br/><font size="16"><b>{value}</b></font>',
                STYLE_CARD,
            ))
            styles.append(("BACKGROUND", (col_idx, row_idx), (col_idx, row_idx), bg))
            idx += 1
        rows.append(row)

    table = Table(rows, colWidths=[doc_width / 2] * 2, rowHeights=[18 * mm, 18 * mm])
    table.setStyle(TableStyle(
        styles + [
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("BOX", (0, 0), (-1, -1), 0.5, WHITE),
            ("INNERGRID", (0, 0), (-1, -1), 2, WHITE),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ]
    ))
    return table


def _build_simple_table(headers, rows, col_widths, accent_color):
    data = [[Paragraph(header, STYLE_TABLE_HEADER) for header in headers]]
    data.extend(rows)

    table = Table(data, colWidths=col_widths, repeatRows=1)
    styles = [
        ("BACKGROUND", (0, 0), (-1, 0), accent_color),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 1, BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ]

    for idx in range(1, len(data)):
        bg_color = ROW_ALT if idx % 2 == 0 else WHITE
        styles.append(("BACKGROUND", (0, idx), (-1, idx), bg_color))

    table.setStyle(TableStyle(styles))
    return table


def _build_weeks_table(doc_width, sheet_stats):
    rows = []
    for stat in sheet_stats:
        status = "Com dados" if stat["rows_with_data"] > 0 else "Sem dados"
        status_color = GREEN if stat["rows_with_data"] > 0 else RED
        rows.append([
            Paragraph(_truncate(stat["sheet"], 24), STYLE_TABLE_CELL),
            Paragraph(str(stat["rows_read"]), STYLE_TABLE_CELL),
            Paragraph(str(stat["rows_with_data"]), STYLE_TABLE_CELL),
            Paragraph(f'<font color="{_color_hex(status_color)}"><b>{status}</b></font>', STYLE_TABLE_CELL),
        ])

    return _build_simple_table(
        ["Semana", "Linhas lidas", "Com dados", "Status"],
        rows,
        [doc_width * 0.42, doc_width * 0.16, doc_width * 0.16, doc_width * 0.26],
        NAVY,
    )


def _build_rank_table(title, items, as_hours, accent_color, width, name_limit=28):
    top_items = _top_items(items, limit=5)
    rows = []
    for idx, item in enumerate(top_items, start=1):
        rows.append([
            Paragraph(str(idx), STYLE_TABLE_CELL),
            Paragraph(_truncate(_rank_label(item["display"]), name_limit), STYLE_TABLE_CELL),
            Paragraph(_fmt_metric(item["hours"], as_hours), STYLE_TABLE_CELL),
        ])

    if not rows:
        rows.append([
            Paragraph("-", STYLE_TABLE_CELL),
            Paragraph("Sem dados", STYLE_MUTED),
            Paragraph("-", STYLE_TABLE_CELL),
        ])

    title_row = [Paragraph(title, ParagraphStyle(
        f"title_{title}",
        parent=STYLE_BODY,
        fontName="Helvetica-Bold",
        fontSize=10,
        textColor=WHITE,
        alignment=TA_LEFT,
    ))]
    title_table = Table([title_row], colWidths=[width])
    title_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), accent_color),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))

    body = _build_simple_table(
        ["#", "Equipamento", "Total"],
        rows,
        [width * 0.10, width * 0.60, width * 0.30],
        colors.HexColor("#8CA3BA"),
    )
    return [title_table, body]


def _build_rank_grid(doc_width, machines, trucks, vehicles):
    rank_width = (doc_width - 12) / 3
    machine_rank = _build_rank_table("Top 5 máquinas", machines, True, BLUE, rank_width, name_limit=18)
    truck_rank = _build_rank_table("Top 5 caminhões", trucks, False, ORANGE, rank_width, name_limit=18)
    vehicle_rank = _build_rank_table("Top 5 veículos", vehicles, False, GREEN, rank_width, name_limit=18)

    grid = Table(
        [[[machine_rank[0], machine_rank[1]], [truck_rank[0], truck_rank[1]], [vehicle_rank[0], vehicle_rank[1]]]],
        colWidths=[rank_width, rank_width, rank_width],
    )
    grid.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return grid


def _build_highlights_table(doc_width, machines, trucks, vehicles):
    rows = [
        [
            Paragraph("Máquina destaque", STYLE_TABLE_CELL),
            Paragraph(_top_summary(machines, True), STYLE_TABLE_CELL),
        ],
        [
            Paragraph("Caminhão destaque", STYLE_TABLE_CELL),
            Paragraph(_top_summary(trucks, False), STYLE_TABLE_CELL),
        ],
        [
            Paragraph("Veículo destaque", STYLE_TABLE_CELL),
            Paragraph(_top_summary(vehicles, False), STYLE_TABLE_CELL),
        ],
    ]

    return _build_simple_table(
        ["Categoria", "Resumo"],
        rows,
        [doc_width * 0.28, doc_width * 0.72],
        BLUE,
    )


def _build_detail_table(title, subtitle, items, weeks, as_hours, accent_color):
    story = [Paragraph(title, STYLE_SECTION), Paragraph(subtitle, STYLE_MUTED), Spacer(1, 4 * mm)]

    active_items = [item for item in items if item["hours"] > 0]
    if not active_items:
        empty = Table([[Paragraph("Sem dados disponíveis para esta categoria.", STYLE_MUTED)]], colWidths=[175 * mm])
        empty.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), LIGHT_BG),
            ("BOX", (0, 0), (-1, -1), 1, BORDER),
            ("TOPPADDING", (0, 0), (-1, -1), 12),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(empty)
        return story

    rows = []
    sorted_items = sorted(active_items, key=lambda item: item["hours"], reverse=True)
    for idx, item in enumerate(sorted_items, start=1):
        avg_value = item["hours"] / max(weeks, 1)
        status = "Alerta" if _is_alert(item["display"], item["hours"]) else "OK"
        status_color = RED if status == "Alerta" else GREEN
        rows.append([
            Paragraph(str(idx), STYLE_TABLE_CELL),
            Paragraph(_truncate(item["display"], 34), STYLE_TABLE_CELL),
            Paragraph(_truncate(item.get("operator") or "N/A", 18), STYLE_TABLE_CELL),
            Paragraph(_fmt_metric(item["hours"], as_hours), STYLE_TABLE_CELL),
            Paragraph(_fmt_metric(avg_value, as_hours), STYLE_TABLE_CELL),
            Paragraph(f"{item.get('weeks_active', 0)}/{weeks}", STYLE_TABLE_CELL),
            Paragraph(f'<font color="{_color_hex(status_color)}"><b>{status}</b></font>', STYLE_TABLE_CELL),
        ])

    table = _build_simple_table(
        ["#", "Equipamento", "Operador", "Total", "Média/semana", "Semanas", "Status"],
        rows,
        [8 * mm, 58 * mm, 32 * mm, 22 * mm, 24 * mm, 18 * mm, 20 * mm],
        accent_color,
    )
    story.append(table)
    return story


def create_monthly_report(excel_path, obs_sheet_name, output_folder, selected_sheets=None, weekly_data_loader=None):
    try:
        xl = pd.ExcelFile(excel_path)
        available_sheets = xl.sheet_names
    except Exception as exc:
        print(f"ERRO ao abrir Excel: {exc}")
        return None

    available_lookup = {}
    for sheet_name in available_sheets:
        normalized = str(sheet_name or "").strip()
        if normalized and normalized not in available_lookup:
            available_lookup[normalized] = sheet_name

    obs_sheet_name = available_lookup.get(str(obs_sheet_name or "").strip(), obs_sheet_name)

    if selected_sheets:
        valid_sheets = []
        for sheet in selected_sheets:
            normalized = str(sheet or "").strip()
            actual_sheet_name = available_lookup.get(normalized)
            if (
                actual_sheet_name
                and actual_sheet_name != obs_sheet_name
                and actual_sheet_name not in valid_sheets
            ):
                valid_sheets.append(actual_sheet_name)
    else:
        valid_sheets = [sheet for sheet in available_sheets if sheet != obs_sheet_name]

    if not valid_sheets:
        print("ERRO: Nenhuma aba semanal encontrada.")
        return None

    print(f"[*] Abas semanais candidatas: {valid_sheets}")

    weekly_data_loader = weekly_data_loader or get_weekly_data
    machine_totals = {}
    sheet_stats = []

    for sheet in valid_sheets:
        print(f"    -> Lendo: {sheet}")
        sheet_items = weekly_data_loader(excel_path, sheet)
        valid_items = [item for item in sheet_items if float(item.get("hours", 0) or 0) > 0]
        sheet_stats.append(
            {
                "sheet": sheet,
                "rows_read": len(sheet_items),
                "rows_with_data": len(valid_items),
            }
        )
        print(f"       linhas lidas: {len(sheet_items)} | com dados válidos: {len(valid_items)}")

        for item in sheet_items:
            key = " ".join(item["machine"].upper().split())
            if key not in machine_totals:
                machine_totals[key] = {
                    "display": item["machine"],
                    "hours": 0.0,
                    "operator": item.get("operator") or "N/A",
                    "weeks_active": 0,
                }

            machine_totals[key]["hours"] += item["hours"]
            if item["hours"] > 0:
                machine_totals[key]["weeks_active"] += 1
            if item.get("operator") and item.get("operator") != "N/A":
                machine_totals[key]["operator"] = item["operator"]

    populated_sheets = [item["sheet"] for item in sheet_stats if item["rows_with_data"] > 0]

    if len(populated_sheets) == 1:
        print(f"[!] Apenas uma aba teve dados válidos no mensal: {populated_sheets[0]}")
    elif len(populated_sheets) > 1:
        print(f"[*] Abas com dados válidos no mensal: {', '.join(populated_sheets)}")
    else:
        print("[!] Nenhuma aba teve dados válidos para compor o mensal.")

    if not machine_totals or not populated_sheets:
        print("Sem dados.")
        return None

    aggregated = list(machine_totals.values())
    machines = [item for item in aggregated if _is_machine(item["display"])]
    trucks = [item for item in aggregated if not _is_machine(item["display"]) and _is_truck(item["display"])]
    vehicles = [item for item in aggregated if not _is_machine(item["display"]) and not _is_truck(item["display"])]
    active_items = [item for item in aggregated if item["hours"] > 0]
    total_h = sum(item["hours"] for item in machines)
    total_km = sum(item["hours"] for item in trucks + vehicles)
    weeks = len(valid_sheets)
    period = f"{valid_sheets[0]}  ->  {valid_sheets[-1]}" if weeks > 1 else valid_sheets[0]
    gen_date = datetime.now().strftime("%d/%m/%Y  %H:%M")

    monthly_output_folder = os.path.join(output_folder, "mensal")
    if not os.path.exists(monthly_output_folder):
        os.makedirs(monthly_output_folder)

    month_label = _extract_month_label(valid_sheets)
    pdf_path = os.path.join(monthly_output_folder, f"Relatorio_Mensal_{month_label}.pdf")
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        topMargin=32 * mm,
        bottomMargin=16 * mm,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
    )

    kpi_cards = [
        ("Horas de máquinas", _fmt_hours(total_h), NAVY),
        ("KM rodados", f"{int(total_km)} km", GREEN),
        ("Semanas com dados", str(len(populated_sheets)), PURPLE),
        ("Equipamentos ativos", str(len(active_items)), ORANGE),
    ]

    story = [
        Paragraph("Resumo do período", STYLE_SECTION),
        Paragraph(
            "Resumo mensal consolidado com foco nas semanas que realmente tiveram dados aproveitados pelo sistema.",
            STYLE_BODY,
        ),
        Spacer(1, 5 * mm),
        _build_kpi_cards(doc.width, kpi_cards),
        Spacer(1, 5 * mm),
        Paragraph("Semanas consideradas", STYLE_SECTION),
        Paragraph("A tabela abaixo mostra exatamente quais abas foram lidas e quais delas contribuíram com dados válidos para o mensal.", STYLE_MUTED),
        Spacer(1, 3 * mm),
        _build_weeks_table(doc.width, sheet_stats),
        Spacer(1, 5 * mm),
        Paragraph("Destaques principais", STYLE_SECTION),
        Spacer(1, 2 * mm),
        _build_highlights_table(doc.width, machines, trucks, vehicles),
        Spacer(1, 5 * mm),
        _build_rank_grid(doc.width, machines, trucks, vehicles),
    ]

    story.append(PageBreak())
    story.extend(_build_detail_table(
        "Máquinas e Equipamentos",
        "Detalhamento mensal das máquinas com dados válidos.",
        machines,
        weeks,
        True,
        NAVY,
    ))

    story.append(PageBreak())
    story.extend(_build_detail_table(
        "Veículos e Apoio",
        "Detalhamento mensal dos caminhões com dados válidos.",
        trucks,
        weeks,
        False,
        ORANGE,
    ))

    story.append(PageBreak())
    story.extend(_build_detail_table(
        "Veículos Leves e Apoio",
        "Detalhamento mensal dos veículos leves com dados válidos.",
        vehicles,
        weeks,
        False,
        GREEN,
    ))

    def on_page(canvas, document):
        _draw_page_chrome(canvas, document, period, gen_date, weeks)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return pdf_path
