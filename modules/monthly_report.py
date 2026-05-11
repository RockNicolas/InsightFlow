
import os
import pandas as pd
import unicodedata
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame,
								Table, TableStyle, Paragraph, Spacer)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.legends import Legend
from modules.data_processor import get_weekly_data

# ── Paleta de cores ──────────────────────────────────────────────────────────
NAVY        = colors.HexColor('#1B3A5C')
GOLD        = colors.HexColor('#F4A720')
BLUE_CARD   = colors.HexColor('#1565C0')
GREEN_CARD  = colors.HexColor('#2E7D32')
PURPLE_CARD = colors.HexColor('#6A1B9A')
ORANGE_CARD = colors.HexColor('#BF360C')
SEC_MACH    = colors.HexColor('#0D3B6E')
SEC_VEH     = colors.HexColor('#1B5E20')
ROW_EVEN    = colors.HexColor('#EEF2F7')
TOTAL_BG    = colors.HexColor('#263238')
WHITE       = colors.white
GRAY_TEXT   = colors.HexColor('#78909C')
LIGHT_BLUE  = colors.HexColor('#B0BEC5')

LOCS_HORA      = ["LOC 01", "LOC 02", "LOC 05", "LOC 08"]
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
	text = unicodedata.normalize('NFKD', str(name).upper())
	return ''.join(c for c in text if not unicodedata.combining(c))


def _is_retro(name):
	n = _norm_name(name)
	return 'RETRO' in n


def _is_caminhao(name):
	n = _norm_name(name)
	return 'CAMINHAO' in n or 'TRUCK' in n


def _is_veiculo(name):
	n = _norm_name(name)
	keys = ['VEICULO', 'CARRO', 'PICKUP', 'VIATURA', 'UTILITARIO', 'SUV', 'MOTO']
	return any(k in n for k in keys)


def _top3(items, matcher):
	filtered = [i for i in items if matcher(i['display']) and i['hours'] > 0]
	return sorted(filtered, key=lambda x: x['hours'], reverse=True)[:3]


def _top3_text(title, rows, as_hours=False):
	if not rows:
		return f"<b>{title}</b><br/><font color='#666666'>Sem dados para este periodo</font>"

	lines = [f"<b>{title}</b>"]
	for idx, item in enumerate(rows, start=1):
		value = _fmt_hours(item['hours']) if as_hours else f"{int(item['hours'])} km"
		name = str(item['display'])[:28]
		lines.append(f"{idx}. {name} - <b>{value}</b>")
	return '<br/>'.join(lines)


def _fmt_rank_value(value, as_hours=False):
	return _fmt_hours(value) if as_hours else f"{int(value)} km"


def _build_top3_ranking_graph(width, top_retro, top_caminhao, top_veiculo):
	"""Mantido por compatibilidade; use _build_top3_category_chart para layout vertical."""
	return _build_top3_category_chart(width, 'TOP 3 RETRO (HORAS)', top_retro, True, colors.HexColor('#1E88E5'))


def _build_top3_category_chart(width, title, rows, as_hours, bar_color):
	"""Cria um grafico vertical grande para uma categoria do Top 3."""
	chart_h = 52 * mm
	d = Drawing(width, chart_h)
	d.add(Rect(0, 0, width, chart_h, fillColor=colors.HexColor('#F5F7FA'), strokeColor=colors.HexColor('#CFD8E3')))
	d.add(String(width / 2, chart_h - 11, title, fontName='Helvetica-Bold', fontSize=9, fillColor=NAVY, textAnchor='middle'))

	if not rows:
		d.add(String(width / 2, chart_h / 2, 'Sem dados para este periodo', fontName='Helvetica', fontSize=9, fillColor=GRAY_TEXT, textAnchor='middle'))
		return d

	labels = [f"{idx}. {str(item['display'])[:10]}" for idx, item in enumerate(rows, start=1)]
	values = [float(item['hours']) for item in rows]

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
	bar.categoryAxis.labels.boxAnchor = 'n'
	bar.bars[0].fillColor = bar_color
	bar.bars[0].strokeColor = bar_color
	bar.barSpacing = 8
	bar.groupSpacing = 12
	bar.barLabelFormat = '%0.0f'
	bar.barLabels.nudge = 6
	bar.barLabels.fontSize = 7
	bar.barLabels.fillColor = colors.HexColor('#37474F')
	bar.barLabels.boxAnchor = 's'
	bar.barLabels.visible = True
	d.add(bar)

	unit = 'horas' if as_hours else 'km'
	d.add(String(width - 8, chart_h - 11, unit, fontName='Helvetica', fontSize=7, fillColor=GRAY_TEXT, textAnchor='end'))

	return d


def _p(text, size=9, bold=False, color='#000000', align=TA_LEFT):
	"""Cria um Paragraph estilizado."""
	return Paragraph(text, ParagraphStyle(
		'auto',
		fontName='Helvetica-Bold' if bold else 'Helvetica',
		fontSize=size,
		textColor=colors.HexColor(color),
		alignment=align,
		leading=size * 1.5,
		leftPadding=2, rightPadding=2,
		spaceBefore=0, spaceAfter=0,
	))


def _chart_base_items(machines, vehicles):
	"""Seleciona a base dos graficos: horas de maquinas, ou KM se nao houver maquinas."""
	if machines:
		return machines, 'hours', 'Horas'
	return vehicles, 'hours', 'KM'


def _normalize_chart_data(items, key_name, limit=8):
	ordered = sorted(items, key=lambda x: x[key_name], reverse=True)
	if not ordered:
		return ["Sem dados"], [0]

	labels = [str(i['display'])[:12] for i in ordered[:limit]]
	values = [float(i[key_name]) for i in ordered[:limit]]

	if len(ordered) > limit:
		rest = sum(float(i[key_name]) for i in ordered[limit:])
		labels.append('Outros')
		values.append(rest)

	return labels, values


def _round_axis_max(max_value):
	if max_value <= 0:
		return 10
	steps = [10, 25, 50, 100, 250, 500, 1000]
	for step in steps:
		if max_value <= step:
			return step
	return int(((max_value // 1000) + 1) * 1000)


def _build_column_chart(machines, vehicles, width):
	items, metric, label = _chart_base_items(machines, vehicles)
	categories, values = _normalize_chart_data(items, metric, limit=10)

	chart_w = width
	chart_h = 68 * mm
	d = Drawing(chart_w, chart_h)
	d.add(Rect(0, 0, chart_w, chart_h, fillColor=colors.HexColor('#F4F6F8'), strokeColor=colors.HexColor('#D5DCE3')))
	d.add(String(chart_w / 2, chart_h - 12, 'Grafico - Colunas', fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))

	bar = VerticalBarChart()
	bar.x = 12 * mm
	bar.y = 12 * mm
	bar.height = 38 * mm
	bar.width = chart_w - 40 * mm
	bar.data = [values]
	bar.valueAxis.valueMin = 0
	bar.valueAxis.valueMax = _round_axis_max(max(values) * 1.1 if values else 0)
	bar.valueAxis.valueStep = max(1, int(bar.valueAxis.valueMax / 5))
	bar.categoryAxis.categoryNames = categories
	bar.categoryAxis.labels.angle = 70
	bar.categoryAxis.labels.boxAnchor = 'ne'
	bar.categoryAxis.labels.fontSize = 6
	bar.categoryAxis.labels.dy = -2
	bar.valueAxis.labels.fontSize = 7
	bar.bars[0].fillColor = colors.HexColor('#1C88C7')
	bar.bars[0].strokeColor = colors.HexColor('#0C5E97')
	bar.barSpacing = 3
	d.add(bar)

	legend = Legend()
	legend.x = chart_w - 24 * mm
	legend.y = chart_h - 22 * mm
	legend.alignment = 'right'
	legend.fontName = 'Helvetica'
	legend.fontSize = 8
	legend.colorNamePairs = [(colors.HexColor('#1C88C7'), label)]
	d.add(legend)

	return d


def _build_pie_chart(machines, vehicles, width):
	items, metric, _ = _chart_base_items(machines, vehicles)
	labels, values = _normalize_chart_data(items, metric, limit=7)

	chart_w = width
	chart_h = 68 * mm
	d = Drawing(chart_w, chart_h)
	d.add(Rect(0, 0, chart_w, chart_h, fillColor=colors.HexColor('#F4F6F8'), strokeColor=colors.HexColor('#D5DCE3')))
	d.add(String(chart_w / 2, chart_h - 12, 'Grafico - Pizza', fontName='Helvetica-Bold', fontSize=12, textAnchor='middle'))

	pie = Pie()
	pie.x = 12 * mm
	pie.y = 10 * mm
	pie.width = 42 * mm
	pie.height = 42 * mm
	pie.data = values
	pie.labels = ['' for _ in values]
	pie.slices.strokeColor = colors.white
	pie.slices.strokeWidth = 0.5

	palette = [
		colors.HexColor('#1C88C7'), colors.HexColor('#F25544'),
		colors.HexColor('#58A55C'), colors.HexColor('#9A6FB0'),
		colors.HexColor('#FF8A26'), colors.HexColor('#1CA7B8'),
		colors.HexColor('#7CB342'), colors.HexColor('#546E7A'),
	]
	for i in range(len(values)):
		pie.slices[i].fillColor = palette[i % len(palette)]

	d.add(pie)

	legend = Legend()
	legend.x = 65 * mm
	legend.y = 50 * mm
	legend.fontName = 'Helvetica'
	legend.fontSize = 7
	legend.dx = 8
	legend.dy = 8
	legend.deltay = 10
	legend.alignment = 'left'
	legend.colorNamePairs = [
		(palette[i % len(palette)], labels[i]) for i in range(len(labels))
	]
	d.add(legend)

	return d


# ── Ponto de entrada ─────────────────────────────────────────────────────────
def create_monthly_report(excel_path, obs_sheet_name, output_folder, selected_sheets=None, weekly_data_loader=None):
	"""Gera um PDF mensal usando apenas as abas selecionadas."""
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
			if (
				sheet_name
				and sheet_name in available_sheets
				and sheet_name != obs_sheet_name
				and sheet_name not in valid_sheets
			):
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
			key = " ".join(item['machine'].upper().split())
			if key not in machine_totals:
				machine_totals[key] = {
					'display': item['machine'],
					'hours': 0.0,
					'operator': item['operator'],
				}
			machine_totals[key]['hours'] += item['hours']

	if not machine_totals:
		print("Sem dados.")
		return None

	aggregated = list(machine_totals.values())
	machines   = [i for i in aggregated if _is_time(i['display'])]
	vehicles   = [i for i in aggregated if not _is_time(i['display'])]
	total_h    = sum(i['hours'] for i in machines)
	total_km   = sum(i['hours'] for i in vehicles)
	weeks      = len(valid_sheets)
	n_equip    = len(aggregated)
	period     = f"{valid_sheets[0]}  ->  {valid_sheets[-1]}" if weeks > 1 else valid_sheets[0]
	gen_date   = datetime.now().strftime("%d/%m/%Y  %H:%M")

	return _build_pdf(machines, vehicles, total_h, total_km,
					  weeks, n_equip, period, gen_date,
					  output_folder, valid_sheets)


# ── Geracao do PDF ───────────────────────────────────────────────────────────
def _build_pdf(machines, vehicles, total_h, total_km,
			   weeks, n_equip, period, gen_date,
			   output_folder, valid_sheets):

	if not os.path.exists(output_folder):
		os.makedirs(output_folder)

	logo_path = os.path.join("assets", "company", "company_2.png")
	safe      = valid_sheets[-1].replace(' ', '_').replace('/', '-')
	path      = os.path.join(output_folder, f"Report_MENSAL_{safe}.pdf")

	W, H     = A4
	HEADER_H = 55 * mm
	FOOTER_H = 14 * mm
	MARGIN   = 12 * mm
	DECO     = 3
	FW       = W - 2 * MARGIN

	frame = Frame(
		MARGIN, FOOTER_H + 4 * mm,
		FW,
		H - HEADER_H - DECO - FOOTER_H - 12 * mm,
		id='main', leftPadding=0, rightPadding=0,
		topPadding=5 * mm, bottomPadding=0,
	)

	def on_page(canv, doc):
		canv.saveState()

		# Header navy
		canv.setFillColor(NAVY)
		canv.rect(0, H - HEADER_H, W, HEADER_H, fill=1, stroke=0)

		# Linha dourada abaixo do header
		canv.setFillColor(GOLD)
		canv.rect(0, H - HEADER_H - DECO, W, DECO, fill=1, stroke=0)

		# Logo
		if os.path.exists(logo_path):
			canv.drawImage(logo_path, 12 * mm, H - HEADER_H + 8 * mm,
						   width=35 * mm, height=35 * mm,
						   preserveAspectRatio=True, mask='auto')

		# Titulo
		canv.setFillColor(WHITE)
		canv.setFont("Helvetica-Bold", 20)
		canv.drawCentredString(W / 2, H - 22 * mm, "RELATORIO MENSAL DE PRODUCAO")
		canv.setFont("Helvetica", 10)
		canv.setFillColor(LIGHT_BLUE)
		canv.drawCentredString(W / 2, H - 34 * mm, f"Periodo:  {period}")

		# Meta (canto superior direito)
		canv.setFont("Helvetica", 8)
		canv.setFillColor(GRAY_TEXT)
		canv.drawRightString(W - 12 * mm, H - 15 * mm, f"Gerado em  {gen_date}")
		canv.drawRightString(W - 12 * mm, H - 26 * mm,
							 f"{weeks} semana(s)  .  {n_equip} equipamentos")

		# Footer navy
		canv.setFillColor(NAVY)
		canv.rect(0, 0, W, FOOTER_H, fill=1, stroke=0)
		canv.setFillColor(GOLD)
		canv.rect(0, FOOTER_H, W, 2, fill=1, stroke=0)

		cy = (FOOTER_H - 8) / 2
		canv.setFillColor(LIGHT_BLUE)
		canv.setFont("Helvetica", 8)
		canv.drawString(12 * mm, cy, "InsightFlow  .  Relatorios Operacionais")
		canv.drawRightString(W - 12 * mm, cy, f"Pagina {doc.page}")

		summary = f"Horas: {_fmt_hours(total_h)}   |   KM: {int(total_km)}"
		canv.setFillColor(GOLD)
		canv.setFont("Helvetica-Bold", 8)
		canv.drawCentredString(W / 2, cy, summary)

		canv.restoreState()

	template = PageTemplate(id='main', frames=[frame], onPage=on_page)
	doc = BaseDocTemplate(path, pagesize=A4, pageTemplates=[template])

	# ── Cards de resumo ──────────────────────────────────────────────────────
	CW = FW / 4
	card_style = ParagraphStyle('card', fontName='Helvetica', fontSize=9,
								textColor=WHITE, alignment=TA_CENTER, leading=14)

	def make_card(label, value):
		return Paragraph(
			f'<font size="7" color="#CCCCCC">{label}</font><br/>'
			f'<font size="13"><b>{value}</b></font>',
			card_style,
		)

	cards_table = Table(
		[[make_card("TOTAL DE HORAS", _fmt_hours(total_h)),
		  make_card("TOTAL DE KM", f"{int(total_km)} km"),
		  make_card("SEMANAS ANALISADAS", str(weeks)),
		  make_card("EQUIPAMENTOS", str(n_equip))]],
		colWidths=[CW] * 4,
		rowHeights=[16 * mm],
	)
	cards_table.setStyle(TableStyle([
		('BACKGROUND', (0, 0), (0, 0), BLUE_CARD),
		('BACKGROUND', (1, 0), (1, 0), GREEN_CARD),
		('BACKGROUND', (2, 0), (2, 0), PURPLE_CARD),
		('BACKGROUND', (3, 0), (3, 0), ORANGE_CARD),
		('ALIGN',      (0, 0), (-1, -1), 'CENTER'),
		('VALIGN',     (0, 0), (-1, -1), 'MIDDLE'),
		('LINEAFTER',  (0, 0), (2, 0), 1.5, WHITE),
	]))

	# ── Tabela principal ─────────────────────────────────────────────────────
	COL = [FW * 0.38, FW * 0.38, FW * 0.24]
	R_H = 9  * mm
	S_H = 7  * mm
	H_H = 8  * mm

	data = []
	cmds = []
	rh   = []
	r    = 0

	# Cabecalho das colunas
	data.append([
		_p('   MAQUINA / PLACA', 9, True, '#FFFFFF'),
		_p('OPERADOR',           9, True, '#FFFFFF', TA_CENTER),
		_p('TOTAL MENSAL',       9, True, '#FFFFFF', TA_RIGHT),
	])
	cmds += [
		('BACKGROUND',    (0, r), (-1, r), NAVY),
		('VALIGN',        (0, r), (-1, r), 'MIDDLE'),
		('TOPPADDING',    (0, r), (-1, r), 5),
		('BOTTOMPADDING', (0, r), (-1, r), 5),
		('LINEBELOW',     (0, r), (-1, r), 1.5, GOLD),
	]
	rh.append(H_H); r += 1

	# Secao: MAQUINAS
	data.append([_p('  >>  MAQUINAS E EQUIPAMENTOS', 8, True, '#FFFFFF'), '', ''])
	cmds += [
		('BACKGROUND',    (0, r), (-1, r), SEC_MACH),
		('SPAN',          (0, r), (-1, r)),
		('VALIGN',        (0, r), (-1, r), 'MIDDLE'),
		('TOPPADDING',    (0, r), (-1, r), 4),
		('BOTTOMPADDING', (0, r), (-1, r), 4),
	]
	rh.append(S_H); r += 1

	for idx, item in enumerate(machines):
		nm, hrs, op = item['display'], item['hours'], item['operator']
		alert = _is_alert(nm, hrs)
		fc    = '#C62828' if alert else '#1A237E'
		data.append([
			_p(f'   {str(nm)[:43]}', 9, alert, fc),
			_p(str(op)[:35],         9, False, '#333333', TA_CENTER),
			_p(_fmt_hours(hrs),      9, True,  fc, TA_RIGHT),
		])
		bg = ROW_EVEN if idx % 2 == 0 else WHITE
		cmds += [
			('BACKGROUND',    (0, r), (-1, r), bg),
			('VALIGN',        (0, r), (-1, r), 'MIDDLE'),
			('TOPPADDING',    (0, r), (-1, r), 3),
			('BOTTOMPADDING', (0, r), (-1, r), 3),
			('LINEBELOW',     (0, r), (-1, r), 0.3, colors.HexColor('#DDDDDD')),
		]
		rh.append(R_H); r += 1

	# Secao: VEICULOS
	data.append([_p('  >>  VEICULOS E APOIO', 8, True, '#FFFFFF'), '', ''])
	cmds += [
		('BACKGROUND',    (0, r), (-1, r), SEC_VEH),
		('SPAN',          (0, r), (-1, r)),
		('VALIGN',        (0, r), (-1, r), 'MIDDLE'),
		('TOPPADDING',    (0, r), (-1, r), 4),
		('BOTTOMPADDING', (0, r), (-1, r), 4),
	]
	rh.append(S_H); r += 1

	for idx, item in enumerate(vehicles):
		nm, hrs, op = item['display'], item['hours'], item['operator']
		alert = _is_alert(nm, hrs)
		fc    = '#C62828' if alert else '#1B5E20'
		data.append([
			_p(f'   {str(nm)[:43]}', 9, alert, fc),
			_p(str(op)[:35],         9, False, '#333333', TA_CENTER),
			_p(f'{int(hrs)} KM',     9, True,  fc, TA_RIGHT),
		])
		bg = ROW_EVEN if idx % 2 == 0 else WHITE
		cmds += [
			('BACKGROUND',    (0, r), (-1, r), bg),
			('VALIGN',        (0, r), (-1, r), 'MIDDLE'),
			('TOPPADDING',    (0, r), (-1, r), 3),
			('BOTTOMPADDING', (0, r), (-1, r), 3),
			('LINEBELOW',     (0, r), (-1, r), 0.3, colors.HexColor('#DDDDDD')),
		]
		rh.append(R_H); r += 1

	# Total de horas
	data.append([
		_p('   >>  TOTAL GERAL DE HORAS (MAQUINAS)', 9, True, '#FFFFFF'),
		'',
		_p(_fmt_hours(total_h), 10, True, '#F4A720', TA_RIGHT),
	])
	cmds += [
		('BACKGROUND',    (0, r), (-1, r), TOTAL_BG),
		('SPAN',          (0, r), (1, r)),
		('VALIGN',        (0, r), (-1, r), 'MIDDLE'),
		('TOPPADDING',    (0, r), (-1, r), 5),
		('BOTTOMPADDING', (0, r), (-1, r), 5),
		('LINEABOVE',     (0, r), (-1, r), 1.5, GOLD),
	]
	rh.append(H_H); r += 1

	# Total de km
	data.append([
		_p('   >>  TOTAL GERAL DE KM (VEICULOS E APOIO)', 9, True, '#FFFFFF'),
		'',
		_p(f"{int(total_km)} KM", 10, True, '#69F0AE', TA_RIGHT),
	])
	cmds += [
		('BACKGROUND',    (0, r), (-1, r), TOTAL_BG),
		('SPAN',          (0, r), (1, r)),
		('VALIGN',        (0, r), (-1, r), 'MIDDLE'),
		('TOPPADDING',    (0, r), (-1, r), 5),
		('BOTTOMPADDING', (0, r), (-1, r), 5),
		('LINEBELOW',     (0, r), (-1, r), 1, NAVY),
	]
	rh.append(H_H); r += 1

	# Borda externa
	cmds += [
		('BOX',  (0, 0), (-1, -1), 1.5, NAVY),
		('GRID', (0, 0), (-1, -1), 0, colors.transparent),
	]

	main_table = Table(data, colWidths=COL, rowHeights=rh)
	main_table.setStyle(TableStyle(cmds))

	top_retro = _top3(machines, _is_retro)
	top_caminhao = _top3(vehicles, _is_caminhao)
	top_veiculo = _top3(vehicles, lambda n: _is_veiculo(n) or not _is_caminhao(n))
	retro_chart = _build_top3_category_chart(FW, 'TOP 3 RETRO (HORAS)', top_retro, True, colors.HexColor('#1E88E5'))
	caminhao_chart = _build_top3_category_chart(FW, 'TOP 3 CAMINHAO (KM)', top_caminhao, False, colors.HexColor('#43A047'))
	veiculo_chart = _build_top3_category_chart(FW, 'TOP 3 VEICULO (KM)', top_veiculo, False, colors.HexColor('#FB8C00'))

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

