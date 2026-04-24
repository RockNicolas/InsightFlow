from fpdf import FPDF
import os

def format_value(machine_name, value, locs_permitidas):
    """
    Regra refinada: 
    - Se contiver 'MC' -> Horas.
    - Se for uma das LOCs permitidas -> Horas.
    - Restante -> KM.
    """
    m_upper = machine_name.upper()
    
    is_loc_hora = any(loc in m_upper for loc in locs_permitidas)
    is_mc = "MC" in m_upper

    if is_mc or is_loc_hora:
        hours = int(value)
        minutes = int(round((value - hours) * 60))
        return f"{hours}:{minutes:02d}h"
    else:
        return f"{int(value)} KM"

def format_final_meter(machine_name, value, locs_permitidas):
    """Formata leitura final de horimetro/km."""
    m_upper = machine_name.upper()

    if value is None:
        return "-"

    # Retroescavadeiras exibem sem ponto de milhar.
    if "RETROESCAVADEIRA" in m_upper or "RETRO" in m_upper:
        return f"{value:.1f}".replace(".", ",")

    is_loc_hora = any(loc in m_upper for loc in locs_permitidas)
    is_mc = "MC" in m_upper

    if is_mc or is_loc_hora:
        return f"{value:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{int(value):,}".replace(",", ".")

def create_pdf_report(data, sheet_name, output_path):
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    pdf = FPDF()
    pdf.add_page()

    logo_path = os.path.join("assets", "company", "company_2.png")
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=85, y=3, w=40)
        pdf.ln(30)
    else:
        pdf.ln(10)

    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 10, f"RELATÓRIO SEMANAL - {sheet_name}", ln=True, align='C')
    pdf.ln(5)

    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)
    pdf.cell(70, 10, "MÁQUINA / PLACA", 1, 0, 'C', True)
    pdf.cell(50, 10, "OPERADORES", 1, 0, 'C', True)
    pdf.cell(35, 10, "PRODUÇÃO", 1, 0, 'C', True)
    pdf.cell(35, 10, "HORI SEMANAL", 1, 1, 'C', True)
    
    pdf.set_font("Arial", "", 9)

    total_horas_decimal = 0
    total_km_acumulado = 0
    
    locs_hora = ["LOC 01", "LOC 02", "LOC 05", "LOC 08"]
    lista_vermelha = ["MC 01", "MC 13"] + locs_hora

    for item in data:
        m_upper = item['machine'].upper()
        
        is_machine_time = "MC" in m_upper or any(loc in m_upper for loc in locs_hora)
        
        if is_machine_time:
            total_horas_decimal += item['hours']
        else:
            total_km_acumulado += item['hours']

        is_red = any(x in m_upper for x in lista_vermelha) or item['hours'] == 0
        if is_red:
            pdf.set_text_color(255, 0, 0)
            pdf.set_font("Arial", "B", 9)
        else:
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", "", 9)

        valor_formatado = format_value(item['machine'], item['hours'], locs_hora)
        final_formatado = format_final_meter(item['machine'], item.get('final_meter'), locs_hora)

        pdf.cell(70, 10, str(item['machine'])[:37], 1)
        pdf.cell(50, 10, str(item['operator'])[:26], 1)
        pdf.cell(35, 10, valor_formatado, 1, 0, 'R')
        pdf.cell(35, 10, final_formatado, 1, 1, 'R')

    pdf.ln(2)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(240, 240, 240)

    total_h_string = format_value("TOTAL HORAS", total_horas_decimal, ["TOTAL"])
    pdf.cell(155, 10, "TOTAL GERAL DE HORAS (MÁQUINAS)", 1, 0, 'L', True)
    pdf.cell(40, 10, total_h_string, 1, 1, 'R', True)

    pdf.cell(155, 10, "TOTAL GERAL DE KM (VEÍCULOS E APOIO)", 1, 0, 'L', True)
    pdf.cell(40, 10, f"{int(total_km_acumulado)} KM", 1, 1, 'R', True)
    
    filename = f"Report_{sheet_name.replace(' ', '_')}.pdf"
    path = os.path.join(output_path, filename)
    pdf.output(path)
    return path