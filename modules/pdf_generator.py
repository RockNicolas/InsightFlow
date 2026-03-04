from fpdf import FPDF
import os

def create_pdf_report(data, sheet_name, output_path):
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    
    logo_path = "assets\company\company_2.png" 
    if os.path.exists(logo_path):
        pdf.image(logo_path, x=85, y=2, w=40)
        pdf.ln(25)
    else:
        print(f"Aviso: Logotipo {logo_path} não encontrado.")
        pdf.ln(10)

    pdf.set_font("Arial", "B", 16)
    pdf.cell(190, 15, f"RELATORIO SEMANAL - {sheet_name}", ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)
    pdf.cell(85, 10, "MÁQUINA/PLACA", 1, 0, 'C', True)
    pdf.cell(65, 10, "OPERADORES", 1, 0, 'C', True)
    pdf.cell(40, 10, "HORAS/KM", 1, 1, 'C', True)
    
    pdf.set_font("Arial", "", 9)
    total_week = 0
    
    lista_vermelha = [
        "RETROESCAVADEIRA MC01", 
        "PC 210 - ESCAVADEIRA MC 13",
        "MOBI - LOCADO"
    ]

    for item in data:
        if item['machine'] in lista_vermelha or item['hours'] == 0:
            pdf.set_text_color(255, 0, 0) 
            pdf.set_font("Arial", "B", 9)
        else:
            pdf.set_text_color(0, 0, 0)  
            pdf.set_font("Arial", "", 9)  

        pdf.cell(85, 10, str(item['machine'])[:45], 1)
        pdf.cell(65, 10, str(item['operator'])[:35], 1)
        pdf.cell(40, 10, f"{item['hours']:.2f}", 1, 1, 'R')
        
        total_week += item['hours']
        
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(150, 10, "TOTAL HORAS/KM", 1, 0, 'L', True)
    pdf.cell(40, 10, f"{total_week:.2f}", 1, 1, 'R', True)
    
    filename = f"Report_{sheet_name.replace(' ', '_')}.pdf"
    path = os.path.join(output_path, filename)
    pdf.output(path)
    return path