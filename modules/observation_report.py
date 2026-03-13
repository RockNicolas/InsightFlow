import os
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

def create_observation_report(excel_path, sheet_name, output_folder):
    pdf_path = os.path.join(output_folder, f"Relatorio_Obs_{sheet_name}.pdf")
    
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name).fillna("")

        doc = SimpleDocTemplate(pdf_path, pagesize=landscape(A4))
        elements = []
        
        styles = getSampleStyleSheet()
        elements.append(Paragraph(f"Relatório de Observações - {sheet_name}", styles['Title']))

        data_list = [df.columns.values.tolist()] + df.values.tolist()

        table = Table(data_list, hAlign='CENTER')
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 1, colors.black), # Traços pretos
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        elements.append(table)
        doc.build(elements)
        return pdf_path
    except Exception as e:
        print(f"[-] Erro ao gerar observações: {e}")
        return None