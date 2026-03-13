import os
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT

def create_observation_report(excel_path, sheet_name, output_folder):
    pdf_path = os.path.join(output_folder, f"Relatorio_Obs_{sheet_name}.pdf")
    logo_path = os.path.join("assets", "company", "company_2.png")

    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None).fillna("")
        
        doc = SimpleDocTemplate(
            pdf_path, 
            pagesize=landscape(A4),
            rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
        )
        elements = []
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle('Title', parent=styles['Title'], fontSize=18, 
        alignment=TA_RIGHT, rightIndent=230)
        titulo = Paragraph(f"CONTROLE DE OBSERVAÇÕES MENSAL", title_style)
        
        if os.path.exists(logo_path):
            img = Image(logo_path, width=150, height=80) 
            header_table = Table([[img, titulo]], colWidths=[160, doc.width - 160])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ]))
            elements.append(header_table)
            elements.append(Spacer(1, 15))
        else:
            elements.append(titulo)
            elements.append(Spacer(1, 10))

        date_style = ParagraphStyle(
            'DateStyle',
            fontName='Helvetica-Bold',
            fontSize=16,
            textColor=colors.whitesmoke,
            alignment=TA_RIGHT,
            rightIndent=-45
        )

        def process_val(val):
            s = str(val).strip()
            if s in ["0", "0.0", "00", "0,0", "nan", "NaN", "None", ""]:
                return ""
            try:
                if "-" in s and len(s) >= 10:
                    return pd.to_datetime(s).strftime('%d/%m/%Y')
            except:
                pass
            if s.endswith(".0"):
                s = s[:-2]
            return s

        processed_data = []
        blue_rows = [] 

        for i, row in df.iterrows():
            clean_row = [process_val(item) for item in row]
            is_date_row = any('/202' in str(cell) for cell in clean_row)
            
            if any(str(c) != "" for c in clean_row):
                if is_date_row:
                    blue_rows.append(len(processed_data))
                    clean_row = [Paragraph(c, date_style) if '/' in str(c) else c for c in clean_row]
                processed_data.append(clean_row)

        if not processed_data:
            return None

        num_cols = len(processed_data[0])
        col_width = doc.width / num_cols
        table = Table(processed_data, colWidths=[col_width] * num_cols, hAlign='CENTER')
        
        main_style = [
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('OUTLINE', (0, 0), (-1, -1), 1, colors.black),
            ('VERTICALGRID', (0, 0), (-1, -1), 1, colors.black),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black), 
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ]

        for row_idx in blue_rows:
            main_style.append(('BACKGROUND', (0, row_idx), (-1, row_idx), colors.HexColor('#2E5A88')))
            main_style.append(('LINEBELOW', (0, row_idx), (-1, row_idx), 1.5, colors.black))
            main_style.append(('TOPPADDING', (0, row_idx), (-1, row_idx), 12))
            main_style.append(('BOTTOMPADDING', (0, row_idx), (-1, row_idx), 12))

        table.setStyle(TableStyle(main_style))
        elements.append(table)
        
        doc.build(elements)
        return pdf_path

    except Exception as e:
        print(f"[-] Erro: {e}")
        return None