import os
import pandas as pd
from dotenv import load_dotenv
from modules.data_processor import get_weekly_data
from modules.pdf_generator import create_pdf_report
from modules.dashboard_generator import generate_monthly_dashboard 

load_dotenv()

def main():
    input_folder = os.getenv("INPUT_FOLDER", "inputs")
    output_folder = os.getenv("OUTPUT_FOLDER", "outputs")
    filename = os.getenv("EXCEL_FILENAME") 
    sheet_name = os.getenv("SELECTED_SHEET")

    excel_path = os.path.join(input_folder, filename)

    print(f"\n>>> SISTEMA INICIADO: InsightFlow")
    print("-" * 45)
    
    if not os.path.exists(excel_path):
        print(f"ERRO: Ficheiro '{filename}' não encontrado em {input_folder}.")
        return

    print(f"[*] Extraindo frotas da aba: {sheet_name}")
    data = get_weekly_data(excel_path, sheet_name)

    if not data:
        print("[-] Nenhum dado válido encontrado.")
        return

    print(f"[+] Gerando PDF Semanal...")
    pdf_file = create_pdf_report(data, sheet_name, output_folder)
    print(f"[OK] PDF gerado: {pdf_file}")

    df_semana = pd.DataFrame(data)
    df_semana['Semana'] = sheet_name 

    csv_name = f"data_{sheet_name.replace(' ', '_').replace('.', '_')}.csv"
    csv_path = os.path.join(output_folder, csv_name)
    df_semana.to_csv(csv_path, index=False)
    print(f"[*] Dados integrados na base histórica: {csv_name}")

    print("-" * 45)
    escolha = input("Deseja atualizar o Dashboard Comparativo Mensal agora? (S/N): ").upper()
    
    if escolha == 'S':
        print("[*] Consolidando semanas e criando gráficos...")
        img_dashboard = generate_monthly_dashboard(output_folder)
        if img_dashboard:
            print(f"SUCCESS: Dashboard profissional gerado em {img_dashboard}")
    
    print("\n>>> PROCESSO FINALIZADO COM SUCESSO!\n")

if __name__ == "__main__":
    main()