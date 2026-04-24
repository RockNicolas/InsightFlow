import os
from dotenv import load_dotenv
from modules.monthly_report import create_monthly_report

load_dotenv()


def main():
    input_folder = os.getenv("INPUT_FOLDER", "inputs")
    output_folder = os.getenv("OUTPUT_FOLDER", "outputs")
    filename = os.getenv("EXCEL_FILENAME")
    sheet_obs = os.getenv("ABA_OBSERVACOES")

    print(f"\n>>> SYSTEM STARTED: InsightFlow - Relatório Mensal")

    if not filename:
        print("ERRO: Variável EXCEL_FILENAME não definida no .env")
        return

    excel_path = os.path.join(input_folder, filename)

    if not os.path.exists(excel_path):
        print(f"ERRO: Arquivo '{filename}' não encontrado em '{input_folder}'.")
        return

    print(f"[*] Excel: {excel_path}")
    print(f"[*] Outputs: {output_folder}")
    print(f"[*] Aba de observações ignorada: {sheet_obs}\n")

    pdf_path = create_monthly_report(excel_path, sheet_obs, output_folder)

    if pdf_path:
        print(f"\nSUCCESS: Relatório mensal gerado: {pdf_path}\n")
    else:
        print("\nERRO: Falha ao gerar relatório mensal.\n")


if __name__ == "__main__":
    main()
