import os
from dotenv import load_dotenv
from modules.data_processor import get_weekly_data
from modules.pdf_generator import create_pdf_report
from modules.observation_report import create_observation_report 

load_dotenv()

def main():
    input_folder = os.getenv("INPUT_FOLDER", "inputs")
    output_folder = os.getenv("OUTPUT_FOLDER", "outputs")
    filename = os.getenv("EXCEL_FILENAME") 
    sheet_weekly = os.getenv("SELECTED_SHEET")
    sheet_obs = os.getenv("ABA_OBSERVACOES")

    excel_path = os.path.join(input_folder, filename)

    print(f"\n>>> SYSTEM STARTED: InsightFlow")
    
    if not os.path.exists(excel_path):
        print(f"ERROR: File '{filename}' not found.")
        return
    if sheet_weekly:
        print(f"[*] Extracting frotas from: {sheet_weekly}")
        data = get_weekly_data(excel_path, sheet_weekly)
        if data:
            pdf_weekly = create_pdf_report(data, sheet_weekly, output_folder)
            print(f"SUCCESS: Weekly report generated: {pdf_weekly}")

    if sheet_obs:
        print(f"\n[*] Processing Observations from: {sheet_obs}")
        pdf_obs = create_observation_report(excel_path, sheet_obs, output_folder)
        if pdf_obs:
            print(f"SUCCESS: Observation report generated: {pdf_obs}\n")

if __name__ == "__main__":
    main()