import os
from dotenv import load_dotenv
from modules.data_processor import get_weekly_data
from modules.pdf_generator import create_pdf_report

load_dotenv()

def main():
    input_folder = os.getenv("INPUT_FOLDER", "inputs")
    output_folder = os.getenv("OUTPUT_FOLDER", "outputs")
    filename = os.getenv("EXCEL_FILENAME") 
    sheet_name = os.getenv("SELECTED_SHEET")

    excel_path = os.path.join(input_folder, filename)

    print(f"\n>>> SYSTEM STARTED: InsightFlow")
    
    if not os.path.exists(excel_path):
        print(f"ERROR: File '{filename}' not found.")
        return

    print(f"[*] Extracting all 42 frotas from: {sheet_name}")
    data = get_weekly_data(excel_path, sheet_name)

    if not data:
        print("[-] No valid data found.")
        return

    print(f"[+] Concatenating names and generating PDF...")
    pdf_file = create_pdf_report(data, sheet_name, output_folder)
    print(f"SUCCESS: Integrated report generated at {pdf_file}\n")

if __name__ == "__main__":
    main()