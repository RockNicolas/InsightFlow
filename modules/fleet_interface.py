from modules.monthly_report import create_monthly_report as shared_create_monthly_report
from modules.observation_report import create_observation_report as shared_create_observation_report
from modules.pdf_generator import create_pdf_report as shared_create_pdf_report


def create_weekly_report(excel_path, sheet_name, output_folder, weekly_data_loader):
    weekly_data = weekly_data_loader(excel_path, sheet_name)
    if not weekly_data:
        return None
    return shared_create_pdf_report(weekly_data, sheet_name, output_folder)


def create_monthly_report(excel_path, obs_sheet_name, output_folder, weekly_data_loader, selected_sheets=None):
    return shared_create_monthly_report(
        excel_path,
        obs_sheet_name,
        output_folder,
        selected_sheets=selected_sheets,
        weekly_data_loader=weekly_data_loader,
    )


def create_observation_report(excel_path, sheet_name, output_folder):
    return shared_create_observation_report(excel_path, sheet_name, output_folder)
