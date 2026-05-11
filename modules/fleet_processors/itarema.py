from modules.data_processor import get_weekly_data
from modules.fleet_common import detect_observation_sheet
from modules.fleet_interface import (
    create_monthly_report as shared_create_monthly_report,
    create_observation_report as shared_create_observation_report,
    create_weekly_report as shared_create_weekly_report,
)


PROFILE_LABEL = "Itarema"


def create_weekly_report(excel_path, sheet_name, output_folder):
    return shared_create_weekly_report(excel_path, sheet_name, output_folder, get_weekly_data)


def create_monthly_report(excel_path, obs_sheet_name, output_folder, selected_sheets=None):
    return shared_create_monthly_report(
        excel_path,
        obs_sheet_name,
        output_folder,
        get_weekly_data,
        selected_sheets=selected_sheets,
    )


def create_observation_report(excel_path, sheet_name, output_folder):
    return shared_create_observation_report(excel_path, sheet_name, output_folder)
