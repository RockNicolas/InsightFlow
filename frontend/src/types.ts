export interface FleetProfile {
  key: string
  label: string
}

export interface ReportResult {
  label: string
  filename: string
  download_path: string
}

export interface SessionState {
  uploaded_name: string
  fleet_profiles: FleetProfile[]
  selected_fleet_key: string
  selected_fleet_label: string
  sheet_names: string[]
  weekly_options: string[]
  weekly_sheet: string
  observation_sheet: string
  generate_all_month_sheets: boolean
  results: ReportResult[]
  output_folder: string
}

export interface GeneratePayload {
  weekly_sheet: string
  observation_sheet: string
  generate_weekly: boolean
  generate_observation: boolean
  generate_monthly: boolean
  generate_all_month_sheets: boolean
}

export interface ApiResponse<T> {
  ok: boolean
  message: string
  data: T
}
