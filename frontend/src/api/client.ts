import type {
  CreateEquipmentPayload,
  DeleteEquipmentPayload,
  MaintenanceFleetResponse,
  UpdateEquipmentPayload,
  UpdateMaintenancePayload,
} from '../constants/maintenanceFleet'
import type { ApiResponse, GeneratePayload, SessionState } from '../types'

async function request<T>(url: string, init?: RequestInit): Promise<ApiResponse<T>> {
  const response = await fetch(url, {
    credentials: 'include',
    ...init,
  })

  let payload: ApiResponse<T>
  try {
    payload = (await response.json()) as ApiResponse<T>
  } catch {
    if (response.status === 405) {
      throw new Error(
        'Cadastro indisponível: reinicie o servidor (Ctrl+C no terminal e python main.py).',
      )
    }
    throw new Error('Falha na comunicação com o servidor.')
  }

  if (!response.ok) {
    if (response.status === 405) {
      throw new Error(
        'Cadastro indisponível: reinicie o servidor (Ctrl+C no terminal e python main.py).',
      )
    }
    if (payload.ok === undefined) {
      throw new Error(payload.message || 'Falha na comunicação com o servidor.')
    }
  }
  return payload
}

export async function fetchSession(): Promise<ApiResponse<SessionState>> {
  return request<SessionState>('/api/session')
}

export async function uploadWorkbook(
  fleetProfile: string,
  file: File,
): Promise<ApiResponse<SessionState>> {
  const formData = new FormData()
  formData.append('fleet_profile', fleetProfile)
  formData.append('excel_file', file)

  return request<SessionState>('/api/load-sheets', {
    method: 'POST',
    body: formData,
  })
}

export async function generateReports(
  payload: GeneratePayload,
): Promise<ApiResponse<SessionState>> {
  return request<SessionState>('/api/generate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

export function downloadUrl(path: string): string {
  return `/downloads/${path}`
}

export async function fetchMaintenanceFleet(): Promise<ApiResponse<MaintenanceFleetResponse>> {
  return request<MaintenanceFleetResponse>('/api/maintenance/fleet')
}

export async function saveMaintenanceRecord(
  payload: UpdateMaintenancePayload,
): Promise<ApiResponse<MaintenanceFleetResponse>> {
  return request<MaintenanceFleetResponse>('/api/maintenance/record', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

export async function createEquipment(
  payload: CreateEquipmentPayload,
): Promise<ApiResponse<MaintenanceFleetResponse>> {
  return request<MaintenanceFleetResponse>('/api/maintenance/equipment', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

export async function updateEquipment(
  payload: UpdateEquipmentPayload,
): Promise<ApiResponse<MaintenanceFleetResponse>> {
  return request<MaintenanceFleetResponse>('/api/maintenance/equipment', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}

export async function deleteEquipment(
  payload: DeleteEquipmentPayload,
): Promise<ApiResponse<MaintenanceFleetResponse>> {
  return request<MaintenanceFleetResponse>('/api/maintenance/equipment', {
    method: 'DELETE',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  })
}
