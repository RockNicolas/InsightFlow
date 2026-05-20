import type { ApiResponse, GeneratePayload, SessionState } from '../types'

async function request<T>(url: string, init?: RequestInit): Promise<ApiResponse<T>> {
  const response = await fetch(url, {
    credentials: 'include',
    ...init,
  })

  const payload = (await response.json()) as ApiResponse<T>
  if (!response.ok && payload.ok === undefined) {
    throw new Error('Falha na comunicação com o servidor.')
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
