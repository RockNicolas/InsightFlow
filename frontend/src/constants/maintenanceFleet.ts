import type { MaintenanceCategory } from './maintenanceCategories'

export interface MaintenanceEquipment {
  id: string
  type: string
  code: string
  note?: string
  alert?: boolean
  lastMaintenance?: string | null
}

export type MaintenanceFleetMap = Record<MaintenanceCategory, MaintenanceEquipment[]>

export interface MaintenanceFleetResponse {
  fleet: MaintenanceFleetMap
}

export interface UpdateMaintenancePayload {
  category: MaintenanceCategory
  equipment_id: string
  last_maintenance: string
}

export interface CreateEquipmentPayload {
  category: MaintenanceCategory
  type: string
  code: string
  note?: string
  alert?: boolean
}

export interface UpdateEquipmentPayload {
  category: MaintenanceCategory
  equipment_id: string
  type: string
  code: string
  note?: string
  alert?: boolean
}

export interface DeleteEquipmentPayload {
  category: MaintenanceCategory
  equipment_id: string
}

type FleetSeed = Omit<MaintenanceEquipment, 'id' | 'lastMaintenance'>[]

const FLEET_SEED: Record<MaintenanceCategory, FleetSeed> = {
  machine: [
    { type: 'RETROESCAVADEIRA', code: 'MC 01' },
    { type: 'RETROESCAVADEIRA', code: 'MC 02' },
    { type: 'RETROESCAVADEIRA', code: 'MC 03' },
    { type: 'RETROESCAVADEIRA', code: 'MC 06' },
    { type: 'RETROESCAVADEIRA', code: 'MC 08' },
    { type: 'RETROESCAVADEIRA', code: 'MC 09' },
    { type: 'RETROESCAVADEIRA', code: 'MC 10' },
    { type: 'RETROESCAVADEIRA', code: 'MC 11' },
  ],
  truck: [
    { type: 'CAMINHÃO', code: 'PRANCHA', alert: true },
    { type: 'CAÇAMBA', code: 'RIH7F79' },
    { type: 'CAÇAMBA', code: 'RIL0A98' },
    { type: 'VOLVO', code: 'SBC3I31' },
    { type: 'MICROÔNIBUS', code: 'LLF8B75' },
    { type: 'M.BENZ - ÔNIBUS', code: 'PRÓPRIO' },
    { type: 'CAÇAMBA', code: 'ASFALTO', note: 'LOCADA', alert: true },
  ],
  vehicle: [
    { type: 'STRADA', code: 'RIB1F06' },
    { type: 'STRADA', code: 'PNV1A69' },
    { type: 'UNO VIVACE', code: 'OHY8267' },
    { type: 'GOL', code: 'OCB0H10' },
    { type: 'MOBI', code: 'RIB1I10' },
    { type: 'MOBI', code: 'RIA2I10' },
  ],
}

export function buildEquipmentId(
  category: MaintenanceCategory,
  item: Pick<MaintenanceEquipment, 'type' | 'code' | 'note'>,
): string {
  return `${category}:${item.type}|${item.code}|${item.note ?? ''}`
}

export function createDefaultFleet(): MaintenanceFleetMap {
  const fleet = {} as MaintenanceFleetMap
  for (const category of Object.keys(FLEET_SEED) as MaintenanceCategory[]) {
    fleet[category] = FLEET_SEED[category].map((item) => ({
      ...item,
      id: buildEquipmentId(category, item),
      lastMaintenance: null,
    }))
  }
  return fleet
}

export const DEFAULT_MAINTENANCE_FLEET = createDefaultFleet()
