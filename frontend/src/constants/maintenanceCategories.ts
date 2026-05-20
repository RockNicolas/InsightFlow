export type MaintenanceCategory = 'machine' | 'truck' | 'vehicle'

export interface MaintenanceCategoryConfig {
  id: MaintenanceCategory
  label: string
  title: string
  description: string
  unit: 'horas' | 'km'
}

export const MAINTENANCE_CATEGORIES: MaintenanceCategoryConfig[] = [
  {
    id: 'machine',
    label: 'Máquinas',
    title: 'Manutenção de máquinas',
    description:
      'Controle preventivo e corretivo de máquinas pesadas, retroescavadeiras e equipamentos com horímetro.',
    unit: 'horas',
  },
  {
    id: 'truck',
    label: 'Caminhões',
    title: 'Manutenção de caminhões',
    description:
      'Ordens de serviço, revisões e histórico para caminhões, caçambas, pranchas e frota pesada rodoviária.',
    unit: 'km',
  },
  {
    id: 'vehicle',
    label: 'Veículos',
    title: 'Manutenção de veículos',
    description:
      'Acompanhamento de veículos leves e de apoio: revisões, quilometragem e registros por placa.',
    unit: 'km',
  },
]

export const DEFAULT_MAINTENANCE_CATEGORY: MaintenanceCategory = 'machine'
