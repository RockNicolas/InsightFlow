import type { MaintenanceCategory } from './maintenanceCategories'

export interface MaintenanceEquipment {
  /** Nome ou modelo (coluna esquerda na planilha) */
  type: string
  /** Código MC, placa ou identificador (coluna direita) */
  code: string
  /** Texto intermediário, ex.: LOCADA em caçamba locada */
  note?: string
  /** Destaque vermelho no código ou na observação (PRANCHA, LOCADA) */
  alert?: boolean
  /** Data ou descrição da última manutenção (ex.: 15/03/2026) */
  lastMaintenance?: string | null
}

/** Frota de referência — aba Manutenção (não altera relatórios de frota). */
export const MAINTENANCE_FLEET: Record<MaintenanceCategory, MaintenanceEquipment[]> = {
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
    { type: 'CAÇAMBA', code: 'ASFAL1', note: 'LOCADA', alert: true },
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
