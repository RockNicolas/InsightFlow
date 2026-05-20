import type { MaintenanceEquipment } from '../constants/maintenanceFleet'

import { EquipmentCell, type EquipmentFormData } from './EquipmentCell'
import { MaintenanceDateCell } from './MaintenanceDateCell'

interface EquipmentListProps {
  items: MaintenanceEquipment[]
  savingId: string | null
  onSaveMaintenance: (equipmentId: string, lastMaintenance: string) => Promise<void>
  onUpdateEquipment: (equipmentId: string, data: EquipmentFormData) => Promise<void>
  onDeleteEquipment: (equipmentId: string) => Promise<void>
  emptyLabel?: string
}

export function EquipmentList({
  items,
  savingId,
  onSaveMaintenance,
  onUpdateEquipment,
  onDeleteEquipment,
  emptyLabel = 'Nenhum equipamento cadastrado.',
}: EquipmentListProps) {
  if (items.length === 0) {
    return (
      <div className="empty-state">
        <p>{emptyLabel}</p>
      </div>
    )
  }

  return (
    <div className="equipment-table-wrap">
      <table className="equipment-table">
        <thead>
          <tr>
            <th scope="col">Equipamento / Placa</th>
            <th scope="col" className="equipment-table__maintenance-head">
              Última manutenção
            </th>
            <th scope="col" className="equipment-table__fill-head" aria-hidden />
          </tr>
        </thead>
        <tbody>
          {items.map((item) => (
            <tr key={item.id}>
              <td className="equipment-table__equipment">
                <EquipmentCell
                  item={item}
                  saving={savingId === item.id}
                  onUpdate={onUpdateEquipment}
                  onSaveMaintenance={onSaveMaintenance}
                  onDelete={onDeleteEquipment}
                />
              </td>
              <td className="equipment-table__maintenance">
                <MaintenanceDateCell item={item} />
              </td>
              <td className="equipment-table__fill" aria-hidden />
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
