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
        <colgroup>
          <col className="equipment-table__col-equipment" />
          <col className="equipment-table__col-maintenance" />
          <col className="equipment-table__col-spacer" />
        </colgroup>
        <thead>
          <tr>
            <th scope="col">Equipamento / Placa</th>
            <th scope="col">Última manutenção</th>
            <th scope="col" className="equipment-table__spacer-head" aria-hidden />
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
                  onDelete={onDeleteEquipment}
                />
              </td>
              <td className="equipment-table__maintenance">
                <MaintenanceDateCell
                  item={item}
                  saving={savingId === item.id}
                  onSave={onSaveMaintenance}
                />
              </td>
              <td className="equipment-table__spacer" aria-hidden />
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
