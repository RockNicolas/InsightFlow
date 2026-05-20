import type { ReactNode } from 'react'

import type { MaintenanceEquipment } from '../constants/maintenanceFleet'

import './EquipmentList.scss'

interface EquipmentListProps {
  items: MaintenanceEquipment[]
  emptyLabel?: string
}

function formatEquipment(item: MaintenanceEquipment): ReactNode {
  const sep = <span className="equipment-list__sep"> — </span>

  if (item.note) {
    return (
      <>
        <span className="equipment-table__name">{item.type}</span>
        {sep}
        <span className={item.alert ? 'equipment-list__alert' : undefined}>{item.note}</span>
        {sep}
        <span className="equipment-list__code">{item.code}</span>
      </>
    )
  }

  return (
    <>
      <span className="equipment-table__name">{item.type}</span>
      {sep}
      <span
        className={
          item.alert ? 'equipment-list__code equipment-list__alert' : 'equipment-list__code'
        }
      >
        {item.code}
      </span>
    </>
  )
}

function formatLastMaintenance(value?: string | null): ReactNode {
  if (!value || !value.trim()) {
    return <span className="equipment-list__empty">Sem registro</span>
  }
  return value
}

export function EquipmentList({ items, emptyLabel = 'Nenhum equipamento cadastrado.' }: EquipmentListProps) {
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
            <tr key={`${item.type}-${item.code}-${item.note ?? ''}`}>
              <td className="equipment-table__equipment">{formatEquipment(item)}</td>
              <td className="equipment-table__maintenance">
                {formatLastMaintenance(item.lastMaintenance)}
              </td>
              <td className="equipment-table__spacer" aria-hidden />
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
