import type { MaintenanceEquipment } from '../constants/maintenanceFleet'

interface MaintenanceDateCellProps {
  item: MaintenanceEquipment
}

export function MaintenanceDateCell({ item }: MaintenanceDateCellProps) {
  return (
    <div className="maintenance-date-cell">
      <span className="maintenance-date-cell__value">
        {item.lastMaintenance?.trim() ? (
          item.lastMaintenance
        ) : (
          <span className="equipment-list__empty">Sem registro</span>
        )}
      </span>
    </div>
  )
}
