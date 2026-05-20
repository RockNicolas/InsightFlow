import { useEffect, useState } from 'react'

import type { MaintenanceEquipment } from '../constants/maintenanceFleet'

import { ConfirmDialog } from './ConfirmDialog'
import { EditEquipmentModal, type EditEquipmentModalData } from './EditEquipmentModal'
import { PencilIcon, TrashIcon } from './MaintenanceActionIcons'

export interface EquipmentFormData {
  type: string
  code: string
  note: string
  alert: boolean
}

interface EquipmentCellProps {
  item: MaintenanceEquipment
  saving: boolean
  onUpdate: (equipmentId: string, data: EquipmentFormData) => Promise<void>
  onSaveMaintenance: (equipmentId: string, lastMaintenance: string) => Promise<void>
  onDelete: (equipmentId: string) => Promise<void>
}

function formatLabel(item: MaintenanceEquipment): string {
  if (item.note) {
    return `${item.type} — ${item.note} — ${item.code}`
  }
  return `${item.type} — ${item.code}`
}

function formatDisplay(item: MaintenanceEquipment) {
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

export function EquipmentCell({
  item,
  saving,
  onUpdate,
  onSaveMaintenance,
  onDelete,
}: EquipmentCellProps) {
  const [modalOpen, setModalOpen] = useState(false)
  const [confirmOpen, setConfirmOpen] = useState(false)

  const equipmentLabel = formatLabel(item)

  useEffect(() => {
    setModalOpen(false)
    setConfirmOpen(false)
  }, [item.id, item.type, item.code, item.note, item.alert, item.lastMaintenance])

  const handleSave = async (data: EditEquipmentModalData) => {
    const equipmentPayload = {
      type: data.type,
      code: data.code,
      note: data.note,
      alert: data.alert,
    }

    const equipmentChanged =
      equipmentPayload.type !== item.type ||
      equipmentPayload.code !== item.code ||
      equipmentPayload.note !== (item.note ?? '') ||
      equipmentPayload.alert !== Boolean(item.alert)

    const maintenanceChanged = data.lastMaintenance !== (item.lastMaintenance || '')

    if (!equipmentChanged && !maintenanceChanged) {
      setModalOpen(false)
      return
    }

    if (equipmentChanged) {
      await onUpdate(item.id, equipmentPayload)
    }

    if (maintenanceChanged) {
      await onSaveMaintenance(item.id, data.lastMaintenance)
    }

    setModalOpen(false)
  }

  const handleConfirmDelete = async () => {
    await onDelete(item.id)
    setConfirmOpen(false)
  }

  return (
    <>
      <div className="equipment-cell">
        <span className="equipment-cell__label">{formatDisplay(item)}</span>
        <div className="equipment-cell__icons">
          <button
            type="button"
            className="maintenance-icon-btn maintenance-icon-btn--edit"
            onClick={() => setModalOpen(true)}
            disabled={saving}
            title="Editar equipamento"
            aria-label={`Editar equipamento ${equipmentLabel}`}
          >
            <PencilIcon />
          </button>
          <button
            type="button"
            className="maintenance-icon-btn maintenance-icon-btn--delete"
            onClick={() => setConfirmOpen(true)}
            disabled={saving}
            title="Excluir equipamento"
            aria-label={`Excluir equipamento ${equipmentLabel}`}
          >
            <TrashIcon />
          </button>
        </div>
      </div>

      <EditEquipmentModal
        open={modalOpen}
        item={item}
        saving={saving}
        onCancel={() => setModalOpen(false)}
        onSave={handleSave}
      />

      <ConfirmDialog
        open={confirmOpen}
        title="Excluir equipamento?"
        message="O cadastro será removido da frota, incluindo o histórico de manutenção de"
        highlight={equipmentLabel}
        confirmLabel="Excluir"
        cancelLabel="Cancelar"
        variant="danger"
        loading={saving}
        onCancel={() => setConfirmOpen(false)}
        onConfirm={() => void handleConfirmDelete()}
      />
    </>
  )
}
