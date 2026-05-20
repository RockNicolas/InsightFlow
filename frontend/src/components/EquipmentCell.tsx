import { useEffect, useState } from 'react'

import type { MaintenanceEquipment } from '../constants/maintenanceFleet'

import { ConfirmDialog } from './ConfirmDialog'
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

export function EquipmentCell({ item, saving, onUpdate, onDelete }: EquipmentCellProps) {
  const [editing, setEditing] = useState(false)
  const [confirmOpen, setConfirmOpen] = useState(false)
  const [type, setType] = useState(item.type)
  const [code, setCode] = useState(item.code)
  const [note, setNote] = useState(item.note ?? '')
  const [alert, setAlert] = useState(Boolean(item.alert))

  const equipmentLabel = formatLabel(item)

  useEffect(() => {
    setType(item.type)
    setCode(item.code)
    setNote(item.note ?? '')
    setAlert(Boolean(item.alert))
    setEditing(false)
    setConfirmOpen(false)
  }, [item.id, item.type, item.code, item.note, item.alert])

  const resetForm = () => {
    setType(item.type)
    setCode(item.code)
    setNote(item.note ?? '')
    setAlert(Boolean(item.alert))
  }

  const handleSave = async () => {
    const payload = {
      type: type.trim(),
      code: code.trim(),
      note: note.trim(),
      alert,
    }

    if (
      payload.type === item.type &&
      payload.code === item.code &&
      payload.note === (item.note ?? '') &&
      payload.alert === Boolean(item.alert)
    ) {
      setEditing(false)
      return
    }

    await onUpdate(item.id, payload)
    setEditing(false)
  }

  const handleConfirmDelete = async () => {
    await onDelete(item.id)
    setConfirmOpen(false)
  }

  if (editing) {
    return (
      <div className="equipment-cell equipment-cell--editing">
        <div className="equipment-cell__fields">
          <label>
            <span>Equipamento</span>
            <input
              type="text"
              value={type}
              onChange={(event) => setType(event.target.value)}
              disabled={saving}
              required
            />
          </label>
          <label>
            <span>Placa / Código</span>
            <input
              type="text"
              value={code}
              onChange={(event) => setCode(event.target.value)}
              disabled={saving}
              required
            />
          </label>
          <label>
            <span>Observação</span>
            <input
              type="text"
              value={note}
              onChange={(event) => setNote(event.target.value)}
              disabled={saving}
              placeholder="Opcional"
            />
          </label>
        </div>
        <label className="equipment-cell__alert">
          <input
            type="checkbox"
            checked={alert}
            onChange={(event) => setAlert(event.target.checked)}
            disabled={saving}
          />
          <span>Destaque vermelho</span>
        </label>
        <div className="equipment-cell__edit-actions">
          <button
            type="button"
            className="button secondary"
            onClick={() => {
              resetForm()
              setEditing(false)
            }}
            disabled={saving}
          >
            Cancelar
          </button>
          <button type="button" className="button primary" onClick={() => void handleSave()} disabled={saving}>
            {saving ? 'Salvando...' : 'Salvar'}
          </button>
        </div>
      </div>
    )
  }

  return (
    <>
      <div className="equipment-cell">
        <span className="equipment-cell__display">{formatDisplay(item)}</span>
        <div className="equipment-cell__icons">
          <button
            type="button"
            className="maintenance-icon-btn maintenance-icon-btn--edit"
            onClick={() => {
              resetForm()
              setEditing(true)
            }}
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
