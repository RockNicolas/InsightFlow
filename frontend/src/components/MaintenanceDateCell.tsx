import { useEffect, useRef, useState } from 'react'

import type { MaintenanceEquipment } from '../constants/maintenanceFleet'
import { displayToInputValue, inputValueToDisplay } from '../utils/maintenanceDate'

import { ConfirmDialog } from './ConfirmDialog'
import { PencilIcon, TrashIcon } from './MaintenanceActionIcons'
interface MaintenanceDateCellProps {
  item: MaintenanceEquipment
  saving: boolean
  onSave: (equipmentId: string, lastMaintenance: string) => Promise<void>
}

export function MaintenanceDateCell({
  item,
  saving,
  onSave,
}: MaintenanceDateCellProps) {
  const [editing, setEditing] = useState(false)
  const [confirmOpen, setConfirmOpen] = useState(false)
  const [inputValue, setInputValue] = useState(() => displayToInputValue(item.lastMaintenance))
  const inputRef = useRef<HTMLInputElement>(null)

  const equipmentLabel = `${item.type} — ${item.code}`

  useEffect(() => {
    setInputValue(displayToInputValue(item.lastMaintenance))
    setEditing(false)
    setConfirmOpen(false)
  }, [item.lastMaintenance, item.id])

  useEffect(() => {
    if (editing) {
      inputRef.current?.focus()
    }
  }, [editing])

  const handleEdit = () => {
    setInputValue(displayToInputValue(item.lastMaintenance))
    setEditing(true)
  }

  const handleDateCommit = async (value: string) => {
    const display = inputValueToDisplay(value)
    if (display === (item.lastMaintenance || '')) {
      setEditing(false)
      return
    }
    await onSave(item.id, display)
    setEditing(false)
  }

  const handleConfirmDelete = async () => {
    await onSave(item.id, '')
    setConfirmOpen(false)
  }

  if (editing) {
    return (
      <div className="maintenance-date-cell maintenance-date-cell--editing">
        <input
          ref={inputRef}
          type="date"
          className="maintenance-date-cell__input"
          value={inputValue}
          onChange={(event) => setInputValue(event.target.value)}
          onBlur={() => handleDateCommit(inputValue)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              void handleDateCommit(inputValue)
            }
            if (event.key === 'Escape') {
              setEditing(false)
            }
          }}
          disabled={saving}
          aria-label={`Editar manutenção de ${item.type} ${item.code}`}
        />
      </div>
    )
  }

  return (
    <>
      <div className="maintenance-date-cell">
        <span className="maintenance-date-cell__value">
          {item.lastMaintenance?.trim() ? (
            item.lastMaintenance
          ) : (
            <span className="equipment-list__empty">Sem registro</span>
          )}
        </span>
        <div className="maintenance-date-cell__icons">
          <button
            type="button"
            className="maintenance-icon-btn maintenance-icon-btn--edit"
            onClick={handleEdit}
            disabled={saving}
            title="Editar manutenção"
            aria-label={`Editar manutenção de ${item.type} ${item.code}`}
          >
            <PencilIcon />
          </button>
          <button
            type="button"
            className="maintenance-icon-btn maintenance-icon-btn--delete"
            onClick={() => setConfirmOpen(true)}
            disabled={saving || !item.lastMaintenance}
            title="Excluir manutenção"
            aria-label={`Excluir manutenção de ${item.type} ${item.code}`}
          >
            <TrashIcon />
          </button>
        </div>
      </div>

      <ConfirmDialog
        open={confirmOpen}
        title="Excluir manutenção?"
        message="O registro de última manutenção será removido para"
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
