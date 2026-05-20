import { useEffect, useState } from 'react'

import type { MaintenanceEquipment } from '../constants/maintenanceFleet'
import { displayToInputValue, inputValueToDisplay } from '../utils/maintenanceDate'

import type { EquipmentFormData } from './EquipmentCell'

export interface EditEquipmentModalData extends EquipmentFormData {
  lastMaintenance: string
}

interface EditEquipmentModalProps {
  open: boolean
  item: MaintenanceEquipment
  saving: boolean
  onCancel: () => void
  onSave: (data: EditEquipmentModalData) => Promise<void>
}

export function EditEquipmentModal({
  open,
  item,
  saving,
  onCancel,
  onSave,
}: EditEquipmentModalProps) {
  const [type, setType] = useState(item.type)
  const [code, setCode] = useState(item.code)
  const [note, setNote] = useState(item.note ?? '')
  const [alert, setAlert] = useState(Boolean(item.alert))
  const [maintenanceInput, setMaintenanceInput] = useState(() =>
    displayToInputValue(item.lastMaintenance),
  )

  useEffect(() => {
    if (!open) {
      return
    }

    setType(item.type)
    setCode(item.code)
    setNote(item.note ?? '')
    setAlert(Boolean(item.alert))
    setMaintenanceInput(displayToInputValue(item.lastMaintenance))
  }, [open, item.id, item.type, item.code, item.note, item.alert, item.lastMaintenance])

  useEffect(() => {
    if (!open) {
      return
    }

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !saving) {
        onCancel()
      }
    }

    document.addEventListener('keydown', onKeyDown)
    document.body.style.overflow = 'hidden'

    return () => {
      document.removeEventListener('keydown', onKeyDown)
      document.body.style.overflow = ''
    }
  }, [open, saving, onCancel])

  if (!open) {
    return null
  }

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault()

    await onSave({
      type: type.trim(),
      code: code.trim(),
      note: note.trim(),
      alert,
      lastMaintenance: inputValueToDisplay(maintenanceInput),
    })
  }

  return (
    <div
      className="edit-equipment-modal__backdrop"
      role="presentation"
      onClick={() => {
        if (!saving) {
          onCancel()
        }
      }}
    >
      <div
        className="edit-equipment-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="edit-equipment-modal-title"
        onClick={(event) => event.stopPropagation()}
      >
        <h3 id="edit-equipment-modal-title" className="edit-equipment-modal__title">
          Editar equipamento
        </h3>

        <form className="edit-equipment-modal__form" onSubmit={(event) => void handleSubmit(event)}>
          <div className="edit-equipment-modal__fields">
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
            <label>
              <span>Última manutenção</span>
              <input
                type="date"
                value={maintenanceInput}
                onChange={(event) => setMaintenanceInput(event.target.value)}
                disabled={saving}
              />
              <span className="edit-equipment-modal__hint">
                Deixe em branco se ainda não houver registro.
              </span>
            </label>
          </div>

          <label className="edit-equipment-modal__alert">
            <input
              type="checkbox"
              checked={alert}
              onChange={(event) => setAlert(event.target.checked)}
              disabled={saving}
            />
            <span>Destaque vermelho</span>
          </label>

          <div className="edit-equipment-modal__actions">
            <button
              type="button"
              className="button secondary"
              onClick={onCancel}
              disabled={saving}
            >
              Cancelar
            </button>
            <button type="submit" className="button primary" disabled={saving}>
              {saving ? 'Salvando...' : 'Salvar'}
            </button>
          </div>
        </form>
      </div>
    </div>
  )
}
