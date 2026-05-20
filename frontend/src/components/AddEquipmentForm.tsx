import { useState } from 'react'

import type { MaintenanceCategory } from '../constants/maintenanceCategories'
import { MAINTENANCE_CATEGORIES } from '../constants/maintenanceCategories'

export interface AddEquipmentFormData {
  type: string
  code: string
  note: string
  alert: boolean
}

interface AddEquipmentFormProps {
  category: MaintenanceCategory
  open: boolean
  saving: boolean
  onToggle: () => void
  onSubmit: (data: AddEquipmentFormData) => Promise<void>
}

export function AddEquipmentForm({
  category,
  open,
  saving,
  onToggle,
  onSubmit,
}: AddEquipmentFormProps) {
  const [type, setType] = useState('')
  const [code, setCode] = useState('')
  const [note, setNote] = useState('')
  const [alert, setAlert] = useState(false)

  const categoryLabel =
    MAINTENANCE_CATEGORIES.find((item) => item.id === category)?.label ?? category

  const reset = () => {
    setType('')
    setCode('')
    setNote('')
    setAlert(false)
  }

  const handleSubmit = async (event: React.FormEvent) => {
    event.preventDefault()
    await onSubmit({ type, code, note, alert })
    reset()
  }

  return (
    <div className="add-equipment">
      <button
        type="button"
        className={`button ${open ? 'secondary' : 'primary'} add-equipment__toggle`}
        onClick={onToggle}
        aria-expanded={open}
      >
        <span className="add-equipment__toggle-icon" aria-hidden="true">
          {open ? '×' : '+'}
        </span>
        {open ? 'Fechar cadastro' : 'Cadastrar equipamento'}
      </button>

      {open ? (
        <form
          className={`add-equipment__panel add-equipment__panel--${category}`}
          onSubmit={(event) => void handleSubmit(event)}
        >
          <header className="add-equipment__header">
            <div className="add-equipment__header-text">
              <span className="add-equipment__eyebrow">Novo cadastro</span>
              <h3 className="add-equipment__title">Adicionar à frota</h3>
              <p className="add-equipment__hint">
                O item entra em <strong>{categoryLabel}</strong>. Depois, use o lápis na tabela para
                registrar a manutenção.
              </p>
            </div>
            <span className="add-equipment__badge">{categoryLabel}</span>
          </header>

          <div className="add-equipment__grid">
            <label className="add-equipment__field">
              <span className="add-equipment__label">Equipamento</span>
              <input
                className="add-equipment__input"
                type="text"
                placeholder="Ex.: RETROESCAVADEIRA"
                value={type}
                onChange={(event) => setType(event.target.value)}
                disabled={saving}
                required
                autoComplete="off"
              />
            </label>

            <label className="add-equipment__field">
              <span className="add-equipment__label">Placa / Código</span>
              <input
                className="add-equipment__input"
                type="text"
                placeholder="Ex.: MC 12 ou RIB1F06"
                value={code}
                onChange={(event) => setCode(event.target.value)}
                disabled={saving}
                required
                autoComplete="off"
              />
            </label>

            <label className="add-equipment__field">
              <span className="add-equipment__label">
                Observação
                <em className="add-equipment__optional">opcional</em>
              </span>
              <input
                className="add-equipment__input"
                type="text"
                placeholder="Ex.: LOCADA"
                value={note}
                onChange={(event) => setNote(event.target.value)}
                disabled={saving}
                autoComplete="off"
              />
            </label>
          </div>

          <label className={`checkbox add-equipment__alert ${alert ? 'is-on' : ''}`}>
            <input
              type="checkbox"
              checked={alert}
              onChange={(event) => setAlert(event.target.checked)}
              disabled={saving}
            />
            <span className="checkbox-box" aria-hidden="true" />
            <span className="checkbox-content">
              <span className="checkbox-title">Destaque vermelho</span>
              <span className="checkbox-description">
                Marca o equipamento como PRANCHA ou LOCADA na lista.
              </span>
            </span>
            <span className="add-equipment__alert-dot" aria-hidden="true" />
          </label>

          <footer className="add-equipment__footer">
            <button
              type="button"
              className="button secondary"
              onClick={() => {
                reset()
                onToggle()
              }}
              disabled={saving}
            >
              Cancelar
            </button>
            <button type="submit" className="button primary" disabled={saving}>
              {saving ? 'Cadastrando...' : 'Cadastrar equipamento'}
            </button>
          </footer>
        </form>
      ) : null}
    </div>
  )
}
