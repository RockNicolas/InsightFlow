import { useCallback, useEffect, useState } from 'react'
import { Link } from 'react-router-dom'

import {
  createEquipment,
  deleteEquipment,
  fetchMaintenanceFleet,
  saveMaintenanceRecord,
  updateEquipment,
} from '../api/client'
import { AddEquipmentForm, type AddEquipmentFormData } from '../components/AddEquipmentForm'
import type { EquipmentFormData } from '../components/EquipmentCell'
import { AppNav } from '../components/AppNav'
import { CategoryTabs } from '../components/CategoryTabs'
import { EquipmentList } from '../components/EquipmentList'
import { Messages } from '../components/Messages'
import {
  DEFAULT_MAINTENANCE_CATEGORY,
  MAINTENANCE_CATEGORIES,
  type MaintenanceCategory,
} from '../constants/maintenanceCategories'
import {
  DEFAULT_MAINTENANCE_FLEET,
  type MaintenanceFleetMap,
} from '../constants/maintenanceFleet'

export function MaintenancePage() {
  const [activeCategory, setActiveCategory] =
    useState<MaintenanceCategory>(DEFAULT_MAINTENANCE_CATEGORY)
  const [fleet, setFleet] = useState<MaintenanceFleetMap>(DEFAULT_MAINTENANCE_FLEET)
  const [loading, setLoading] = useState(true)
  const [savingId, setSavingId] = useState<string | null>(null)
  const [addingEquipment, setAddingEquipment] = useState(false)
  const [showAddForm, setShowAddForm] = useState(false)
  const [messages, setMessages] = useState<Array<{ type: 'success' | 'error'; text: string }>>([])

  const current = MAINTENANCE_CATEGORIES.find((item) => item.id === activeCategory)!
  const items = fleet[activeCategory] ?? []

  const loadFleet = useCallback(async () => {
    setLoading(true)
    try {
      const response = await fetchMaintenanceFleet()
      if (response.ok) {
        setFleet(response.data.fleet)
      } else {
        setFleet(DEFAULT_MAINTENANCE_FLEET)
        setMessages([
          {
            type: 'error',
            text:
              response.message ||
              'API indisponível. Reinicie o servidor: pare com Ctrl+C e rode python main.py de novo.',
          },
        ])
      }
    } catch {
      setFleet(DEFAULT_MAINTENANCE_FLEET)
      setMessages([
        {
          type: 'error',
          text: 'Servidor desatualizado? Reinicie com Ctrl+C e execute python main.py novamente.',
        },
      ])
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    loadFleet()
  }, [loadFleet])

  const handleSaveMaintenance = async (equipmentId: string, lastMaintenance: string) => {
    setSavingId(equipmentId)
    setMessages([])
    try {
      const response = await saveMaintenanceRecord({
        category: activeCategory,
        equipment_id: equipmentId,
        last_maintenance: lastMaintenance,
      })
      if (response.ok) {
        setFleet(response.data.fleet)
        setMessages([{ type: 'success', text: response.message }])
      } else {
        setMessages([
          {
            type: 'error',
            text:
              response.message ||
              'Nao foi possivel salvar. Reinicie o servidor (Ctrl+C e python main.py).',
          },
        ])
      }
    } catch {
      setMessages([
        {
          type: 'error',
          text: 'Nao foi possivel salvar. Reinicie o servidor (Ctrl+C e python main.py).',
        },
      ])
    } finally {
      setSavingId(null)
    }
  }

  const handleUpdateEquipment = async (equipmentId: string, data: EquipmentFormData) => {
    setSavingId(equipmentId)
    setMessages([])
    try {
      const response = await updateEquipment({
        category: activeCategory,
        equipment_id: equipmentId,
        type: data.type,
        code: data.code,
        note: data.note,
        alert: data.alert,
      })
      if (response.ok) {
        setFleet(response.data.fleet)
        setMessages([{ type: 'success', text: response.message }])
      } else {
        setMessages([{ type: 'error', text: response.message }])
      }
    } catch (error) {
      const text =
        error instanceof Error ? error.message : 'Nao foi possivel atualizar o equipamento.'
      setMessages([{ type: 'error', text }])
    } finally {
      setSavingId(null)
    }
  }

  const handleDeleteEquipment = async (equipmentId: string) => {
    setSavingId(equipmentId)
    setMessages([])
    try {
      const response = await deleteEquipment({
        category: activeCategory,
        equipment_id: equipmentId,
      })
      if (response.ok) {
        setFleet(response.data.fleet)
        setMessages([{ type: 'success', text: response.message }])
      } else {
        setMessages([{ type: 'error', text: response.message }])
      }
    } catch (error) {
      const text =
        error instanceof Error ? error.message : 'Nao foi possivel excluir o equipamento.'
      setMessages([{ type: 'error', text }])
    } finally {
      setSavingId(null)
    }
  }

  const handleAddEquipment = async (data: AddEquipmentFormData) => {
    setAddingEquipment(true)
    setMessages([])

    try {
      const response = await createEquipment({
        category: activeCategory,
        type: data.type,
        code: data.code,
        note: data.note,
        alert: data.alert,
      })

      if (response.ok) {
        setFleet(response.data.fleet)
        setMessages([{ type: 'success', text: response.message }])
        setShowAddForm(false)
      } else {
        setMessages([{ type: 'error', text: response.message }])
      }
    } catch (error) {
      const text =
        error instanceof Error
          ? error.message
          : 'Nao foi possivel cadastrar. Reinicie o servidor (Ctrl+C e python main.py).'
      setMessages([{ type: 'error', text }])
    } finally {
      setAddingEquipment(false)
    }
  }

  return (
    <main className="page page--wide">
      <AppNav />

      <section className="hero hero-with-action">
        <div>
          <span className="eyebrow">InsightFlow</span>
          <h1>Manutenção de frota</h1>
          <p>
            Cadastre equipamentos e registre a última manutenção. Tudo fica salvo no banco
            PostgreSQL (Neon).
          </p>
        </div>
        <Link to="/" className="button secondary">
          Relatórios de frota
        </Link>
      </section>

      <Messages items={messages} />

      <section className="card maintenance-card">
        <CategoryTabs
          tabs={MAINTENANCE_CATEGORIES.map(({ id, label }) => ({ id, label }))}
          active={activeCategory}
          onChange={setActiveCategory}
        />

        <div className="card-header">
          <h2>{current.title}</h2>
          <p>{current.description}</p>
        </div>

        <div className="maintenance-stats">
          <article className={`stat-pill stat-pill--${current.id}`}>
            <span className="label">Categoria</span>
            <strong>{current.label}</strong>
          </article>
          <article className="stat-pill">
            <span className="label">Equipamentos</span>
            <strong>{items.length}</strong>
          </article>
          <article className="stat-pill">
            <span className="label">Com manutenção</span>
            <strong>
              {items.filter((item) => item.lastMaintenance?.trim()).length}
            </strong>
          </article>
        </div>

        <AddEquipmentForm
          category={activeCategory}
          open={showAddForm}
          saving={addingEquipment}
          onToggle={() => setShowAddForm((value) => !value)}
          onSubmit={handleAddEquipment}
        />

        <div className="maintenance-panel" role="tabpanel">
          <span className="label maintenance-fleet-label">Frota cadastrada</span>
          {loading ? (
            <div className="empty-state">
              <p>Carregando frota...</p>
            </div>
          ) : (
            <EquipmentList
              items={items}
              savingId={savingId}
              onSaveMaintenance={handleSaveMaintenance}
              onUpdateEquipment={handleUpdateEquipment}
              onDeleteEquipment={handleDeleteEquipment}
            />
          )}
          <p className="field-note maintenance-fleet-note">
            Use o lápis na coluna Equipamento para editar nome, placa, observação e data de manutenção.
            A lixeira remove o cadastro da frota.
          </p>
        </div>
      </section>
    </main>
  )
}
