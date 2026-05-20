import { useState } from 'react'
import { Link } from 'react-router-dom'

import { AppNav } from '../components/AppNav'
import { CategoryTabs } from '../components/CategoryTabs'
import {
  DEFAULT_MAINTENANCE_CATEGORY,
  MAINTENANCE_CATEGORIES,
  type MaintenanceCategory,
} from '../constants/maintenanceCategories'

export function MaintenancePage() {
  const [activeCategory, setActiveCategory] =
    useState<MaintenanceCategory>(DEFAULT_MAINTENANCE_CATEGORY)

  const current = MAINTENANCE_CATEGORIES.find((item) => item.id === activeCategory)!

  return (
    <main className="page">
      <AppNav />

      <section className="hero hero-with-action">
        <div>
          <span className="eyebrow">InsightFlow</span>
          <h1>Manutenção de frota</h1>
          <p>
            Área separada dos relatórios. Escolha o tipo de equipamento para registrar e
            acompanhar manutenções.
          </p>
        </div>
        <Link to="/" className="button secondary">
          Relatórios de frota
        </Link>
      </section>

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
            <span className="label">Medição</span>
            <strong>{current.unit === 'horas' ? 'Horímetro (horas)' : 'Quilometragem (km)'}</strong>
          </article>
        </div>

        <div className="empty-state maintenance-panel" role="tabpanel">
          <p>
            Em breve: cadastro de ordens de serviço, histórico e alertas para{' '}
            <strong>{current.label.toLowerCase()}</strong>.
          </p>
          <p className="field-note">
            Os relatórios semanais e mensais continuam apenas em &quot;Relatórios de frota&quot;.
          </p>
        </div>
      </section>
    </main>
  )
}
