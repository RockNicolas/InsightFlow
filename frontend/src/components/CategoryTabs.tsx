import type { MaintenanceCategory } from '../constants/maintenanceCategories'

interface CategoryTab {
  id: MaintenanceCategory
  label: string
}

interface CategoryTabsProps {
  tabs: CategoryTab[]
  active: MaintenanceCategory
  onChange: (id: MaintenanceCategory) => void
}

export function CategoryTabs({ tabs, active, onChange }: CategoryTabsProps) {
  return (
    <div className="category-tabs" role="tablist" aria-label="Tipo de equipamento">
      {tabs.map((tab) => (
        <button
          key={tab.id}
          type="button"
          role="tab"
          aria-selected={active === tab.id}
          className={`category-tab category-tab--${tab.id}${active === tab.id ? ' is-active' : ''}`}
          onClick={() => onChange(tab.id)}
        >
          {tab.label}
        </button>
      ))}
    </div>
  )
}
