import { NavLink } from 'react-router-dom'

export function AppNav() {
  return (
    <nav className="app-nav" aria-label="Navegação principal">
      <NavLink
        to="/"
        end
        className={({ isActive }) => `nav-link${isActive ? ' is-active' : ''}`}
      >
        Relatórios de frota
      </NavLink>
      <NavLink
        to="/manutencao"
        className={({ isActive }) => `nav-link${isActive ? ' is-active' : ''}`}
      >
        Manutenção de veículos
      </NavLink>
    </nav>
  )
}
