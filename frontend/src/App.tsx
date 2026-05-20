import { BrowserRouter, Route, Routes } from 'react-router-dom'
import { MaintenancePage } from './pages/MaintenancePage'
import { ReportsPage } from './pages/ReportsPage'

export default function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<ReportsPage />} />
        <Route path="/manutencao" element={<MaintenancePage />} />
      </Routes>
    </BrowserRouter>
  )
}
