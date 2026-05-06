import { useMemo, useState } from 'react'
import Anomalies from './Anomalies'
import AnomalyBulletins from './AnomalyBulletins'
import ReportingAgent from './ReportingAgent'
import ControllerReporting from './ControllerReporting'
import ControllerReportingList from './ControllerReportingList'
import PlanControle from './PlanControle'
import PlanActionCorrectif from './PlanActionCorrectif'
import './Dashboard.css'

interface DashboardProps {
  userName?: string
  userRole?: string
  userEmail?: string
}

type Tab =
  | 'anomalie'
  | 'bulletins'
  | 'reporting-agent'
  | 'reporting-saisie'
  | 'reporting-validation'
  | 'plan-controle'
  | 'plan-action-correctif'

interface NavEntry {
  key: Tab
  label: string
}

const MANAGER_ROLES = ['Chef_Departement', 'Directeur']

export default function Dashboard({ userName, userRole, userEmail }: DashboardProps) {
  const isManager = !!userRole && MANAGER_ROLES.includes(userRole)
  const isController = userRole === 'Controleur'

  const navItems = useMemo<NavEntry[]>(() => {
    const items: NavEntry[] = [
      { key: 'anomalie', label: 'Anomalies' },
      { key: 'bulletins', label: "Bulletins d'anomalies" },
      { key: 'reporting-agent', label: 'Reporting par Agent' },
    ]
    if (isController || isManager) {
      items.push({ key: 'reporting-saisie', label: 'Saisie journal' })
    }
    if (isManager) {
      items.push({ key: 'reporting-validation', label: 'Validation reportings' })
    }
    items.push(
      { key: 'plan-controle', label: 'Plan de Contrôle' },
      { key: 'plan-action-correctif', label: "Plan d'Action Correctif" },
    )
    return items
  }, [isManager, isController])

  const [activeTab, setActiveTab] = useState<Tab>('anomalie')

  return (
    <div className="dashboard">
      <header className="dashboard-topbar">
        <div className="topbar-left">
          <h1 className="topbar-title">ReportingDCPO</h1>
        </div>
        <div className="topbar-user">
          <span className="topbar-user-name">{userName}</span>
          <span className="topbar-user-job">{userRole}</span>
        </div>
      </header>

      <div className="dashboard-body">
        <nav className="dashboard-nav" aria-label="Navigation principale">
          {navItems.map(item => (
            <button
              key={item.key}
              type="button"
              className={`nav-item ${activeTab === item.key ? 'active' : ''}`}
              onClick={() => setActiveTab(item.key)}
              aria-current={activeTab === item.key ? 'page' : undefined}
            >
              {item.label}
            </button>
          ))}
        </nav>

        <div className="dashboard-content">
          {activeTab === 'anomalie' && (
            <Anomalies userName={userName} userEmail={userEmail} />
          )}
          {activeTab === 'bulletins' && <AnomalyBulletins />}
          {activeTab === 'reporting-agent' && <ReportingAgent />}
          {activeTab === 'reporting-saisie' && (
            <ControllerReporting userName={userName} userEmail={userEmail} />
          )}
          {activeTab === 'reporting-validation' && (
            <ControllerReportingList userName={userName} userEmail={userEmail} />
          )}
          {activeTab === 'plan-controle' && <PlanControle />}
          {activeTab === 'plan-action-correctif' && <PlanActionCorrectif />}
        </div>
      </div>
    </div>
  )
}
