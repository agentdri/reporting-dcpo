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

interface NavLeaf {
  type: 'leaf'
  key: Tab
  label: string
}

interface NavGroup {
  type: 'group'
  key: string
  label: string
  children: NavLeaf[]
}

type NavEntry = NavLeaf | NavGroup

const MANAGER_ROLES = ['Chef_Departement', 'Directeur']

export default function Dashboard({ userName, userRole, userEmail }: DashboardProps) {
  const isManager = !!userRole && MANAGER_ROLES.includes(userRole)
  const isController = userRole === 'Controleur'

  const navItems = useMemo<NavEntry[]>(() => {
    const items: NavEntry[] = [
      {
        type: 'group',
        key: 'anomalies',
        label: 'Anomalies',
        children: [
          { type: 'leaf', key: 'anomalie', label: 'Liste des anomalies' },
          { type: 'leaf', key: 'bulletins', label: "Bulletins d'anomalies" },
        ],
      },
      { type: 'leaf', key: 'reporting-agent', label: 'Reporting par Agent' },
    ]
    if (isController || isManager) {
      items.push({ type: 'leaf', key: 'reporting-saisie', label: 'Rapport quotidien' })
    }
    if (isManager) {
      items.push({ type: 'leaf', key: 'reporting-validation', label: 'Validation reportings' })
    }
    items.push(
      { type: 'leaf', key: 'plan-controle', label: 'Plan de Contrôle' },
      { type: 'leaf', key: 'plan-action-correctif', label: "Plan d'Action Correctif" },
    )
    return items
  }, [isManager, isController])

  const [activeTab, setActiveTab] = useState<Tab>('anomalie')
  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>({ anomalies: true })

  const toggleGroup = (key: string) => {
    setOpenGroups(prev => ({ ...prev, [key]: !prev[key] }))
  }

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
          {navItems.map(item => {
            if (item.type === 'leaf') {
              return (
                <button
                  key={item.key}
                  type="button"
                  className={`nav-item ${activeTab === item.key ? 'active' : ''}`}
                  onClick={() => setActiveTab(item.key)}
                  aria-current={activeTab === item.key ? 'page' : undefined}
                >
                  {item.label}
                </button>
              )
            }
            const isOpen = openGroups[item.key] ?? true
            const hasActiveChild = item.children.some(c => c.key === activeTab)
            return (
              <div key={item.key} className="nav-group">
                <button
                  type="button"
                  className={`nav-item nav-parent ${hasActiveChild ? 'active-parent' : ''}`}
                  onClick={() => toggleGroup(item.key)}
                  aria-expanded={isOpen}
                >
                  <span>{item.label}</span>
                  <span className="nav-chevron" aria-hidden="true">{isOpen ? '▾' : '▸'}</span>
                </button>
                {isOpen && (
                  <div className="nav-children" role="group" aria-label={item.label}>
                    {item.children.map(child => (
                      <button
                        key={child.key}
                        type="button"
                        className={`nav-item nav-child ${activeTab === child.key ? 'active' : ''}`}
                        onClick={() => setActiveTab(child.key)}
                        aria-current={activeTab === child.key ? 'page' : undefined}
                      >
                        {child.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            )
          })}
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
