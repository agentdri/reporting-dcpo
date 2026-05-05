import { useState } from 'react'
import Anomalies from './Anomalies'
import ReportingAgent from './ReportingAgent'
import ActiviteControleur from './ActiviteControleur'
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
  | 'reporting-agent'
  | 'activite-controleur'
  | 'plan-controle'
  | 'plan-action-correctif'

export default function Dashboard({ userName, userRole, userEmail }: DashboardProps) {
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
        <nav className="dashboard-nav">
          <button
            className={`nav-item ${activeTab === 'anomalie' ? 'active' : ''}`}
            onClick={() => setActiveTab('anomalie')}
          >
            Anomalies
          </button>
          <button
            className={`nav-item ${activeTab === 'reporting-agent' ? 'active' : ''}`}
            onClick={() => setActiveTab('reporting-agent')}
          >
            Reporting par Agent
          </button>
          <button
            className={`nav-item ${activeTab === 'activite-controleur' ? 'active' : ''}`}
            onClick={() => setActiveTab('activite-controleur')}
          >
            Activite Controleur
          </button>
          <button
            className={`nav-item ${activeTab === 'plan-controle' ? 'active' : ''}`}
            onClick={() => setActiveTab('plan-controle')}
          >
            Plan de Controle
          </button>
          <button
            className={`nav-item ${activeTab === 'plan-action-correctif' ? 'active' : ''}`}
            onClick={() => setActiveTab('plan-action-correctif')}
          >
            Plan d'Action Correctif
          </button>
        </nav>

        <div className="dashboard-content">
          {activeTab === 'anomalie' && (
            <Anomalies userName={userName} userEmail={userEmail} />
          )}
          {activeTab === 'reporting-agent' && <ReportingAgent />}
          {activeTab === 'activite-controleur' && <ActiviteControleur />}
          {activeTab === 'plan-controle' && <PlanControle />}
          {activeTab === 'plan-action-correctif' && <PlanActionCorrectif />}
        </div>
      </div>
    </div>
  )
}
