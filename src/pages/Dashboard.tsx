/**
 * ============================================================================
 * DASHBOARD — INTERFACE PRINCIPALE DE L'APPLICATION
 * ============================================================================
 *
 * Affiché uniquement pour les utilisateurs autorisés (cf. App.tsx).
 *
 * Structure visuelle :
 *
 *   ┌──────────────────────────────────────────────────────────────────┐
 *   │ TOPBAR : Logo ReportingDCPO  ........  Nom utilisateur + rôle    │
 *   ├──────────┬───────────────────────────────────────────────────────┤
 *   │          │                                                       │
 *   │ SIDEBAR  │              CONTENU (page courante)                  │
 *   │          │                                                       │
 *   │ - Anomalies (groupe expansible)                                  │
 *   │   - Liste des anomalies                                          │
 *   │   - Bulletins d'anomalies                                        │
 *   │ - Reporting par Agent                                            │
 *   │ - Rapport quotidien (controllers + managers)                     │
 *   │ - Validation reportings (managers seulement)                     │
 *   │ - Plan de Contrôle                                               │
 *   │ - Plan d'Action Correctif                                        │
 *   │          │                                                       │
 *   └──────────┴───────────────────────────────────────────────────────┘
 *
 * Logique de routage :
 *   - State activeTab : key de l'onglet actif
 *   - Le rendu du contenu utilise des "&&" plutôt qu'un router → simple,
 *     léger, suffisant pour 7 onglets
 *
 * Logique des permissions :
 *   - Tous les utilisateurs autorisés voient les onglets standards
 *   - "Rapport quotidien" : visible pour Controleur OU manager
 *   - "Validation reportings" : visible UNIQUEMENT pour les managers
 * ============================================================================
 */

import { useMemo, useState } from 'react'
import Anomalies from './Anomalies'
import AnomalyBulletins from './AnomalyBulletins'
import ReportingAgent from './ReportingAgent'
import ControllerReporting from './ControllerReporting'
import ControllerReportingList from './ControllerReportingList'
import PlanControle from './PlanControle'
import PlanActionCorrectif from './PlanActionCorrectif'
import './Dashboard.css'

/**
 * Props passées par App.tsx après authentification réussie.
 */
interface DashboardProps {
  userName?: string  // displayName Office 365
  userRole?: string  // fonction.Value depuis DCPO_LISTE_USER
  userEmail?: string // mail Office 365 (utile pour les pages enfants)
}

/**
 * Énumération de tous les onglets possibles.
 *
 * Chaque key correspond à :
 *   - un item de navigation (nav-item)
 *   - une condition de rendu dans la zone .dashboard-content
 *
 * Pour ajouter un nouvel onglet : étendre ce type, l'ajouter à navItems,
 * et ajouter la condition de rendu en bas du composant.
 */
type Tab =
  | 'anomalie'
  | 'bulletins'
  | 'reporting-agent'
  | 'reporting-saisie'
  | 'reporting-validation'
  | 'plan-controle'
  | 'plan-action-correctif'

/**
 * Item de nav simple (feuille de l'arbre).
 *   - type discriminant 'leaf' (vs 'group') pour le pattern matching
 *   - key : valeur de Tab à activer au clic
 *   - label : texte affiché
 */
interface NavLeaf {
  type: 'leaf'
  key: Tab
  label: string
}

/**
 * Groupe de nav (parent expansible avec enfants).
 *   - key : identifiant string (NON une Tab car le groupe lui-même
 *     n'est pas un onglet ; c'est juste un bouton expand/collapse)
 *   - children : tableau de NavLeaf imbriqués
 */
interface NavGroup {
  type: 'group'
  key: string
  label: string
  children: NavLeaf[]
}

/** Union pour le tableau de nav (mélange feuilles et groupes). */
type NavEntry = NavLeaf | NavGroup

/**
 * Liste des rôles considérés comme "manager".
 * Ces rôles ont accès à des fonctionnalités supplémentaires
 * (validation des reportings).
 */
const MANAGER_ROLES = ['Chef_Departement', 'Directeur']

export default function Dashboard({ userName, userRole, userEmail }: DashboardProps) {
  // ─── Calculs dérivés du rôle ─────────────────────────────────────────
  // Booléens pour conditionner l'affichage des onglets sensibles
  const isManager = !!userRole && MANAGER_ROLES.includes(userRole)
  const isController = userRole === 'Controleur'

  /**
   * Construction dynamique de la nav selon le rôle.
   *
   * useMemo pour éviter de reconstruire le tableau à chaque render
   * (les dépendances [isManager, isController] ne changent jamais
   * pendant la session, mais on garde le useMemo pour la pureté).
   */
  const navItems = useMemo<NavEntry[]>(() => {
    const items: NavEntry[] = [
      // Groupe Anomalies : 2 sous-onglets toujours visibles
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
    // Onglets conditionnels selon rôle
    if (isController || isManager) {
      items.push({ type: 'leaf', key: 'reporting-saisie', label: 'Rapport quotidien' })
    }
    if (isManager) {
      items.push({ type: 'leaf', key: 'reporting-validation', label: 'Validation reportings' })
    }
    // Onglets toujours en queue
    items.push(
      { type: 'leaf', key: 'plan-controle', label: 'Plan de Contrôle' },
      { type: 'leaf', key: 'plan-action-correctif', label: "Plan d'Action Correctif" },
    )
    return items
  }, [isManager, isController])

  // ─── États ─────────────────────────────────────────────────────────────

  /** Onglet courant. 'anomalie' (Liste) par défaut au montage. */
  const [activeTab, setActiveTab] = useState<Tab>('anomalie')

  /**
   * État d'expansion des groupes de navigation.
   * Map keyName → true (ouvert) / false (fermé).
   * Initialisé avec 'anomalies' ouvert pour que l'utilisateur voie tout
   * de suite le sous-menu (cohérent avec activeTab par défaut).
   */
  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>({ anomalies: true })

  /** Toggle l'état d'un groupe (expansion / contraction). */
  const toggleGroup = (key: string) => {
    setOpenGroups(prev => ({ ...prev, [key]: !prev[key] }))
  }

  return (
    <div className="dashboard">
      {/* ─── TOPBAR : titre app + identité utilisateur ─────────────── */}
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
        {/* ─── SIDEBAR : navigation principale ─────────────────────── */}
        <nav className="dashboard-nav" aria-label="Navigation principale">
          {navItems.map(item => {
            // Branche feuille : un simple bouton qui change activeTab
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
            // Branche groupe : parent expansible + enfants conditionnels
            const isOpen = openGroups[item.key] ?? true
            // hasActiveChild : style distinct pour le parent quand un enfant est actif
            const hasActiveChild = item.children.some(c => c.key === activeTab)
            return (
              <div key={item.key} className="nav-group">
                {/* Bouton parent : toggle expand/collapse */}
                <button
                  type="button"
                  className={`nav-item nav-parent ${hasActiveChild ? 'active-parent' : ''}`}
                  onClick={() => toggleGroup(item.key)}
                  aria-expanded={isOpen}
                >
                  <span>{item.label}</span>
                  <span className="nav-chevron" aria-hidden="true">{isOpen ? '▾' : '▸'}</span>
                </button>
                {/* Enfants visibles uniquement si groupe ouvert */}
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

        {/* ─── CONTENU : composant correspondant à activeTab ─────── */}
        {/* Pattern simple "&& render" : un seul des conditions matche */}
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
