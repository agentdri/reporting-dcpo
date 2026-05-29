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
 *   - "Rapport quotidien" (Saisie + Mes rapports) : visible UNIQUEMENT pour
 *     les Controleurs (les managers ne soumettent pas de rapport eux-mêmes)
 *   - "Validation reportings" : visible UNIQUEMENT pour les managers
 *     (Chef_Departement, Directeur) — leur permet de voir et valider les
 *     rapports soumis par les contrôleurs
 * ============================================================================
 */

import { useMemo, useState } from 'react'
import Anomalies from './Anomalies'
import AnomalyBulletins from './AnomalyBulletins'
import ReportingAgent from './ReportingAgent'
import ControllerReporting from './ControllerReporting'
import ControllerMyReports from './ControllerMyReports'
import ControllerReportingList from './ControllerReportingList'
import PlanControle from './PlanControle'
import PlanActionCorrectif from './PlanActionCorrectif'
import './Dashboard.css'

/**
 * Props passées par App.tsx après authentification réussie.
 */
interface DashboardProps {
  userName?: string  // displayName Office 365
  userRole?: string  // rôle EFFECTIF (peut être un rôle simulé) — pilote la nav/permissions
  userEmail?: string // mail Office 365 (utile pour les pages enfants)
  /**
   * Rôle RÉEL de l'utilisateur (DCPO_LISTE_USER.fonction). Sert uniquement à
   * déterminer quels rôles il peut simuler via le sélecteur de la topbar.
   */
  realUserRole?: string
  /**
   * Callback de changement de rôle effectif (simulation). Appelé par le
   * sélecteur de la topbar. Si absent, le sélecteur n'est pas affiché.
   */
  onRoleChange?: (role: string) => void
}

/**
 * Rôles que chaque rôle réel peut "assumer" (simuler), dans l'ordre
 * d'affichage du sélecteur. Le premier est toujours le rôle réel lui-même
 * (= revenir à son rôle normal).
 *
 * Règle métier :
 *   - Directeur        → Directeur, Chef_Departement, Controleur
 *   - Chef_Departement → Chef_Departement, Controleur
 *   - Controleur       → (absent du map → pas de sélecteur, pas de simulation)
 */
const ROLE_SWITCH_OPTIONS: Record<string, string[]> = {
  Directeur: ['Directeur', 'Chef_Departement', 'Controleur'],
  Chef_Departement: ['Chef_Departement', 'Controleur'],
}

/** Libellés lisibles pour l'affichage des rôles dans le sélecteur. */
const ROLE_LABELS: Record<string, string> = {
  Directeur: 'Directeur',
  Chef_Departement: 'Chef de département',
  Controleur: 'Contrôleur',
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
  | 'reporting-mes-rapports'
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

export default function Dashboard({ userName, userRole, userEmail, realUserRole, onRoleChange }: DashboardProps) {
  // ─── Calculs dérivés du rôle ─────────────────────────────────────────
  // Booléens pour conditionner l'affichage des onglets sensibles
  const isManager = !!userRole && MANAGER_ROLES.includes(userRole)
  const isController = userRole === 'Controleur'

  // Rôles que l'utilisateur peut simuler, en fonction de son rôle RÉEL.
  // Vide (et donc pas de sélecteur) si le rôle réel n'autorise pas la simulation
  // (ex: Controleur) ou si le callback de changement n'est pas fourni.
  const switchableRoles = (realUserRole && onRoleChange) ? (ROLE_SWITCH_OPTIONS[realUserRole] ?? []) : []

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
          { type: 'leaf', key: 'bulletins', label: "Fiche récapitulatif de l'anomalie" },
        ],
      },
      { type: 'leaf', key: 'reporting-agent', label: 'Reporting par Agent' },
    ]
    // Onglets conditionnels selon rôle
    //
    // Règle métier :
    //   - Un Controleur SOUMET des rapports et consulte SES propres rapports
    //     → groupe "Rapport quotidien" (Saisie + Mes rapports)
    //   - Un manager (Chef_Departement / Directeur) NE soumet PAS de rapport,
    //     il VALIDE ceux soumis par les contrôleurs → onglet "Validation reportings"
    //   - Les deux ensembles sont disjoints (un user ne peut pas être les deux)
    if (isController) {
      // Groupe "Rapport quotidien" : saisie + historique personnel (controleur uniquement)
      items.push({
        type: 'group',
        key: 'rapport-quotidien',
        label: 'Rapport quotidien',
        children: [
          { type: 'leaf', key: 'reporting-saisie', label: 'Saisie du rapport' },
          { type: 'leaf', key: 'reporting-mes-rapports', label: 'Mes rapports' },
        ],
      })
    }
    if (isManager) {
      // Validation des rapports soumis par les contrôleurs (manager uniquement)
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
   * Initialisé avec 'anomalies' et 'rapport-quotidien' ouverts pour que
   * l'utilisateur voie tout de suite les sous-menus disponibles.
   */
  const [openGroups, setOpenGroups] = useState<Record<string, boolean>>({
    anomalies: true,
    'rapport-quotidien': true,
  })

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
          {/* Sélecteur de rôle : affiché uniquement pour les rôles qui peuvent
              en simuler d'autres (Directeur / Chef_Departement). Sinon, simple
              libellé du rôle. */}
          {switchableRoles.length > 0 ? (
            <label className="topbar-role-switch">
              <span className="topbar-role-switch-icon" aria-hidden="true">👤</span>
              <select
                value={userRole ?? ''}
                onChange={e => onRoleChange?.(e.target.value)}
                aria-label="Changer de rôle (simulation)"
              >
                {switchableRoles.map(role => (
                  <option key={role} value={role}>
                    {ROLE_LABELS[role] ?? role}
                    {role === realUserRole ? ' (mon rôle)' : ''}
                  </option>
                ))}
              </select>
            </label>
          ) : (
            <span className="topbar-user-job">{ROLE_LABELS[userRole ?? ''] ?? userRole}</span>
          )}
          {/* Indicateur visuel quand l'utilisateur simule un rôle ≠ son rôle réel */}
          {realUserRole && userRole !== realUserRole && (
            <span className="topbar-role-simulated" title={`Rôle réel : ${ROLE_LABELS[realUserRole] ?? realUserRole}`}>
              rôle simulé
            </span>
          )}
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
            <Anomalies userName={userName} userEmail={userEmail} userRole={userRole} />
          )}
          {activeTab === 'bulletins' && <AnomalyBulletins />}
          {activeTab === 'reporting-agent' && <ReportingAgent />}
          {activeTab === 'reporting-saisie' && (
            <ControllerReporting userName={userName} userEmail={userEmail} />
          )}
          {activeTab === 'reporting-mes-rapports' && (
            <ControllerMyReports userName={userName} userEmail={userEmail} />
          )}
          {activeTab === 'reporting-validation' && (
            <ControllerReportingList userName={userName} userEmail={userEmail} />
          )}
          {activeTab === 'plan-controle' && (
            <PlanControle userName={userName} userEmail={userEmail} userRole={userRole} />
          )}
          {activeTab === 'plan-action-correctif' && (
            <PlanActionCorrectif userName={userName} userEmail={userEmail} userRole={userRole} />
          )}
        </div>
      </div>
    </div>
  )
}
