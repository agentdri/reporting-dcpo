/**
 * ============================================================================
 * DASHBOARD — INTERFACE PRINCIPALE DE L'APPLICATION
 * ============================================================================
 *
 * Affiché uniquement pour les utilisateurs autorisés (cf. App.tsx).
 *
 * Structure visuelle :
 *
 *   ┌────────────────────────────────────────────────────────────────────┐
 *   │ TOPBAR : Logo  .....  Nom user + [sélecteur de rôle] + badge       │
 *   ├──────────┬─────────────────────────────────────────────────────────┤
 *   │          │                                                         │
 *   │ SIDEBAR  │              CONTENU (page courante)                    │
 *   │          │                                                         │
 *   │ - Anomalies (groupe expansible)                                    │
 *   │   - Liste des anomalies                                            │
 *   │   - Fiche récapitulatif de l'anomalie                              │
 *   │ - Reporting par Agent                                              │
 *   │ - Rapport quotidien (controleurs seulement) (groupe)               │
 *   │   - Saisie du rapport                                              │
 *   │   - Mes rapports                                                   │
 *   │ - Validation reportings (managers seulement)                       │
 *   │ - Plan de Contrôle                                                 │
 *   │ - Plan d'Action Correctif                                          │
 *   │ - Configuration (managers seulement) (groupe)                      │
 *   │   - Directions                                                     │
 *   │   - Gestion des utilisateurs                                       │
 *   │          │                                                         │
 *   └──────────┴─────────────────────────────────────────────────────────┘
 *
 * ROUTAGE
 * -------
 *   - State activeTab : key de l'onglet actif
 *   - Le rendu du contenu utilise des "&&" plutôt qu'un router (suffisant
 *     pour ~10 onglets, pas d'URL profonde)
 *   - Les groupes (Anomalies, Rapport quotidien, Configuration) sont des
 *     parents EXPANSIBLES (toggle) — leur clé n'est PAS un Tab.
 *
 * PERMISSIONS (gouvernées par MANAGER_ROLES)
 * ------------------------------------------
 *   - Tous les autorisés : Anomalies, Reporting par Agent, Plan de Contrôle,
 *     Plan d'Action Correctif
 *   - Controleurs UNIQUEMENT : groupe "Rapport quotidien" (ils soumettent
 *     leur reporting journalier — les managers eux ne le soumettent pas)
 *   - Managers UNIQUEMENT : "Validation reportings" + groupe "Configuration"
 *     (Directions, Gestion des utilisateurs)
 *
 * SIMULATION DE RÔLE (topbar)
 * ---------------------------
 * Si le rôle RÉEL de l'utilisateur (realUserRole) est Chef_Departement ou
 * Directeur, la topbar expose un sélecteur (ROLE_SWITCH_OPTIONS) qui permet
 * de "rétrograder" temporairement le rôle effectif :
 *   - Directeur        → peut simuler Chef_Departement ou Controleur
 *   - Chef_Departement → peut simuler Controleur
 *   - Controleur       → pas de simulation (déjà le rôle le plus bas)
 *
 * C'est une simulation CÔTÉ CLIENT uniquement (le rôle réel reste inchangé
 * en base). Permet de vérifier l'UX d'un autre rôle, de faire des démos ou
 * du support. Un badge "rôle simulé" s'affiche tant que actif ≠ réel.
 *
 * AJOUTER UN NOUVEL ONGLET — Marche à suivre
 * ------------------------------------------
 *   1. Étendre le type Tab en haut de ce fichier
 *   2. Ajouter la `leaf` (ou modifier un groupe) dans navItems
 *   3. Ajouter le rendu conditionnel `{activeTab === 'xxx' && <Page ... />}`
 *      en bas du composant
 *   4. Importer le composant en haut
 *   5. Si la page est restreinte à un rôle : conditionner avec isManager /
 *      isController dans navItems
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
import UserManagement from './UserManagement'
import DirectionManagement from './DirectionManagement'
import DashboardPowerBI from './DashboardPowerBI'

/**
 * URLs publiques (Publier sur le web) des dashboards Power BI.
 *
 * Centralisées ici pour éviter d'avoir à chercher dans le code à chaque
 * mise à jour. Pour publier un nouveau rapport : Power BI Desktop →
 * Publier → Power BI Service → Fichier → Publier sur le web → copier l'URL.
 */
const POWER_BI_DASHBOARDS = {
  numerisation: 'https://app.powerbi.com/view?r=eyJrIjoiY2MyMzA2NjctMDllYS00MWY3LTkxM2ItZTg2MDNiNWM1ZWM3IiwidCI6IjJiZDgyYTY4LTJjN2QtNGM0My1iMDgwLTliMDY0NDEwZjZjZiJ9',
  creance: 'https://app.powerbi.com/view?r=eyJrIjoiNDFjYjRjZjItOWNmOC00YzE1LTg5MzYtZjljNmZkNDZmNTcyIiwidCI6IjJiZDgyYTY4LTJjN2QtNGM0My1iMDgwLTliMDY0NDEwZjZjZiJ9',
  comex: 'https://app.powerbi.com/view?r=eyJrIjoiZTAwMjVkZTItOTI3YS00MmU2LTg5OTktZTY1YWExMjRlYzdhIiwidCI6IjJiZDgyYTY4LTJjN2QtNGM0My1iMDgwLTliMDY0NDEwZjZjZiJ9',
}
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
  | 'user-management'
  | 'config-directions'      // Groupe Configuration → Directions (référentiel)
  | 'dashboard-numerisation' // Groupe Dashboards Power BI → Numérisation
  | 'dashboard-creance'      // Groupe Dashboards Power BI → Créance Hors Bilan
  | 'dashboard-comex'        // Groupe Dashboards Power BI → Apurement Comex

/**
 * Item de nav simple (feuille de l'arbre).
 *   - type discriminant 'leaf' (vs 'group') pour le pattern matching
 *   - key : identifiant unique. Pour les liens INTERNES, c'est une valeur
 *     de Tab activée au clic. Pour les liens EXTERNES (`href` défini), c'est
 *     juste un identifiant React (typage élargi à string)
 *   - label : texte affiché
 *   - icon : (optionnel) emoji ou caractère unicode affiché à gauche du label
 *     pour faciliter la lecture visuelle de la sidebar. Pas d'asset à charger,
 *     compatible multiplateforme.
 *   - href : (optionnel) URL externe. Si défini, l'item est rendu comme un
 *     `<a target="_blank">` qui ouvre l'URL dans un nouvel onglet sans
 *     changer l'onglet interne actif. Sert à intégrer des applications
 *     externes au dashboard (ex: autres apps Power Apps métier).
 */
interface NavLeaf {
  type: 'leaf'
  key: Tab | string
  label: string
  icon?: string
  href?: string
}

/**
 * Groupe de nav (parent expansible avec enfants).
 *   - key : identifiant string (NON une Tab car le groupe lui-même
 *     n'est pas un onglet ; c'est juste un bouton expand/collapse)
 *   - children : tableau de NavLeaf imbriqués
 *   - icon : (optionnel) emoji affiché à gauche du label du groupe
 */
interface NavGroup {
  type: 'group'
  key: string
  label: string
  children: NavLeaf[]
  icon?: string
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
        icon: '⚠️',
        children: [
          { type: 'leaf', key: 'anomalie', label: 'Anomalies en cours', icon: '📋' },
          { type: 'leaf', key: 'bulletins', label: 'Anomalies résolues', icon: '📄' },
        ],
      },
      { type: 'leaf', key: 'reporting-agent', label: 'Reporting par Agent', icon: '👤' },
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
        icon: '📅',
        children: [
          { type: 'leaf', key: 'reporting-saisie', label: 'Saisie du rapport', icon: '📝' },
          { type: 'leaf', key: 'reporting-mes-rapports', label: 'Mes rapports', icon: '📂' },
        ],
      })
    }
    if (isManager) {
      // Validation des rapports soumis par les contrôleurs (manager uniquement)
      items.push({ type: 'leaf', key: 'reporting-validation', label: 'Validation reportings', icon: '✅' })
    }
    // Onglets toujours en queue
    items.push(
      { type: 'leaf', key: 'plan-controle', label: 'Plan de Contrôle', icon: '🎯' },
      { type: 'leaf', key: 'plan-action-correctif', label: "Plan d'Action Correctif", icon: '🛠️' },
    )
    // Groupe "Dashboards Power BI" — visible par tous les rôles autorisés.
    // Chaque sous-onglet rend un iframe pointant vers un rapport Power BI
    // publié (URL publique app.powerbi.com/view?r=...). Voir DashboardPowerBI.tsx
    // pour le composant générique.
    //
    // Inclut aussi un lien externe vers l'application Power Apps "Clôture des
    // journées" qui ouvre dans un nouvel onglet (cf. champ `href`).
    items.push({
      type: 'group',
      key: 'dashboards',
      label: 'Dashboards',
      icon: '📊',
      children: [
        { type: 'leaf', key: 'dashboard-numerisation', label: 'Numérisation des journées', icon: '📑' },
        { type: 'leaf', key: 'dashboard-creance', label: 'Créance Hors Bilan', icon: '💰' },
        { type: 'leaf', key: 'dashboard-comex', label: 'Apurement Comex', icon: '📈' },
        {
          type: 'leaf',
          key: 'ext-cloture-journees',
          label: 'Clôture des journées',
          icon: '📒',
          // App Power Apps externe — ouvre dans un nouvel onglet
          // (cf. logique des leaves avec `href` dans le rendu de la sidebar).
          href: 'https://apps.powerapps.com/play/e/e78a17af-caf0-e888-989b-beca000173f8/a/8d342307-2ae1-4d55-ac08-414d26a60449',
        },
      ],
    })
    // Groupe "Configuration" : référentiels métier + gestion des accès.
    // Réservé aux managers (Chef_Departement / Directeur) — restriction
    // d'accès aux paramètres partagés pour éviter les modifications anarchiques.
    //
    // Pour ajouter un nouveau référentiel (ex: catégories, types d'anomalies) :
    //   1. Étendre le type Tab en haut de ce fichier
    //   2. Ajouter une `leaf` dans children: ci-dessous
    //   3. Ajouter le rendu conditionnel correspondant en bas du composant
    if (isManager) {
      items.push({
        type: 'group',
        key: 'configuration',
        label: 'Configuration',
        icon: '⚙️',
        children: [
          { type: 'leaf', key: 'config-directions', label: 'Directions', icon: '🏛️' },
          { type: 'leaf', key: 'user-management', label: 'Gestion des utilisateurs', icon: '👥' },
        ],
      })
    }
    return items
  }, [isManager, isController])

  // ─── États ─────────────────────────────────────────────────────────────

  /** Onglet courant. 'anomalie' (Liste) par défaut au montage. */
  const [activeTab, setActiveTab] = useState<Tab>('anomalie')

  /**
   * Date cible passée à <ControllerReporting> quand on navigue depuis la
   * page "Mes rapports" via le bouton "Modifier et re-soumettre" d'un
   * rapport rejeté.
   *
   *   - undefined → comportement par défaut (date = aujourd'hui)
   *   - 'YYYY-MM-DD' → la saisie pré-charge cette date, ce qui déclenche
   *     dans ControllerReporting le lookup serveur et le pré-remplissage
   *     du rapport rejeté correspondant.
   */
  const [targetReportDate, setTargetReportDate] = useState<string | undefined>(undefined)

  /**
   * Callback passé à ControllerMyReports : redirige vers la saisie en
   * ciblant la date d'un rapport rejeté à corriger.
   *
   * Flux :
   *   1. Utilisateur clique "Modifier" sur la ligne d'un rapport rejeté
   *   2. setTargetReportDate(date) mémorise la date à pré-charger
   *   3. setActiveTab('reporting-saisie') affiche la page de saisie
   *   4. ControllerReporting reçoit targetDate et lance son lookup
   *      → la bannière orange "Rapport rejeté" apparaît + formulaire pré-rempli
   */
  const handleEditRejectedReport = (date: string) => {
    setTargetReportDate(date)
    setActiveTab('reporting-saisie')
  }

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
          <h1 className="topbar-title">
            <span aria-hidden="true" style={{ marginRight: 8 }}>📋</span>
            ReportingDCPO
          </h1>
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
            // Branche feuille : un simple bouton qui change activeTab,
            // OU un lien externe (<a target="_blank">) si `href` est défini.
            if (item.type === 'leaf') {
              // Lien externe : on rend un <a> qui ouvre l'URL dans un nouvel
              // onglet. On NE TOUCHE PAS à activeTab → l'utilisateur revient
              // sur son onglet actuel quand il ferme le nouvel onglet.
              // L'icône "↗" en suffixe signale visuellement la nature externe.
              if (item.href) {
                return (
                  <a
                    key={item.key}
                    href={item.href}
                    target="_blank"
                    rel="noreferrer"
                    className="nav-item"
                    style={{ textDecoration: 'none' }}
                  >
                    {item.icon && <span className="nav-icon" aria-hidden="true">{item.icon}</span>}
                    {item.label}
                    <span aria-hidden="true" style={{ marginLeft: 6, fontSize: 11, opacity: 0.7 }}>↗</span>
                  </a>
                )
              }
              return (
                <button
                  key={item.key}
                  type="button"
                  className={`nav-item ${activeTab === item.key ? 'active' : ''}`}
                  onClick={() => setActiveTab(item.key as Tab)}
                  aria-current={activeTab === item.key ? 'page' : undefined}
                >
                  {/* Icône optionnelle : `aria-hidden` car redondante avec le label texte */}
                  {item.icon && <span className="nav-icon" aria-hidden="true">{item.icon}</span>}
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
                  <span>
                    {item.icon && <span className="nav-icon" aria-hidden="true">{item.icon}</span>}
                    {item.label}
                  </span>
                  <span className="nav-chevron" aria-hidden="true">{isOpen ? '▾' : '▸'}</span>
                </button>
                {/* Enfants visibles uniquement si groupe ouvert.
                    Comme pour les leaves de niveau racine, on supporte les
                    liens externes via la prop `href`. */}
                {isOpen && (
                  <div className="nav-children" role="group" aria-label={item.label}>
                    {item.children.map(child => {
                      if (child.href) {
                        return (
                          <a
                            key={child.key}
                            href={child.href}
                            target="_blank"
                            rel="noreferrer"
                            className="nav-item nav-child"
                            style={{ textDecoration: 'none' }}
                          >
                            {child.icon && <span className="nav-icon" aria-hidden="true">{child.icon}</span>}
                            {child.label}
                            <span aria-hidden="true" style={{ marginLeft: 6, fontSize: 11, opacity: 0.7 }}>↗</span>
                          </a>
                        )
                      }
                      return (
                      <button
                        key={child.key}
                        type="button"
                        className={`nav-item nav-child ${activeTab === child.key ? 'active' : ''}`}
                        onClick={() => setActiveTab(child.key as Tab)}
                        aria-current={activeTab === child.key ? 'page' : undefined}
                      >
                        {child.icon && <span className="nav-icon" aria-hidden="true">{child.icon}</span>}
                        {child.label}
                      </button>
                      )
                    })}
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
            <ControllerReporting
              userName={userName}
              userEmail={userEmail}
              targetDate={targetReportDate}
            />
          )}
          {activeTab === 'reporting-mes-rapports' && (
            <ControllerMyReports
              userName={userName}
              userEmail={userEmail}
              onEditRejected={handleEditRejectedReport}
            />
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
          {activeTab === 'user-management' && (
            <UserManagement userEmail={userEmail} userRole={userRole} />
          )}
          {activeTab === 'config-directions' && <DirectionManagement />}
          {/* ─── Dashboards Power BI ──────────────────────────────────
              Trois rapports embarqués via iframe — cf. composant générique
              DashboardPowerBI qui prend titre + URL en props. */}
          {activeTab === 'dashboard-numerisation' && (
            <DashboardPowerBI
              title="Suivi de numérisation des journées comptables"
              url={POWER_BI_DASHBOARDS.numerisation}
            />
          )}
          {activeTab === 'dashboard-creance' && (
            <DashboardPowerBI
              title="Surveillance Créance Hors Bilan"
              url={POWER_BI_DASHBOARDS.creance}
            />
          )}
          {activeTab === 'dashboard-comex' && (
            <DashboardPowerBI
              title="Apurement Comex"
              url={POWER_BI_DASHBOARDS.comex}
            />
          )}
        </div>
      </div>
    </div>
  )
}
