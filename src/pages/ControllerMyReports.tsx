/**
 * ============================================================================
 * MODULE 3 (volet contrôleur) — MES RAPPORTS
 * ============================================================================
 *
 * Page accessible aux CONTRÔLEURS et MANAGERS (mais pensée d'abord pour
 * les contrôleurs).
 *
 * Rôle :
 *   Permet à un contrôleur de consulter l'historique de SES PROPRES rapports
 *   quotidiens sur une période donnée, sans avoir à passer par le manager.
 *
 *   Use case :
 *     « J'ai soumis un rapport il y a 3 jours, est-ce qu'il a déjà été
 *       validé par mon manager ? Et celui de la semaine dernière ? »
 *
 * Sécurité (filtrage applicatif) :
 *   - Le filtrage côté client se fait sur `controleurEmail === userEmail`
 *     (insensible à la casse)
 *   - C'est un simple confort UX — la sécurité réelle est gérée par
 *     SharePoint (un contrôleur ne devrait voir que ses items selon les
 *     ACLs de la liste, mais ici on charge TOUT puis on filtre côté front,
 *     donc cette page est techniquement consultable par un manager qui
 *     verrait alors uniquement SES propres rapports)
 *
 * Fonctionnalités :
 *   - Filtres :
 *       Période (date du / date au) — par défaut 30 derniers jours
 *       Statut (Soumis / Réalisé / Reporté / Tous)
 *       Recherche libre (dans observations + actions)
 *   - Pattern saisie/applied : les filtres ne se déclenchent qu'au clic
 *     sur "Rechercher" (cohérent avec les autres modules de l'app)
 *   - Stats cards : nombre de rapports / total heures / compteurs par statut
 *   - Tableau triable par défaut date desc (plus récent en haut)
 *   - Modale détail READ-ONLY (pas d'actions de validation, contrairement
 *     à ControllerReportingList.tsx qui est la vue manager) :
 *       Date, statut, totaux, statut journée
 *       Liste des actions de la journée
 *       Observations globales
 *       Pièces jointes (avec icônes)
 *       Historique des décisions managers (si présent)
 *
 * Données :
 *   - Source : listReports() depuis lib/activityService (qui lit SharePoint
 *     DCPO_ACTIVICTE_CONTROLLER)
 *   - Validation history : lue depuis localStorage (cf. activityService)
 *   - Pas de mutation depuis cette page → page lecture seule
 *
 * Réutilisation visuelle :
 *   - Réutilise les classes CSS de ControllerReporting.css (manager-*,
 *     stats-cards, filters-bar, etc.) pour cohérence visuelle avec le
 *     reste du module Reporting
 * ============================================================================
 */

import { useCallback, useEffect, useMemo, useState } from 'react'
import {
  listReports,
  type ActivityReport,
  type ActivityStatus,
} from '../lib/activityService'
import { getAttachmentIcon, getAttachmentIconType } from '../lib/ticketAttachments'
import './ControllerReporting.css'


/**
 * Props passées par Dashboard.tsx — identité du contrôleur courant.
 *   - userName  : DisplayName Office 365 (affichage)
 *   - userEmail : email Office 365 (clé de filtrage)
 */
interface ControllerMyReportsProps {
  userName?: string
  userEmail?: string
}


/**
 * Critères de filtrage de la page.
 *
 * Pourquoi 4 champs séparés (vs un objet "filter") :
 *   Lisibilité accrue dans le binding des inputs et des comparaisons.
 *
 * Le statut '' (chaîne vide) signifie "Tous" — on évite une valeur magique
 * type 'Tous' qui collisionnerait avec ActivityStatus.
 */
interface FilterState {
  dateFrom: string
  dateTo: string
  statut: '' | ActivityStatus
  search: string
}

/**
 * Construit l'état initial des filtres :
 *   - dateFrom : J-30 (les 30 derniers jours par défaut)
 *   - dateTo   : aujourd'hui
 *   - statut   : '' (tous)
 *   - search   : '' (pas de recherche)
 *
 * Cette plage par défaut est un compromis :
 *   - Assez large pour montrer du contenu dès l'ouverture
 *   - Assez restreinte pour ne pas charger des centaines de rapports
 *
 * Si on veut élargir : changer setDate(today.getDate() - N) avec N plus grand.
 */
function defaultFilters(): FilterState {
  const today = new Date()
  const monthAgo = new Date(today)
  monthAgo.setDate(today.getDate() - 30)
  return {
    dateFrom: monthAgo.toISOString().split('T')[0],
    dateTo: today.toISOString().split('T')[0],
    statut: '',
    search: '',
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * HELPERS DE FORMATAGE (locaux à la page — ne servent qu'ici)
 * ────────────────────────────────────────────────────────────────────────── */

/** "X.XX h" pour total heures (locale FR, max 2 décimales). */
function formatHours(value: number): string {
  return `${value.toLocaleString('fr-FR', { maximumFractionDigits: 2 })} h`
}

/** Date jj/mm/aaaa, "—" si vide / invalide. */
function formatDate(value?: string): string {
  if (!value) return '—'
  try { return new Date(value).toLocaleDateString('fr-FR') } catch { return value }
}

/** Date+heure pour les timestamps (submittedAt, validation history). */
function formatDateTime(value?: string): string {
  if (!value) return ''
  try { return new Date(value).toLocaleString('fr-FR') } catch { return value }
}

/**
 * Mappe un statut workflow vers la classe CSS de la pastille colorée.
 *   Soumis  → orange (en attente)
 *   Réalisé → vert (validé)
 *   Reporté → rouge (rejeté)
 */
function statutPillClass(statut: ActivityStatus): string {
  switch (statut) {
    case 'Soumis': return 'manager-pill-pending'
    case 'Réalisé': return 'manager-pill-realise'
    case 'Reporté': return 'manager-pill-reporte'
    default: return 'manager-pill'
  }
}


/* ══════════════════════════════════════════════════════════════════════════
 * COMPOSANT PRINCIPAL
 * ══════════════════════════════════════════════════════════════════════════ */

export default function ControllerMyReports({ userEmail }: ControllerMyReportsProps) {
  /**
   * Email normalisé en lowercase une seule fois ici.
   * Utilisé comme clé de filtrage côté client.
   * Si userEmail est undefined → chaîne vide → tous les rapports seront filtrés
   * out par défaut (cas anormal mais protégé contre crash).
   */
  const email = (userEmail ?? '').toLowerCase()


  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Liste de TOUS les rapports du contrôleur courant (pas filtrés UI).
   * Alimentée à chaque refresh().
   */
  const [reports, setReports] = useState<ActivityReport[]>([])

  /** Indicateur de chargement réseau (true au montage + à chaque actualisation). */
  const [loading, setLoading] = useState(true)

  /**
   * Filtres en cours de saisie (binding inputs).
   * Initialisés à defaultFilters() pour pré-remplir la période.
   */
  const [filters, setFilters] = useState<FilterState>(defaultFilters)

  /**
   * Filtres effectivement APPLIQUÉS au filtrage.
   * Snapshot pris au clic sur "Rechercher".
   * Pré-rempli aussi avec defaultFilters() pour que l'utilisateur voie déjà
   * les rapports des 30 derniers jours dès l'arrivée sur la page (sinon il
   * verrait une liste vide tant qu'il ne clique pas sur Rechercher).
   */
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(defaultFilters)

  /** Rapport sélectionné pour la modale détail (null = pas de modale). */
  const [selected, setSelected] = useState<ActivityReport | null>(null)


  /* ──────────────────────────────────────────────────────────────────────
   * CHARGEMENT DES DONNÉES
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Recharge la liste des rapports du contrôleur depuis SharePoint.
   *
   * Étapes :
   *   1. Court-circuit si pas d'email (cas anormal)
   *   2. listReports() retourne TOUS les rapports en SharePoint
   *   3. Filtrage côté client par controleurEmail === userEmail
   *   4. Stockage local
   *
   * useCallback : référence stable pour éviter de re-trigger l'effet
   *               qui dépend de cette fonction.
   */
  const refresh = useCallback(async () => {
    if (!email) {
      setReports([])
      setLoading(false)
      return
    }
    setLoading(true)
    try {
      const all = await listReports()
      // Filtrage : on ne garde QUE les rapports du contrôleur connecté.
      // Comparaison case-insensitive pour tolérer les variations
      // de casse possibles entre Office 365 et SharePoint.
      const mine = all.filter(r => (r.controleurEmail ?? '').toLowerCase() === email)
      setReports(mine)
    } catch (err) {
      console.error('Erreur chargement Mes rapports', err)
      setReports([])
    } finally {
      setLoading(false)
    }
  }, [email])

  /** Chargement initial au montage. */
  useEffect(() => {
    refresh()
  }, [refresh])


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS DE FILTRES
   * ────────────────────────────────────────────────────────────────────── */

  /** Met à jour partiellement un champ du filtre (binding générique). */
  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }

  /** Snapshot saisie → applied (déclenche la mise à jour de filteredReports). */
  const applyFilters = () => setAppliedFilters(filters)

  /**
   * Réinitialise les filtres à la valeur par défaut (30 derniers jours).
   * Met à jour les deux états (saisie ET applied) pour appliquer immédiatement.
   */
  const resetFilters = () => {
    const def = defaultFilters()
    setFilters(def)
    setAppliedFilters(def)
  }


  /* ──────────────────────────────────────────────────────────────────────
   * CALCULS DÉRIVÉS — PIPELINE reports → filteredReports → stats
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Applique les filtres APPLIQUÉS sur la liste des rapports du contrôleur.
   *
   * Pré-calculs :
   *   - Bornes de date converties UNE fois (Date) puis comparées en .getTime()
   *   - Terme de recherche normalisé en lowercase
   *
   * Ordre des tests (early return) :
   *   1. Statut (test trivial, élimine vite)
   *   2. Bornes de date (parsing + comparaison)
   *   3. Recherche libre (haystack construit à la demande)
   *
   * Recherche libre : cherche dans
   *   - observationsGlobales
   *   - chaque action / objet / observation des lignes
   * → permet de retrouver rapidement un rapport qui contient un mot-clé
   *   ("audit du 12 mai", "régularisation client X", etc.)
   */
  const filteredReports = useMemo(() => {
    const f = appliedFilters
    const fromDate = f.dateFrom ? new Date(`${f.dateFrom}T00:00:00`) : undefined
    const toDate = f.dateTo ? new Date(`${f.dateTo}T23:59:59`) : undefined
    const search = f.search.trim().toLowerCase()

    return reports.filter(r => {
      // Filtre statut
      if (f.statut && r.statut !== f.statut) return false

      // Filtre période (sur la date métier r.date, pas submittedAt)
      if (fromDate || toDate) {
        const rd = r.date ? new Date(r.date) : null
        if (!rd || Number.isNaN(rd.getTime())) return false
        if (fromDate && rd < fromDate) return false
        if (toDate && rd > toDate) return false
      }

      // Recherche libre — construit le haystack uniquement si nécessaire
      if (search) {
        const haystack = [
          r.observationsGlobales ?? '',
          ...(r.lines ?? []).flatMap(l => [l.objet, l.action, l.observation ?? '']),
        ].join(' ').toLowerCase()
        if (!haystack.includes(search)) return false
      }

      return true
    })
  }, [reports, appliedFilters])


  /**
   * Tri final : du plus récent au plus ancien (par date métier r.date).
   * Effectué APRÈS le filtre pour économiser des comparaisons sur les items
   * exclus.
   *
   * À date égale, on départage par submittedAt (le rapport soumis en dernier
   * remonte en premier — utile si un contrôleur soumet plusieurs rapports
   * pour la même date).
   */
  const sortedReports = useMemo(() => {
    return [...filteredReports].sort((a, b) => {
      if (a.date !== b.date) return a.date < b.date ? 1 : -1
      // tie-breaker : submittedAt desc
      return (a.submittedAt < b.submittedAt) ? 1 : -1
    })
  }, [filteredReports])


  /**
   * Statistiques agrégées pour les cards en haut de page.
   * Recalculées sur la liste filtrée (reflète les filtres actuels).
   */
  const stats = useMemo(() => {
    return {
      total: filteredReports.length,
      pending: filteredReports.filter(r => r.statut === 'Soumis').length,
      realise: filteredReports.filter(r => r.statut === 'Réalisé').length,
      reporte: filteredReports.filter(r => r.statut === 'Reporté').length,
      totalHeures: filteredReports.reduce((s, r) => s + (r.totalHeures || 0), 0),
    }
  }, [filteredReports])


  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   * ════════════════════════════════════════════════════════════════════════ */

  return (
    <>
      {/* ─── Header ────────────────────────────────────────────────── */}
      <div className="content-header">
        <h2>Mes rapports — historique d'activité</h2>
        <button className="btn-add" type="button" onClick={refresh} disabled={loading}>
          {loading ? 'Chargement...' : 'Actualiser'}
        </button>
      </div>

      {/* ─── Stats cards ───────────────────────────────────────────── */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Rapports</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.pending}</span>
          <span className="stat-label">En attente</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{stats.realise}</span>
          <span className="stat-label">Validés</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{stats.reporte}</span>
          <span className="stat-label">Reportés</span>
        </div>
        <div className="stat-card montant">
          <span className="stat-value">{formatHours(stats.totalHeures)}</span>
          <span className="stat-label">Total heures</span>
        </div>
      </div>

      {/* ─── Barre de filtres ──────────────────────────────────────── */}
      {/*
        Pattern saisie / applied : les inputs sont liés à `filters`,
        le filtrage utilise `appliedFilters`.
        Snapshot au clic Rechercher OU sur Entrée dans la recherche.
      */}
      <div className="filters-bar">
        <div className="filter-field">
          <label>Date du</label>
          <input
            type="date"
            value={filters.dateFrom}
            onChange={e => updateFilter('dateFrom', e.target.value)}
          />
        </div>
        <div className="filter-field">
          <label>Date au</label>
          <input
            type="date"
            value={filters.dateTo}
            onChange={e => updateFilter('dateTo', e.target.value)}
          />
        </div>
        <div className="filter-field">
          <label>Statut</label>
          <select
            value={filters.statut}
            onChange={e => updateFilter('statut', e.target.value as FilterState['statut'])}
          >
            <option value="">Tous</option>
            <option value="Soumis">Soumis (en attente)</option>
            <option value="Réalisé">Réalisé (validé)</option>
            <option value="Reporté">Reporté (rejeté)</option>
          </select>
        </div>
        <div className="filter-field" style={{ flex: 1, minWidth: 220 }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="Mot-clé dans actions / observations..."
            value={filters.search}
            onChange={e => updateFilter('search', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
          />
        </div>
        <button type="button" className="btn-search-filters" onClick={applyFilters} disabled={loading}>
          Rechercher
        </button>
        <button type="button" className="btn-reset-filters" onClick={resetFilters}>
          Réinitialiser
        </button>
      </div>

      {/* ─── Liste / vide / chargement ─────────────────────────────── */}
      {loading ? (
        <p className="loading-text">Chargement de vos rapports...</p>
      ) : sortedReports.length === 0 ? (
        <div className="manager-empty">
          <p>
            Aucun rapport ne correspond à votre sélection.
            {reports.length === 0 && ' Vous n\'avez encore soumis aucun rapport.'}
          </p>
        </div>
      ) : (
        /* Tableau dense : 1 ligne = 1 rapport. Click → modale détail. */
        <div className="manager-group">
          <table className="manager-table">
            <thead>
              <tr>
                <th>Date</th>
                <th>Total heures</th>
                <th>Temps occupé</th>
                <th>Statut journée</th>
                <th>Statut</th>
                <th>Actions</th>
                <th>Soumis le</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {sortedReports.map(r => (
                <tr key={r.id}>
                  <td><strong>{formatDate(r.date)}</strong></td>
                  <td>{formatHours(r.totalHeures)}</td>
                  <td>{r.tempsOccupe}%</td>
                  <td>{r.statutJournee}</td>
                  <td>
                    <span className={`manager-pill ${statutPillClass(r.statut)}`}>
                      {r.statut}
                    </span>
                  </td>
                  <td>{(r.lines ?? []).length}</td>
                  <td>{formatDateTime(r.submittedAt)}</td>
                  <td>
                    <button
                      type="button"
                      className="btn-detail"
                      onClick={() => setSelected(r)}
                    >
                      Détail
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* ─── Modale détail (read-only) ─────────────────────────────── */}
      {selected && (
        <MyReportDetailModal
          report={selected}
          onClose={() => setSelected(null)}
        />
      )}
    </>
  )
}


/* ══════════════════════════════════════════════════════════════════════════
 * MODALE DÉTAIL — lecture seule
 *
 * Différences avec ManagerDetailModal (ControllerReportingList.tsx) :
 *   - PAS de boutons Valider / Invalider (c'est la vue propriétaire)
 *   - PAS de saisie de motif
 *   - L'historique manager est affiché en lecture seule pour que le
 *     contrôleur sache pourquoi son rapport a été rejeté (motif)
 * ══════════════════════════════════════════════════════════════════════════ */

interface MyReportDetailModalProps {
  report: ActivityReport
  onClose: () => void
}

function MyReportDetailModal({ report, onClose }: MyReportDetailModalProps) {
  /**
   * Raccourci clavier : Escape ferme la modale.
   * Listener au montage, cleanup au démontage (fonction retournée).
   */
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal manager-modal"
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-labelledby="my-report-modal-title"
      >
        <div className="modal-header">
          <h2 id="my-report-modal-title">
            Rapport du {formatDate(report.date)}
          </h2>
          <button className="modal-close" onClick={onClose} aria-label="Fermer">
            &times;
          </button>
        </div>
        <div className="manager-modal-body">
          {/* ─── Sommaire — vue d'ensemble en grille ──────────────── */}
          {/*
            Affiche les indicateurs clés du rapport en bloc compact,
            permet à l'utilisateur de voir l'essentiel d'un coup d'œil.
          */}
          <dl className="manager-summary-grid">
            <div className="manager-summary-cell">
              <dt>Date</dt><dd>{formatDate(report.date)}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Total heures</dt><dd>{formatHours(report.totalHeures)}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Temps occupé</dt><dd>{report.tempsOccupe}%</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Statut journée</dt><dd>{report.statutJournee}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Anomalies détectées</dt><dd>{report.anomaliesDetectees ?? 0}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Statut</dt>
              <dd>
                <span className={`manager-pill ${statutPillClass(report.statut)}`}>
                  {report.statut}
                </span>
              </dd>
            </div>
          </dl>

          {/* ─── Tableau des actions de la journée ──────────────── */}
          {/*
            Réutilise le style manager-table pour cohérence visuelle
            avec la vue manager. Chaque ligne montre une activité
            individuelle reconstituée depuis le champ actionDeLaJournee.
          */}
          <section>
            <h3 style={{ margin: '0 0 8px', fontSize: 13, color: '#1a1a2e' }}>
              Actions de la journée
            </h3>
            {(!report.lines || report.lines.length === 0) ? (
              <p style={{ margin: 0, fontSize: 13, color: '#999', fontStyle: 'italic' }}>
                Aucune action consignée.
              </p>
            ) : (
              <table className="manager-table">
                <thead>
                  <tr>
                    <th>#</th>
                    <th>Domaine</th>
                    <th>Objet</th>
                    <th>Action</th>
                    <th>Durée</th>
                    <th>Observation</th>
                  </tr>
                </thead>
                <tbody>
                  {report.lines.map((line, i) => (
                    <tr key={line.id ?? i}>
                      <td>{i + 1}</td>
                      <td>{line.domaine || '—'}</td>
                      <td>{line.objet || '—'}</td>
                      <td>{line.action || '—'}</td>
                      <td>{line.duree} min</td>
                      <td>{line.observation || '—'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </section>

          {/* ─── Observations globales (si présentes) ───────────── */}
          {report.observationsGlobales && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: '#1a1a2e' }}>
                Observations globales
              </h3>
              <p style={{ margin: 0, fontSize: 13, whiteSpace: 'pre-wrap' }}>
                {report.observationsGlobales}
              </p>
            </section>
          )}

          {/* ─── Pièces jointes (si présentes) ──────────────────── */}
          {/*
            report.attachments est peuplé par reportFromItem() depuis le
            champ urlPieceJointes de SharePoint (cf. activityService).
            Chaque entrée a un { name, url } prêt à afficher.
          */}
          {report.attachments && report.attachments.length > 0 && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: '#1a1a2e' }}>
                Pièces jointes ({report.attachments.length})
              </h3>
              <ul style={{ listStyle: 'none', padding: 0, margin: 0, display: 'flex', flexDirection: 'column', gap: 6 }}>
                {report.attachments.map((a, i) => {
                  // Détecte le type d'icône selon l'extension du nom
                  // pour afficher l'emoji approprié (📄 pdf, 📊 excel, etc.)
                  const iconType = getAttachmentIconType(a.name || a.url)
                  return (
                    <li
                      key={i}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: 8,
                        padding: '6px 10px',
                        background: '#f8fafc',
                        border: '1px solid #eef2f7',
                        borderRadius: 6,
                        fontSize: 13,
                      }}
                    >
                      <span aria-hidden="true">{getAttachmentIcon(iconType)}</span>
                      {a.url ? (
                        <a
                          href={a.url}
                          target="_blank"
                          rel="noreferrer"
                          style={{ color: '#1d4ed8', textDecoration: 'none', fontWeight: 600, wordBreak: 'break-all' }}
                        >
                          {a.name || a.url}
                        </a>
                      ) : (
                        <span>{a.name}</span>
                      )}
                    </li>
                  )
                })}
              </ul>
            </section>
          )}

          {/* ─── Historique des décisions managers (si présent) ─── */}
          {/*
            Lecture seule pour le contrôleur — important pour qu'il
            comprenne pourquoi son rapport a été rejeté (motif).
            La couleur de la décision (vert/rouge) facilite la lecture rapide.
          */}
          {report.validationHistory && report.validationHistory.length > 0 && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: '#1a1a2e' }}>
                Décisions du manager
              </h3>
              <ul className="manager-history">
                {report.validationHistory.map((note, i) => (
                  <li key={i}>
                    <div>
                      <span
                        className={
                          'manager-history-decision ' +
                          (note.decision === 'Réalisé' ? 'realise' : 'reporte')
                        }
                      >
                        {note.decision}
                      </span>
                      {' — '}
                      <span>{note.manager}</span>
                      {' — '}
                      <span>{formatDateTime(note.date)}</span>
                    </div>
                    {note.motif && (
                      <div style={{ marginTop: 4, color: '#444' }}>
                        <em>Motif :</em> {note.motif}
                      </div>
                    )}
                  </li>
                ))}
              </ul>
            </section>
          )}

          {/* ─── Footer ─────────────────────────────────────────── */}
          {/*
            Pas de boutons d'action métier ici (read-only) — juste un
            bouton de fermeture pour cohérence UX.
          */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', borderTop: '1px solid #eee', paddingTop: 14 }}>
            <button type="button" className="btn-cta btn-cta-secondary" onClick={onClose}>
              Fermer
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
