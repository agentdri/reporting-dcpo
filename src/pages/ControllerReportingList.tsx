/**
 * ============================================================================
 * MODULE 3 (volet manager) — VALIDATION DES RAPPORTS QUOTIDIENS
 * ============================================================================
 *
 * Page accessible UNIQUEMENT aux managers (Chef_Departement / Directeur).
 * Permet de valider ou invalider les rapports soumis par les contrôleurs.
 *
 * Fonctionnalités :
 *   - Liste consolidée des rapports, regroupés par (contrôleur, date)
 *   - Filtres : nom/email contrôleur, période, "en attente uniquement"
 *   - Stats : compteurs par statut (Soumis / Réalisé / Reporté), total heures
 *   - Modale détail : résumé du rapport + lignes + lien rapport + observations
 *     + pièces jointes + historique des décisions managers passées
 *   - Actions managers :
 *     - Valider → statut 'Réalisé', avec note horodatée optionnelle
 *     - Invalider → statut 'Reporté', motif obligatoire (saisi en textarea)
 *
 * Persistance :
 *   - Lecture : listReports() depuis activityService (localStorage)
 *   - Validation : validateReport() qui ajoute une ActivityValidationNote
 *     à validationHistory (append-only — audit complet préservé)
 *
 * Note future :
 *   Quand SharePoint ACTIVITE_Controleurs sera disponible, listReports et
 *   validateReport changeront d'implémentation côté lib mais cette page
 *   ne nécessitera pas de modification.
 * ============================================================================
 */

import { useCallback, useEffect, useMemo, useState } from 'react'
import {
  listReports,
  validateReport,
  type ActivityReport,
} from '../lib/activityService'
import './ControllerReporting.css'

/** Identité du manager courant (passée par Dashboard). */
interface ControllerReportingListProps {
  userName?: string
  userEmail?: string
}

/** Critères de filtrage de la liste des rapports. */
interface FilterState {
  controleur: string   // recherche partielle nom OU email
  dateFrom: string     // date min du rapport (YYYY-MM-DD)
  dateTo: string       // date max du rapport
  pendingOnly: boolean // ne montrer QUE les rapports en attente de validation
}

const EMPTY_FILTERS: FilterState = {
  controleur: '',
  dateFrom: '',
  dateTo: '',
  pendingOnly: false,
}

/**
 * Clé d'un groupe dans la vue regroupée : un groupe = (contrôleur, date).
 * Permet d'afficher tous les rapports d'un contrôleur pour un même jour
 * sous un même header (généralement il n'y en a qu'un, mais le manager peut
 * avoir reçu plusieurs envois).
 */
interface GroupKey {
  controleurEmail: string
  controleurName: string
  date: string
}

/**
 * Données agrégées d'un groupe : la liste des rapports + les compteurs
 * par statut + total heures cumulées.
 */
interface Group {
  key: GroupKey
  reports: ActivityReport[]
  totalHeures: number
  pending: number
  realise: number
  reporte: number
}

/** Formatte un nombre d'heures avec 2 décimales max (locale FR). */
function formatHours(value: number): string {
  return `${value.toLocaleString('fr-FR', { maximumFractionDigits: 2 })} h`
}

/** Date jj/mm/aaaa, "—" si vide ou invalide. */
function formatDate(value?: string): string {
  if (!value) return '—'
  try { return new Date(value).toLocaleDateString('fr-FR') } catch { return value }
}

/** Date+heure, '' si vide ou invalide (vide = pas de fallback affiché). */
function formatDateTime(value?: string): string {
  if (!value) return ''
  try { return new Date(value).toLocaleString('fr-FR') } catch { return value }
}

export default function ControllerReportingList({ userName, userEmail }: ControllerReportingListProps) {
  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS
   * ────────────────────────────────────────────────────────────────────── */

  /** Liste des rapports — chargée via useEffect (async SharePoint). */
  const [reports, setReports] = useState<ActivityReport[]>([])
  /** Indicateur de chargement initial (réseau en cours). */
  const [loadingReports, setLoadingReports] = useState(true)
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** Rapport sélectionné pour la modale détail (null = pas de modale). */
  const [selected, setSelected] = useState<ActivityReport | null>(null)

  /**
   * Recharge la liste depuis SharePoint.
   * useCallback : référence stable pour éviter les re-renders inutiles.
   * Async désormais (SharePoint fetch).
   */
  const refresh = useCallback(async () => {
    setLoadingReports(true)
    try {
      const data = await listReports()
      setReports(data)
    } finally {
      setLoadingReports(false)
    }
  }, [])

  // Chargement initial au montage
  useEffect(() => {
    refresh()
  }, [refresh])

  /**
   * Applique tous les filtres sur la liste brute.
   *
   * Ordre des tests (early return false) :
   *   1. pendingOnly (booléen le moins coûteux)
   *   2. controleur (recherche texte)
   *   3. dates (parsing + comparaison)
   *
   * useMemo : recalcul uniquement si reports OU filters change.
   */
  const filtered = useMemo(() => {
    const fromDate = filters.dateFrom ? new Date(`${filters.dateFrom}T00:00:00`) : undefined
    const toDate = filters.dateTo ? new Date(`${filters.dateTo}T23:59:59`) : undefined
    const search = filters.controleur.trim().toLowerCase()
    return reports.filter(r => {
      if (filters.pendingOnly && r.statut !== 'Soumis') return false
      if (search) {
        const haystack = `${r.controleurEmail} ${r.controleurName}`.toLowerCase()
        if (!haystack.includes(search)) return false
      }
      if (fromDate || toDate) {
        const rd = new Date(r.date)
        if (Number.isNaN(rd.getTime())) return false
        if (fromDate && rd < fromDate) return false
        if (toDate && rd > toDate) return false
      }
      return true
    })
  }, [reports, filters])

  /**
   * Regroupe les rapports filtrés par (contrôleur, date).
   *
   * Algorithme :
   *   1. Map<key, Group> où key = `${email}|${date}`
   *   2. Pour chaque rapport : créer le groupe si nécessaire, accumuler
   *   3. Trier les groupes :
   *      - D'abord par date desc (plus récent en haut)
   *      - Puis par nom de contrôleur (alpha)
   */
  const groups = useMemo<Group[]>(() => {
    const map = new Map<string, Group>()
    for (const r of filtered) {
      const k = `${r.controleurEmail}|${r.date}`
      if (!map.has(k)) {
        map.set(k, {
          key: { controleurEmail: r.controleurEmail, controleurName: r.controleurName, date: r.date },
          reports: [],
          totalHeures: 0,
          pending: 0,
          realise: 0,
          reporte: 0,
        })
      }
      const g = map.get(k)!
      g.reports.push(r)
      g.totalHeures += r.totalHeures
      if (r.statut === 'Soumis') g.pending++
      else if (r.statut === 'Réalisé') g.realise++
      else if (r.statut === 'Reporté') g.reporte++
    }
    return Array.from(map.values()).sort((a, b) => {
      if (a.key.date !== b.key.date) return a.key.date < b.key.date ? 1 : -1
      return a.key.controleurName.localeCompare(b.key.controleurName)
    })
  }, [filtered])

  /**
   * Stats globales pour les cards en haut de page.
   * Recalculées sur la liste filtrée (reflète les filtres en cours).
   */
  const totals = useMemo(() => {
    return {
      total: filtered.length,
      pending: filtered.filter(r => r.statut === 'Soumis').length,
      realise: filtered.filter(r => r.statut === 'Réalisé').length,
      reporte: filtered.filter(r => r.statut === 'Reporté').length,
      totalHeures: filtered.reduce((s, r) => s + r.totalHeures, 0),
    }
  }, [filtered])

  /** Helper générique de mise à jour partielle des filtres. */
  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const resetFilters = () => setFilters(EMPTY_FILTERS)

  /**
   * Applique une décision manager (Valider ou Invalider) à un rapport.
   *
   * Règle métier : motif OBLIGATOIRE pour invalidation (Reporté).
   * Garde-fou défensif : alert + abort si motif vide.
   *
   * Après validation :
   *   - refresh() recharge la liste depuis localStorage
   *   - setSelected(updated) garde la modale ouverte avec les données fraîches
   *     (l'utilisateur voit immédiatement la nouvelle entrée dans l'historique)
   */
  const handleValidation = async (report: ActivityReport, decision: 'Réalisé' | 'Reporté', motif?: string) => {
    if (decision === 'Reporté' && !motif?.trim()) {
      alert('Un motif est requis pour invalider un reporting.')
      return
    }
    const updated = await validateReport(report.id, {
      manager: userName ?? 'Manager',
      managerEmail: userEmail,
      decision,
      motif: motif?.trim() || undefined,
    })
    if (updated) {
      await refresh()
      setSelected(updated)
    }
  }

  return (
    <>
      <div className="content-header">
        <h2>Reportings soumis — Validation manager</h2>
        <button className="btn-add" type="button" onClick={refresh} disabled={loadingReports}>
          {loadingReports ? 'Chargement...' : 'Actualiser'}
        </button>
      </div>

      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{totals.total}</span>
          <span className="stat-label">Reportings</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{totals.pending}</span>
          <span className="stat-label">En attente</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{totals.realise}</span>
          <span className="stat-label">Réalisé</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{totals.reporte}</span>
          <span className="stat-label">Reporté</span>
        </div>
        <div className="stat-card montant">
          <span className="stat-value">{formatHours(totals.totalHeures)}</span>
          <span className="stat-label">Total heures</span>
        </div>
      </div>

      <div className="manager-controls">
        <div className="filter-field">
          <label>Contrôleur</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filters.controleur}
            onChange={e => updateFilter('controleur', e.target.value)}
          />
        </div>
        <div className="filter-field">
          <label>Date du</label>
          <input type="date" value={filters.dateFrom} onChange={e => updateFilter('dateFrom', e.target.value)} />
        </div>
        <div className="filter-field">
          <label>Date au</label>
          <input type="date" value={filters.dateTo} onChange={e => updateFilter('dateTo', e.target.value)} />
        </div>
        <label className="filter-toggle">
          <input
            type="checkbox"
            checked={filters.pendingOnly}
            onChange={e => updateFilter('pendingOnly', e.target.checked)}
          />
          En attente uniquement
        </label>
        <button type="button" className="btn-reset-filters" onClick={resetFilters}>Réinitialiser</button>
      </div>

      {loadingReports ? (
        <p className="loading-text">Chargement des rapports...</p>
      ) : groups.length === 0 ? (
        <div className="manager-empty">
          <p>Aucun reporting ne correspond aux critères.</p>
        </div>
      ) : (
        groups.map(group => (
          <section key={`${group.key.controleurEmail}-${group.key.date}`} className="manager-group">
            <header className="manager-group-header">
              <div className="manager-group-title">
                <strong>{group.key.controleurName}</strong>
                <span>{group.key.controleurEmail} — {formatDate(group.key.date)}</span>
              </div>
              <div className="manager-group-stats">
                <span className="manager-pill">{group.reports.length} entrée(s)</span>
                <span className="manager-pill">{formatHours(group.totalHeures)}</span>
                {group.pending > 0 && <span className="manager-pill manager-pill-pending">{group.pending} en attente</span>}
                {group.realise > 0 && <span className="manager-pill manager-pill-realise">{group.realise} validé(s)</span>}
                {group.reporte > 0 && <span className="manager-pill manager-pill-reporte">{group.reporte} reporté(s)</span>}
              </div>
            </header>
            <table className="manager-table">
              <thead>
                <tr>
                  <th>Soumis</th>
                  <th>Total heures</th>
                  <th>Temps occupé</th>
                  <th>Statut journée</th>
                  <th>Statut</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {group.reports.map(r => (
                  <tr key={r.id}>
                    <td>{formatDateTime(r.submittedAt)}</td>
                    <td>{formatHours(r.totalHeures)}</td>
                    <td>{r.tempsOccupe}%</td>
                    <td>{r.statutJournee}</td>
                    <td>
                      <span className={`manager-pill ${
                        r.statut === 'Soumis' ? 'manager-pill-pending'
                          : r.statut === 'Réalisé' ? 'manager-pill-realise'
                            : 'manager-pill-reporte'
                      }`}>{r.statut}</span>
                    </td>
                    <td>
                      <div className="manager-actions-cell">
                        <button className="btn-detail" type="button" onClick={() => setSelected(r)}>Détail</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </section>
        ))
      )}

      {selected && (
        <ManagerDetailModal
          report={selected}
          onClose={() => setSelected(null)}
          onValidate={(decision, motif) => handleValidation(selected, decision, motif)}
        />
      )}
    </>
  )
}

/** Props de la modale détail manager. */
interface ManagerDetailModalProps {
  report: ActivityReport
  onClose: () => void
  onValidate: (decision: 'Réalisé' | 'Reporté', motif?: string) => void
}

/**
 * Modale de détail d'un rapport pour le manager.
 *
 * Sections :
 *   1. Header : nom du contrôleur + bouton fermer
 *   2. Summary grid : date, total h, % occupé, statut journée, anomalies, statut
 *   3. Tableau des actions réalisées (lignes du rapport)
 *   4. Lien rapport externe (si présent)
 *   5. Observations globales (si présentes)
 *   6. Pièces jointes (si présentes)
 *   7. Historique des décisions managers (si non vide)
 *   8. Zone de décision : textarea motif + boutons Valider / Invalider
 *
 * Raccourci : Escape ferme la modale.
 *
 * Logique des boutons :
 *   - Valider : actif sauf si déjà 'Réalisé'
 *   - Invalider : actif sauf si déjà 'Reporté' OU motif vide
 */
function ManagerDetailModal({ report, onClose, onValidate }: ManagerDetailModalProps) {
  const [motif, setMotif] = useState('')

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
        aria-labelledby="manager-modal-title"
      >
        <div className="modal-header">
          <h2 id="manager-modal-title">Reporting — {report.controleurName}</h2>
          <button className="modal-close" onClick={onClose} aria-label="Fermer">&times;</button>
        </div>
        <div className="manager-modal-body">
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
              <dt>Statut</dt><dd>{report.statut}</dd>
            </div>
          </dl>

          <section>
            <h3 style={{ margin: '0 0 8px', fontSize: 13, color: '#1a1a2e' }}>Actions réalisées</h3>
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
                  <tr key={line.id}>
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
          </section>

          {report.lienRapport && (
            <p style={{ margin: 0, fontSize: 13 }}>
              <strong>Lien rapport : </strong>
              <a href={report.lienRapport} target="_blank" rel="noreferrer">{report.lienRapport}</a>
            </p>
          )}

          {report.observationsGlobales && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: '#1a1a2e' }}>Observations globales</h3>
              <p style={{ margin: 0, fontSize: 13, whiteSpace: 'pre-wrap' }}>{report.observationsGlobales}</p>
            </section>
          )}

          {report.attachments && report.attachments.length > 0 && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: '#1a1a2e' }}>Pièces jointes</h3>
              <ul className="bulletin-attachments" style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                {report.attachments.map((a, i) => (
                  <li key={i} style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '6px 0' }}>
                    <span aria-hidden="true">📎</span>
                    {a.url ? <a href={a.url} target="_blank" rel="noreferrer">{a.name}</a> : <span>{a.name}</span>}
                  </li>
                ))}
              </ul>
            </section>
          )}

          {report.validationHistory.length > 0 && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: '#1a1a2e' }}>Historique des décisions</h3>
              <ul className="manager-history">
                {report.validationHistory.map((note, i) => (
                  <li key={i}>
                    <div>
                      <span className={`manager-history-decision ${note.decision === 'Réalisé' ? 'realise' : 'reporte'}`}>
                        {note.decision}
                      </span>
                      {' — '}
                      <span>{note.manager}</span>
                      {' — '}
                      <span>{formatDateTime(note.date)}</span>
                    </div>
                    {note.motif && <div style={{ marginTop: 4, color: '#444' }}><em>Motif :</em> {note.motif}</div>}
                  </li>
                ))}
              </ul>
            </section>
          )}

          <div className="manager-decision-row">
            <label htmlFor="manager-motif" style={{ fontSize: 12, fontWeight: 700, color: '#666' }}>
              Motif (obligatoire pour invalider)
            </label>
            <textarea
              id="manager-motif"
              rows={3}
              placeholder="Saisir le motif si invalidation..."
              value={motif}
              onChange={e => setMotif(e.target.value)}
            />
            <div className="manager-decision-actions">
              <button
                type="button"
                className="btn-validate"
                onClick={() => onValidate('Réalisé')}
                disabled={report.statut === 'Réalisé'}
              >
                ✓ Valider (Réalisé)
              </button>
              <button
                type="button"
                className="btn-invalidate"
                onClick={() => onValidate('Reporté', motif)}
                disabled={report.statut === 'Reporté' || !motif.trim()}
              >
                ✗ Invalider (Reporté)
              </button>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
