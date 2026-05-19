/**
 * ============================================================================
 * MODULE — PLAN D'ACTION CORRECTIF (PAC)
 * ============================================================================
 *
 * Permet aux managers (Chef_Departement / Directeur) de créer et suivre les
 * plans d'action correctifs émis par la DCPO. Les autres utilisateurs ont
 * un accès en lecture seule (consultation de la liste et du détail).
 *
 * Structure :
 *   - Stats cards : compteurs Total / En cours / Exécutées / Non Exécutées
 *   - Barre de filtres : statut, direction, année + bouton "Nouveau PAC"
 *     (visible uniquement pour les managers)
 *   - Tableau paginé des PAC avec colonnes principales
 *   - Modale Détail (lecture seule)
 *   - Modale Création (formulaire complet)
 *
 * Persistance : actuellement localStorage via src/lib/pacService.ts. Le
 * service expose une API stable — l'utilisateur peut basculer vers les
 * listes SharePoint sans toucher à ce composant.
 *
 * Source des champs : feuille "PAC DCPO" du fichier Excel
 * Tableau_de_bord_KPI_DCPO_2026.xlsx fourni par l'utilisateur.
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import {
  createPAC,
  listDirections,
  listPACs,
  getDirectionLabel,
  getPacStatusClass,
  PAC_STATUS_OPTIONS,
  type Direction,
  type Pac,
  type PacStatus,
} from '../lib/pacService'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'


/**
 * Props passées par Dashboard.tsx.
 *
 * Seul userRole pilote la logique métier (canManage). Les autres props sont
 * conservées pour cohérence avec les autres modules (logging futur,
 * pré-remplissage responsable, etc.).
 */
interface PlanActionCorrectifProps {
  userName?: string
  userEmail?: string
  userRole?: string
}

/** Rôles ayant le droit de créer / éditer un PAC. */
const MANAGER_ROLES = ['Chef_Departement', 'Directeur']

/** État vide pour la barre de filtres (réutilisé au reset). */
const EMPTY_FILTERS = {
  statut: '' as '' | PacStatus,
  direction: '',
  annee: '',
  search: '',
}

type FilterState = typeof EMPTY_FILTERS

/** État vide pour le formulaire de création. */
const EMPTY_FORM = {
  sourcePac: '',
  dateCreation: new Date().toISOString().split('T')[0],
  intitule: '',
  descriptionProbleme: '',
  causeImmediate: '',
  causeRacine: '',
  actionsCorrectives: '',
  directionsConcernees: [] as string[],
  echeance: '',
  kpi: '',
  annee: new Date().getFullYear(),
  responsable: '',
  statut: 'En cours' as PacStatus,
  observations: '',
}

/**
 * Formatte une date YYYY-MM-DD vers locale FR (jj/mm/aaaa).
 * Tolérant : renvoie '—' si vide / invalide.
 */
function formatDate(d: string | undefined): string {
  if (!d) return '—'
  const parsed = new Date(d)
  if (Number.isNaN(parsed.getTime())) return d
  return parsed.toLocaleDateString('fr-FR')
}


export default function PlanActionCorrectif({ userRole }: PlanActionCorrectifProps) {
  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * Liste des PAC (rechargée à chaque mutation).
   * Lazy init via useState(() => ...) pour éviter le pattern "fetch in useEffect"
   * → satisfait la règle react-hooks/set-state-in-effect.
   */
  const [pacs, setPacs] = useState<Pac[]>(() => listPACs())
  /** Liste des directions actives (idem lazy init). */
  const [directions, setDirections] = useState<Direction[]>(() => listDirections(true))
  /** Filtres en cours de saisie (binding inputs). */
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** Filtres effectivement appliqués (snapshot au clic Rechercher). */
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** PAC sélectionné pour la modale de détail (null = fermée). */
  const [selected, setSelected] = useState<Pac | null>(null)
  /** Affichage du formulaire de création (toggle). */
  const [showForm, setShowForm] = useState(false)
  /** Snapshot des champs en cours de saisie dans le formulaire de création. */
  const [form, setForm] = useState(EMPTY_FORM)
  /** Erreur de soumission affichée sous le formulaire. */
  const [formError, setFormError] = useState<string | null>(null)

  /** Permission "créer / éditer un PAC". */
  const canManage = !!userRole && MANAGER_ROLES.includes(userRole)


  /* ════════════════════════════════════════════════════════════════════════
   * CHARGEMENT INITIAL + REFRESH
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * Recharge la liste depuis le service après mutation (créa / édition).
   * Pas d'appel via useEffect : on n'a pas besoin de "sync" — le composant
   * peut directement rafraîchir aux moments choisis (après createPAC, après
   * clic sur "Actualiser", etc.).
   */
  const refresh = () => {
    setPacs(listPACs())
    setDirections(listDirections(true))
  }


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION + STATS
   * ════════════════════════════════════════════════════════════════════════ */

  /** PAC après application des filtres (côté client, pas d'OData ici). */
  const filtered = useMemo(() => {
    return pacs.filter(p => {
      if (appliedFilters.statut && p.statut !== appliedFilters.statut) return false
      if (appliedFilters.direction && !p.directionsConcernees.includes(appliedFilters.direction)) return false
      if (appliedFilters.annee && String(p.annee) !== appliedFilters.annee) return false
      if (appliedFilters.search) {
        const needle = appliedFilters.search.toLowerCase()
        const haystack = [
          p.intitule,
          p.sourcePac,
          p.descriptionProbleme,
          p.actionsCorrectives,
          p.responsable,
        ].join(' ').toLowerCase()
        if (!haystack.includes(needle)) return false
      }
      return true
    })
  }, [pacs, appliedFilters])

  /** Stats globales pour les cards (recalculées sur la liste filtrée). */
  const stats = useMemo(() => ({
    total: filtered.length,
    enCours: filtered.filter(p => p.statut === 'En cours').length,
    executees: filtered.filter(p => p.statut === 'Exécutée').length,
    nonExecutees: filtered.filter(p => p.statut === 'Non Exécutée').length,
  }), [filtered])

  /**
   * Pagination — porte sur la liste filtrée. resetKey = signature des
   * filtres pour ramener automatiquement à la page 1 quand l'utilisateur
   * change un critère.
   */
  const pagination = usePagination({
    total: filtered.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  const pagedPacs = useMemo(
    () => filtered.slice(pagination.start, pagination.end),
    [filtered, pagination.start, pagination.end],
  )


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS — FILTRES
   * ════════════════════════════════════════════════════════════════════════ */

  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const applyFilters = () => setAppliedFilters(filters)
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS)
    setAppliedFilters(EMPTY_FILTERS)
  }


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS — FORMULAIRE DE CRÉATION
   * ════════════════════════════════════════════════════════════════════════ */

  /** Helper générique pour patcher un champ du formulaire. */
  const updateForm = <K extends keyof typeof form>(key: K, value: (typeof form)[K]) => {
    setForm(prev => ({ ...prev, [key]: value }))
  }

  /**
   * Ajoute une direction à la sélection (idempotent : ignore si déjà présent).
   * Appelé depuis le select "Ajouter une direction" du formulaire.
   */
  const addDirection = (code: string) => {
    if (!code) return
    setForm(prev => (
      prev.directionsConcernees.includes(code)
        ? prev
        : { ...prev, directionsConcernees: [...prev.directionsConcernees, code] }
    ))
  }

  /** Retire une direction de la sélection (depuis le × sur le chip). */
  const removeDirection = (code: string) => {
    setForm(prev => ({
      ...prev,
      directionsConcernees: prev.directionsConcernees.filter(c => c !== code),
    }))
  }

  const openForm = () => {
    setForm(EMPTY_FORM)
    setFormError(null)
    setShowForm(true)
  }
  const closeForm = () => {
    setShowForm(false)
    setFormError(null)
  }

  const submitForm = () => {
    setFormError(null)
    try {
      createPAC({
        sourcePac: form.sourcePac.trim(),
        dateCreation: form.dateCreation,
        intitule: form.intitule.trim(),
        descriptionProbleme: form.descriptionProbleme.trim(),
        causeImmediate: form.causeImmediate.trim(),
        causeRacine: form.causeRacine.trim(),
        actionsCorrectives: form.actionsCorrectives.trim(),
        directionsConcernees: form.directionsConcernees,
        echeance: form.echeance,
        kpi: form.kpi.trim(),
        annee: Number(form.annee) || new Date().getFullYear(),
        responsable: form.responsable.trim(),
        statut: form.statut,
        observations: form.observations.trim() || undefined,
      })
      refresh()
      closeForm()
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Échec de la création du PAC.')
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   * ════════════════════════════════════════════════════════════════════════ */

  return (
    <>
      {/* ─── Header ────────────────────────────────────────────────── */}
      <div className="content-header">
        <h2>Plan d'Action Correctif</h2>
        {canManage && (
          <button className="btn-add" type="button" onClick={openForm}>
            + Nouveau PAC
          </button>
        )}
      </div>

      {/* ─── Stats cards ───────────────────────────────────────────── */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Total PAC</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.enCours}</span>
          <span className="stat-label">En cours</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{stats.executees}</span>
          <span className="stat-label">Exécutées</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{stats.nonExecutees}</span>
          <span className="stat-label">Non Exécutées</span>
        </div>
      </div>

      {/* ─── Barre de filtres ──────────────────────────────────────── */}
      <div className="filters">
        <div className="filter-field" style={{ flex: 1, minWidth: 220 }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="Intitulé, description, responsable..."
            value={filters.search}
            onChange={e => updateFilter('search', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
          />
        </div>
        <div className="filter-field">
          <label>Statut</label>
          <select value={filters.statut} onChange={e => updateFilter('statut', e.target.value as FilterState['statut'])}>
            <option value="">Tous</option>
            {PAC_STATUS_OPTIONS.map(s => (
              <option key={s} value={s}>{s}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Direction</label>
          <select value={filters.direction} onChange={e => updateFilter('direction', e.target.value)}>
            <option value="">Toutes</option>
            {directions.map(d => (
              <option key={d.code} value={d.code}>{d.code} — {d.libelle}</option>
            ))}
          </select>
        </div>
        <div className="filter-field" style={{ maxWidth: 110 }}>
          <label>Année</label>
          <input
            type="number"
            placeholder="2026"
            value={filters.annee}
            onChange={e => updateFilter('annee', e.target.value)}
          />
        </div>
        <button type="button" className="btn-search-filters" onClick={applyFilters}>
          Rechercher
        </button>
        <button type="button" className="btn-reset-filters" onClick={resetFilters}>
          Réinitialiser
        </button>
      </div>

      {/* ─── Tableau des PAC ───────────────────────────────────────── */}
      {filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {pacs.length === 0
            ? 'Aucun PAC enregistré pour le moment.'
            : 'Aucun PAC ne correspond aux critères.'}
        </p>
      ) : (
        <div className="table-container" style={{ marginTop: 16 }}>
          <table className="anomalies-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Source</th>
                <th>Date</th>
                <th>Intitulé</th>
                <th>Directions</th>
                <th>Échéance</th>
                <th>Responsable</th>
                <th>Statut</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pagedPacs.map((p, i) => (
                <tr key={p.id}>
                  <td>{pagination.start + i + 1}</td>
                  <td>{p.sourcePac || '—'}</td>
                  <td>{formatDate(p.dateCreation)}</td>
                  <td style={{ maxWidth: 280 }}>{p.intitule}</td>
                  <td>
                    {p.directionsConcernees.map(c => (
                      <span key={c} className="ticket-tag" style={{ marginRight: 4 }}>{c}</span>
                    ))}
                  </td>
                  <td>{formatDate(p.echeance)}</td>
                  <td>{p.responsable || '—'}</td>
                  <td>
                    <span className={`manager-pill ${getPacStatusClass(p.statut)}`}>{p.statut}</span>
                  </td>
                  <td>
                    <button type="button" className="btn-cta btn-cta-detail" onClick={() => setSelected(p)}>
                      Détail
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <Pagination
            state={pagination}
            total={filtered.length}
            itemLabel="PAC"
          />
        </div>
      )}

      {/* ─── Modale de détail (lecture seule) ──────────────────────── */}
      {selected && (
        <PacDetailModal pac={selected} onClose={() => setSelected(null)} />
      )}

      {/* ─── Modale de création ────────────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
            <div className="modal-header">
              <h2>Nouveau Plan d'Action Correctif</h2>
              <button className="modal-close" onClick={closeForm}>&times;</button>
            </div>
            <div className="modal-body">
              {/* Champs d'identification */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-intitule">Intitulé du PAC *</label>
                <input
                  id="pac-intitule"
                  type="text"
                  value={form.intitule}
                  onChange={e => updateForm('intitule', e.target.value)}
                  placeholder="Ex: Frais de convention SSP"
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-source">Source PAC</label>
                <input
                  id="pac-source"
                  type="text"
                  value={form.sourcePac}
                  onChange={e => updateForm('sourcePac', e.target.value)}
                  placeholder="Ex: Contrôle administratifs"
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-date">Date d'ouverture</label>
                <input
                  id="pac-date"
                  type="date"
                  value={form.dateCreation}
                  onChange={e => updateForm('dateCreation', e.target.value)}
                />
              </div>

              {/* Diagnostic du problème */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-desc">Description du problème</label>
                <textarea
                  id="pac-desc"
                  rows={2}
                  value={form.descriptionProbleme}
                  onChange={e => updateForm('descriptionProbleme', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-cause-im">Cause immédiate</label>
                <textarea
                  id="pac-cause-im"
                  rows={2}
                  value={form.causeImmediate}
                  onChange={e => updateForm('causeImmediate', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-cause-rac">Cause racine</label>
                <textarea
                  id="pac-cause-rac"
                  rows={2}
                  value={form.causeRacine}
                  onChange={e => updateForm('causeRacine', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-actions">Actions correctives à mener</label>
                <textarea
                  id="pac-actions"
                  rows={2}
                  value={form.actionsCorrectives}
                  onChange={e => updateForm('actionsCorrectives', e.target.value)}
                />
              </div>

              {/* Directions concernées (multi-sélection via dropdown + chips)
                  - Select : on n'affiche que les directions PAS encore sélectionnées
                    (évite l'ajout en double et clarifie ce qui reste à choisir)
                  - Le select revient toujours en value="" après ajout
                    (signal pour React de réafficher le placeholder)
                  - Chips : chaque direction sélectionnée s'affiche avec × pour
                    la retirer (UX standard du multi-select à tags) */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-direction-add">Directions concernées *</label>
                {directions.length === 0 ? (
                  <p style={{ color: '#888', fontSize: 12, margin: 0 }}>
                    Aucune direction configurée.
                  </p>
                ) : (
                  <>
                    <select
                      id="pac-direction-add"
                      value=""
                      onChange={e => {
                        addDirection(e.target.value)
                        // Reset visuel : on remet le select sur "Ajouter..." pour
                        // permettre l'ajout successif sans changer manuellement.
                        e.target.value = ''
                      }}
                    >
                      <option value="">
                        {form.directionsConcernees.length === directions.length
                          ? '— Toutes les directions sélectionnées —'
                          : '— Ajouter une direction —'}
                      </option>
                      {directions
                        .filter(d => !form.directionsConcernees.includes(d.code))
                        .map(d => (
                          <option key={d.code} value={d.code}>
                            {d.code} — {d.libelle}
                          </option>
                        ))}
                    </select>
                    {/* Chips des directions sélectionnées */}
                    {form.directionsConcernees.length > 0 && (
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 6 }}>
                        {form.directionsConcernees.map(code => {
                          const dir = directions.find(d => d.code === code)
                          return (
                            <span
                              key={code}
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: 6,
                                padding: '2px 6px 2px 8px',
                                background: '#e0f2fe',
                                border: '1px solid #bae6fd',
                                borderRadius: 12,
                                fontSize: 12,
                              }}
                            >
                              <strong>{code}</strong>
                              {dir && <span style={{ color: '#444' }}>— {dir.libelle}</span>}
                              <button
                                type="button"
                                onClick={() => removeDirection(code)}
                                aria-label={`Retirer ${code}`}
                                style={{
                                  border: 'none',
                                  background: 'transparent',
                                  cursor: 'pointer',
                                  fontSize: 14,
                                  lineHeight: 1,
                                  color: '#0369a1',
                                  padding: '0 2px',
                                }}
                              >
                                ×
                              </button>
                            </span>
                          )
                        })}
                      </div>
                    )}
                  </>
                )}
              </div>

              {/* Pilotage : échéance, KPI, année, responsable, statut */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                <div className="form-field">
                  <label htmlFor="pac-echeance">Échéance</label>
                  <input
                    id="pac-echeance"
                    type="date"
                    value={form.echeance}
                    onChange={e => updateForm('echeance', e.target.value)}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="pac-annee">Année</label>
                  <input
                    id="pac-annee"
                    type="number"
                    value={form.annee}
                    onChange={e => updateForm('annee', Number(e.target.value))}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="pac-resp">Responsable de mise en œuvre</label>
                  <input
                    id="pac-resp"
                    type="text"
                    value={form.responsable}
                    onChange={e => updateForm('responsable', e.target.value)}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="pac-statut">Statut initial</label>
                  <select
                    id="pac-statut"
                    value={form.statut}
                    onChange={e => updateForm('statut', e.target.value as PacStatus)}
                  >
                    {PAC_STATUS_OPTIONS.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="pac-kpi">KPI</label>
                <input
                  id="pac-kpi"
                  type="text"
                  value={form.kpi}
                  onChange={e => updateForm('kpi', e.target.value)}
                  placeholder="Ex: 100% des dossiers traités dans les délais"
                />
              </div>
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="pac-obs">Observations initiales</label>
                <textarea
                  id="pac-obs"
                  rows={2}
                  value={form.observations}
                  onChange={e => updateForm('observations', e.target.value)}
                />
              </div>

              {formError && (
                <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0' }} role="alert">
                  {formError}
                </p>
              )}

              <div className="modal-actions" style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn-cta btn-cta-detail" onClick={closeForm}>
                  Annuler
                </button>
                <button type="button" className="btn-cta btn-cta-affect" onClick={submitForm}>
                  Créer le PAC
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  )
}


/* ══════════════════════════════════════════════════════════════════════════
 * MODALE DÉTAIL (lecture seule)
 *
 * Affiche TOUS les champs d'un PAC y compris ceux non gérés à la création
 * (dernière évaluation, nouveau délai, MEO DCPO) — ces champs seront édités
 * dans une itération ultérieure via un mode édition.
 * ══════════════════════════════════════════════════════════════════════════ */

function PacDetailModal({ pac, onClose }: { pac: Pac; onClose: () => void }) {
  // Escape ferme la modale (cohérent avec les autres modales du projet)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
        <div className="modal-header">
          <h2>Détail du PAC</h2>
          <button className="modal-close" onClick={onClose}>&times;</button>
        </div>
        <div className="modal-body">
          <dl className="detail-grid">
            <dt>Intitulé</dt><dd><strong>{pac.intitule}</strong></dd>
            <dt>Source</dt><dd>{pac.sourcePac || '—'}</dd>
            <dt>Date d'ouverture</dt><dd>{formatDate(pac.dateCreation)}</dd>
            <dt>Année</dt><dd>{pac.annee}</dd>
            <dt>Statut</dt>
            <dd>
              <span className={`manager-pill ${getPacStatusClass(pac.statut)}`}>{pac.statut}</span>
            </dd>
            <dt>Directions concernées</dt>
            <dd>
              {pac.directionsConcernees.length === 0
                ? '—'
                : pac.directionsConcernees.map(c => (
                    <span key={c} className="ticket-tag" style={{ marginRight: 6 }}>
                      {c} — {getDirectionLabel(c)}
                    </span>
                  ))}
            </dd>
            <dt>Échéance</dt><dd>{formatDate(pac.echeance)}</dd>
            <dt>Responsable</dt><dd>{pac.responsable || '—'}</dd>
            <dt>KPI</dt><dd>{pac.kpi || '—'}</dd>

            <dt>Description du problème</dt><dd>{pac.descriptionProbleme || '—'}</dd>
            <dt>Cause immédiate</dt><dd>{pac.causeImmediate || '—'}</dd>
            <dt>Cause racine</dt><dd>{pac.causeRacine || '—'}</dd>
            <dt>Actions correctives</dt><dd>{pac.actionsCorrectives || '—'}</dd>

            {pac.observations && (<><dt>Observations</dt><dd>{pac.observations}</dd></>)}
            {pac.derniereEvaluation && (<><dt>Dernière évaluation</dt><dd>{formatDate(pac.derniereEvaluation)}</dd></>)}
            {pac.nouveauDelai && (<><dt>Nouveau délai proposé</dt><dd>{formatDate(pac.nouveauDelai)}</dd></>)}
            {pac.meoDcpo && (<><dt>MEO DCPO</dt><dd>{pac.meoDcpo}</dd></>)}
          </dl>
        </div>
      </div>
    </div>
  )
}
