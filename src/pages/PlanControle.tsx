/**
 * ============================================================================
 * MODULE — PLAN DE CONTRÔLE (PLAN ANNUEL DCPO)
 * ============================================================================
 *
 * Permet aux managers de créer et suivre les activités de contrôle prévues
 * sur l'année. Inspiré du fichier Excel "PLAN DE CONTROLE 2025" fourni par
 * l'utilisateur, avec :
 *   - une catégorie (regroupement thématique des contrôles)
 *   - un libellé d'activité
 *   - une fréquence (Quotidienne / Hebdomadaire / Mensuelle / Annuelle)
 *   - un responsable
 *   - des objectifs (qualitatif + chiffré/KPI)
 *   - le planning mensuel (cases Jan-Déc cochées)
 *   - un statut workflow
 *
 * Structure UI :
 *   - Stats cards : Total + compteurs par statut
 *   - Filtres : catégorie, fréquence, responsable, statut, année
 *   - Tableau paginé avec 12 colonnes mois (✓ si planifié)
 *   - Bouton "Nouveau contrôle" (managers uniquement)
 *   - Modale création (formulaire complet avec checkboxes mois)
 *   - Modale détail (lecture seule)
 *
 * Persistance : localStorage via src/lib/planControleService.ts. Une fois
 * la liste SharePoint DCPO_PLAN_CONTROLE créée, basculer côté service
 * (le composant React reste inchangé).
 * ============================================================================
 */

import { useMemo, useState, useEffect } from 'react'
import {
  createControle,
  listControles,
  getControleStatusClass,
  MOIS,
  PLAN_CONTROLE_CATEGORIES,
  PLAN_CONTROLE_RESPONSABLES_SUGGESTIONS,
  FREQUENCE_OPTIONS,
  STATUS_OPTIONS,
  type ControleEntry,
  type ControleFrequence,
  type ControleStatus,
  type MoisCode,
} from '../lib/planControleService'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'


interface PlanControleProps {
  userName?: string
  userEmail?: string
  userRole?: string
}

/** Rôles autorisés à créer / éditer un contrôle. */
const MANAGER_ROLES = ['Chef_Departement', 'Directeur']

/** État vide pour la barre de filtres. */
const EMPTY_FILTERS = {
  categorie: '',
  frequence: '' as '' | ControleFrequence,
  responsable: '',
  statut: '' as '' | ControleStatus,
  annee: '',
  search: '',
}

type FilterState = typeof EMPTY_FILTERS

/** État vide pour le formulaire de création. */
const EMPTY_FORM = {
  libelle: '',
  categorie: PLAN_CONTROLE_CATEGORIES[0],
  objectif: '',
  objectifChiffre: '',
  frequence: 'Mensuelle' as ControleFrequence,
  responsable: '',
  moisPlanifies: [] as MoisCode[],
  annee: new Date().getFullYear(),
  statut: 'À planifier' as ControleStatus,
}


export default function PlanControle({ userRole }: PlanControleProps) {
  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * Lazy init via useState(() => …) pour éviter le pattern fetch-in-useEffect
   * (cf. règle react-hooks/set-state-in-effect appliquée dans tout le projet).
   * ════════════════════════════════════════════════════════════════════════ */

  const [controles, setControles] = useState<ControleEntry[]>(() => listControles())
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [selected, setSelected] = useState<ControleEntry | null>(null)
  const [showForm, setShowForm] = useState(false)
  const [form, setForm] = useState(EMPTY_FORM)
  const [formError, setFormError] = useState<string | null>(null)

  /** Permission "créer un contrôle". */
  const canManage = !!userRole && MANAGER_ROLES.includes(userRole)


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION + STATS
   * ════════════════════════════════════════════════════════════════════════ */

  /** Liste filtrée côté client (toutes les colonnes sont en mémoire). */
  const filtered = useMemo(() => {
    return controles.filter(c => {
      if (appliedFilters.categorie && c.categorie !== appliedFilters.categorie) return false
      if (appliedFilters.frequence && c.frequence !== appliedFilters.frequence) return false
      if (appliedFilters.statut && c.statut !== appliedFilters.statut) return false
      if (appliedFilters.annee && String(c.annee) !== appliedFilters.annee) return false
      if (appliedFilters.responsable) {
        const needle = appliedFilters.responsable.toLowerCase()
        if (!c.responsable.toLowerCase().includes(needle)) return false
      }
      if (appliedFilters.search) {
        const needle = appliedFilters.search.toLowerCase()
        const haystack = `${c.libelle} ${c.objectif} ${c.objectifChiffre}`.toLowerCase()
        if (!haystack.includes(needle)) return false
      }
      return true
    })
  }, [controles, appliedFilters])

  /** Stats globales (sur la liste filtrée). */
  const stats = useMemo(() => ({
    total: filtered.length,
    aPlanifier: filtered.filter(c => c.statut === 'À planifier').length,
    planifies: filtered.filter(c => c.statut === 'Planifié').length,
    enCours: filtered.filter(c => c.statut === 'En cours').length,
    realises: filtered.filter(c => c.statut === 'Réalisé').length,
  }), [filtered])

  const pagination = usePagination({
    total: filtered.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  const pagedControles = useMemo(
    () => filtered.slice(pagination.start, pagination.end),
    [filtered, pagination.start, pagination.end],
  )


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS
   * ════════════════════════════════════════════════════════════════════════ */

  const refresh = () => setControles(listControles())

  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const applyFilters = () => setAppliedFilters(filters)
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS)
    setAppliedFilters(EMPTY_FILTERS)
  }

  const updateForm = <K extends keyof typeof form>(key: K, value: (typeof form)[K]) => {
    setForm(prev => ({ ...prev, [key]: value }))
  }

  /**
   * Toggle un mois dans la sélection (checkbox dans le formulaire).
   * Le tri final est garanti par l'ordre lexicographique des codes '01'..'12'.
   */
  const toggleMois = (code: MoisCode) => {
    setForm(prev => ({
      ...prev,
      moisPlanifies: prev.moisPlanifies.includes(code)
        ? prev.moisPlanifies.filter(m => m !== code)
        : [...prev.moisPlanifies, code].sort(),
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
      createControle({
        libelle: form.libelle,
        categorie: form.categorie,
        objectif: form.objectif,
        objectifChiffre: form.objectifChiffre,
        frequence: form.frequence,
        responsable: form.responsable,
        moisPlanifies: form.moisPlanifies,
        annee: Number(form.annee) || new Date().getFullYear(),
        statut: form.statut,
      })
      refresh()
      closeForm()
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Échec de la création du contrôle.')
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * RENDU
   * ════════════════════════════════════════════════════════════════════════ */

  return (
    <>
      <div className="content-header">
        <h2>Plan de Contrôle</h2>
        {canManage && (
          <button className="btn-add" type="button" onClick={openForm}>
            + Nouveau contrôle
          </button>
        )}
      </div>

      {/* ─── Stats cards ───────────────────────────────────────────── */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Total</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.aPlanifier}</span>
          <span className="stat-label">À planifier</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.planifies}</span>
          <span className="stat-label">Planifiés</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.enCours}</span>
          <span className="stat-label">En cours</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{stats.realises}</span>
          <span className="stat-label">Réalisés</span>
        </div>
      </div>

      {/* ─── Barre de filtres ──────────────────────────────────────── */}
      <div className="filters">
        <div className="filter-field" style={{ flex: 1, minWidth: 220 }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="Libellé, objectif, KPI..."
            value={filters.search}
            onChange={e => updateFilter('search', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
          />
        </div>
        <div className="filter-field">
          <label>Catégorie</label>
          <select value={filters.categorie} onChange={e => updateFilter('categorie', e.target.value)}>
            <option value="">Toutes</option>
            {PLAN_CONTROLE_CATEGORIES.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Fréquence</label>
          <select value={filters.frequence} onChange={e => updateFilter('frequence', e.target.value as FilterState['frequence'])}>
            <option value="">Toutes</option>
            {FREQUENCE_OPTIONS.map(f => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Statut</label>
          <select value={filters.statut} onChange={e => updateFilter('statut', e.target.value as FilterState['statut'])}>
            <option value="">Tous</option>
            {STATUS_OPTIONS.map(s => (
              <option key={s} value={s}>{s}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Responsable</label>
          <input
            type="text"
            placeholder="Nom..."
            value={filters.responsable}
            onChange={e => updateFilter('responsable', e.target.value)}
          />
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

      {/* ─── Tableau ────────────────────────────────────────────────── */}
      {filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {controles.length === 0
            ? 'Aucun contrôle planifié pour le moment.'
            : 'Aucun contrôle ne correspond aux critères.'}
        </p>
      ) : (
        <div className="table-container" style={{ marginTop: 16, overflowX: 'auto' }}>
          <table className="anomalies-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Catégorie</th>
                <th>Libellé</th>
                <th>Fréquence</th>
                <th>Responsable</th>
                {MOIS.map(m => (
                  <th key={m.code} style={{ minWidth: 38, textAlign: 'center' }}>{m.label}</th>
                ))}
                <th>Statut</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pagedControles.map((c, i) => (
                <tr key={c.id}>
                  <td>{pagination.start + i + 1}</td>
                  <td style={{ maxWidth: 200 }}>{c.categorie}</td>
                  <td style={{ maxWidth: 280 }}>{c.libelle}</td>
                  <td>{c.frequence}</td>
                  <td>{c.responsable || '—'}</td>
                  {MOIS.map(m => (
                    <td key={m.code} style={{ textAlign: 'center' }}>
                      {c.moisPlanifies.includes(m.code) ? '✓' : ''}
                    </td>
                  ))}
                  <td>
                    <span className={`manager-pill ${getControleStatusClass(c.statut)}`}>{c.statut}</span>
                  </td>
                  <td>
                    <button type="button" className="btn-cta btn-cta-detail" onClick={() => setSelected(c)}>
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
            itemLabel="contrôles"
          />
        </div>
      )}

      {/* ─── Modale détail ──────────────────────────────────────────── */}
      {selected && (
        <ControleDetailModal entry={selected} onClose={() => setSelected(null)} />
      )}

      {/* ─── Modale création ───────────────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
            <div className="modal-header">
              <h2>Nouveau contrôle</h2>
              <button className="modal-close" onClick={closeForm}>&times;</button>
            </div>
            <div className="modal-body">
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="ctrl-libelle">Libellé du contrôle *</label>
                <input
                  id="ctrl-libelle"
                  type="text"
                  value={form.libelle}
                  onChange={e => updateForm('libelle', e.target.value)}
                  placeholder="Ex: Evaluation du contrôle des opérations SYSTAC"
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="ctrl-categorie">Catégorie *</label>
                <select
                  id="ctrl-categorie"
                  value={form.categorie}
                  onChange={e => updateForm('categorie', e.target.value)}
                >
                  {PLAN_CONTROLE_CATEGORIES.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                <div className="form-field">
                  <label htmlFor="ctrl-freq">Fréquence</label>
                  <select
                    id="ctrl-freq"
                    value={form.frequence}
                    onChange={e => updateForm('frequence', e.target.value as ControleFrequence)}
                  >
                    {FREQUENCE_OPTIONS.map(f => (
                      <option key={f} value={f}>{f}</option>
                    ))}
                  </select>
                </div>
                <div className="form-field">
                  <label htmlFor="ctrl-annee">Année</label>
                  <input
                    id="ctrl-annee"
                    type="number"
                    value={form.annee}
                    onChange={e => updateForm('annee', Number(e.target.value))}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="ctrl-resp">Responsable</label>
                  <input
                    id="ctrl-resp"
                    type="text"
                    list="resp-suggestions"
                    value={form.responsable}
                    onChange={e => updateForm('responsable', e.target.value)}
                    placeholder="Ex: KAMEDA, Tous les contrôleurs..."
                  />
                  {/* Suggestions via datalist : aide à la saisie sans
                      restreindre à la liste (un nouveau responsable est OK) */}
                  <datalist id="resp-suggestions">
                    {PLAN_CONTROLE_RESPONSABLES_SUGGESTIONS.map(r => (
                      <option key={r} value={r} />
                    ))}
                  </datalist>
                </div>
                <div className="form-field">
                  <label htmlFor="ctrl-statut">Statut initial</label>
                  <select
                    id="ctrl-statut"
                    value={form.statut}
                    onChange={e => updateForm('statut', e.target.value as ControleStatus)}
                  >
                    {STATUS_OPTIONS.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="ctrl-obj">Objectif</label>
                <textarea
                  id="ctrl-obj"
                  rows={2}
                  value={form.objectif}
                  onChange={e => updateForm('objectif', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="ctrl-kpi">Objectif chiffré (KPI)</label>
                <textarea
                  id="ctrl-kpi"
                  rows={2}
                  value={form.objectifChiffre}
                  onChange={e => updateForm('objectifChiffre', e.target.value)}
                  placeholder="Ex: 01 rapport de contrôle mensuel"
                />
              </div>

              {/* Planning mensuel : grille 12 cases */}
              <div className="form-field" style={{ marginTop: 8 }}>
                <label>Mois planifiés</label>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 4 }}>
                  {MOIS.map(m => {
                    const checked = form.moisPlanifies.includes(m.code)
                    return (
                      <label
                        key={m.code}
                        style={{
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: 4,
                          padding: '4px 8px',
                          border: '1px solid #ddd',
                          borderRadius: 4,
                          background: checked ? '#e0f2fe' : '#fff',
                          cursor: 'pointer',
                          fontSize: 12,
                          minWidth: 48,
                          justifyContent: 'center',
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleMois(m.code)}
                        />
                        {m.label}
                      </label>
                    )
                  })}
                </div>
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
                  Créer le contrôle
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
 * ══════════════════════════════════════════════════════════════════════════ */

function ControleDetailModal({ entry, onClose }: { entry: ControleEntry; onClose: () => void }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(640px, 100%)' }}>
        <div className="modal-header">
          <h2>Détail du contrôle</h2>
          <button className="modal-close" onClick={onClose}>&times;</button>
        </div>
        <div className="modal-body">
          <dl className="detail-grid">
            <dt>Libellé</dt><dd><strong>{entry.libelle}</strong></dd>
            <dt>Catégorie</dt><dd>{entry.categorie}</dd>
            <dt>Fréquence</dt><dd>{entry.frequence}</dd>
            <dt>Responsable</dt><dd>{entry.responsable || '—'}</dd>
            <dt>Année</dt><dd>{entry.annee}</dd>
            <dt>Statut</dt>
            <dd>
              <span className={`manager-pill ${getControleStatusClass(entry.statut)}`}>{entry.statut}</span>
            </dd>
            <dt>Objectif</dt><dd>{entry.objectif || '—'}</dd>
            <dt>Objectif chiffré</dt><dd>{entry.objectifChiffre || '—'}</dd>
            <dt>Mois planifiés</dt>
            <dd>
              {entry.moisPlanifies.length === 0
                ? '—'
                : entry.moisPlanifies
                    .map(code => MOIS.find(m => m.code === code)?.label ?? code)
                    .join(', ')}
            </dd>
          </dl>
        </div>
      </div>
    </div>
  )
}
