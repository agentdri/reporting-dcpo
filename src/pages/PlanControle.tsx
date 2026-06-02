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
 *   - un responsable (champ Personne Office 365)
 *   - des objectifs (qualitatif + chiffré/KPI)
 *   - un statut workflow
 *
 * Structure UI :
 *   - Stats cards : Total + compteurs par statut
 *   - Filtres : catégorie, fréquence, responsable, statut, année
 *   - Tableau paginé
 *   - Bouton "Nouveau contrôle" (managers uniquement)
 *   - Modale création
 *   - Modale détail (lecture seule)
 *
 * Persistance : liste SharePoint DCPO_LISTE_PLAN_CONTROLE via
 * src/lib/planControleService.ts.
 * ============================================================================
 */

import { useMemo, useState, useEffect } from 'react'
import {
  createControle,
  updateControle,
  appendPlanControleAttachmentUrls,
  listControles,
  getControleStatusClass,
  PLAN_CONTROLE_CATEGORIES,
  FREQUENCE_OPTIONS,
  STATUS_OPTIONS,
  type ControleEntry,
  type ControleFrequence,
  type ControleStatus,
} from '../lib/planControleService'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { UserPicker } from '../components/UserPicker'
import {
  uploadPlanControleAttachment,
  getAttachmentIcon,
  getAttachmentIconType,
} from '../lib/ticketAttachments'


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
  responsableName: '',
  responsableEmail: '',
  annee: new Date().getFullYear(),
  statut: 'À planifier' as ControleStatus,
}


export default function PlanControle({ userEmail, userRole }: PlanControleProps) {
  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * ════════════════════════════════════════════════════════════════════════ */

  const [controles, setControles] = useState<ControleEntry[]>([])
  const [loading, setLoading] = useState(true)
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [selected, setSelected] = useState<ControleEntry | null>(null)
  const [showForm, setShowForm] = useState(false)
  /**
   * ID du contrôle en cours d'édition (null = mode création).
   * Une seule modale sert aux deux modes.
   */
  const [editingId, setEditingId] = useState<string | null>(null)
  const [form, setForm] = useState(EMPTY_FORM)
  const [formError, setFormError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)
  /**
   * Pièce jointe optionnelle. À la création comme à la modification : uploadée
   * via Power Automate APRÈS création/update de l'item SP (le workflow a
   * besoin de l'ID), puis l'URL renvoyée est ajoutée au champ urlPieceJointe.
   */
  const [attachment, setAttachment] = useState<File | null>(null)

  /** Permission "créer un contrôle". */
  const canManage = !!userRole && MANAGER_ROLES.includes(userRole)

  /* ════════════════════════════════════════════════════════════════════════
   * CHARGEMENT DEPUIS SHAREPOINT
   * Le fetch async dans un effet est autorisé (setState dans un callback
   * après await, pas dans le corps synchrone de l'effet).
   * ════════════════════════════════════════════════════════════════════════ */

  useEffect(() => {
    let cancelled = false
    listControles()
      .then(data => { if (!cancelled) setControles(data) })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [])


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION + STATS
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * Liste filtrée côté client (toutes les colonnes sont en mémoire).
   *
   * Restriction de visibilité par rôle :
   *   - Controleur → ne voit QUE les contrôles dont il est responsable
   *     (responsableEmail = userEmail, case-insensitive)
   *   - Manager (Chef_Departement / Directeur) → voit tout
   *
   * Le filtre rôle est appliqué AVANT les filtres de la barre de recherche
   * pour que les compteurs/stats reflètent uniquement ce que l'utilisateur
   * a réellement le droit de voir.
   */
  const filtered = useMemo(() => {
    const myEmail = userEmail?.toLowerCase()
    const restrictToMine = userRole === 'Controleur'
    return controles.filter(c => {
      // ─── Garde de visibilité par rôle (cf. doc ci-dessus) ─────────────
      if (restrictToMine) {
        if (!myEmail || c.responsableEmail.toLowerCase() !== myEmail) return false
      }
      // ─── Filtres de la barre de recherche ────────────────────────────
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
  }, [controles, appliedFilters, userRole, userEmail])

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

  const refresh = async () => {
    setLoading(true)
    try {
      setControles(await listControles())
    } finally {
      setLoading(false)
    }
  }

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

  /** Ouvre le formulaire en mode CRÉATION. */
  const openForm = () => {
    setEditingId(null)
    setForm(EMPTY_FORM)
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
  }

  /** Ouvre le formulaire en mode ÉDITION pré-rempli depuis un contrôle. */
  const openEdit = (ctrl: ControleEntry) => {
    setEditingId(ctrl.id)
    setForm({
      libelle: ctrl.libelle,
      categorie: ctrl.categorie,
      objectif: ctrl.objectif,
      objectifChiffre: ctrl.objectifChiffre,
      frequence: ctrl.frequence,
      responsableName: ctrl.responsable,
      responsableEmail: ctrl.responsableEmail,
      annee: ctrl.annee,
      statut: ctrl.statut,
    })
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
    // Ferme la modale détail pour laisser la place à la modale édition.
    setSelected(null)
  }

  const closeForm = () => {
    setShowForm(false)
    setEditingId(null)
    setAttachment(null)
    setFormError(null)
  }

  /**
   * Crée OU met à jour le contrôle (selon editingId), puis enchaîne l'upload
   * de la pièce jointe et la persistance de son URL dans le champ
   * `urlPieceJointe` via `appendPlanControleAttachmentUrls`.
   *
   * Si la création/update réussit mais l'upload échoue, on garde la modale
   * ouverte avec un message d'erreur — l'item existe déjà en SP, pas de rollback.
   */
  const submitForm = async () => {
    setFormError(null)
    setSaving(true)
    try {
      const input = {
        libelle: form.libelle,
        categorie: form.categorie,
        objectif: form.objectif,
        objectifChiffre: form.objectifChiffre,
        frequence: form.frequence,
        responsableName: form.responsableName,
        responsableEmail: form.responsableEmail,
        annee: Number(form.annee) || new Date().getFullYear(),
        statut: form.statut,
      }

      const saved = editingId
        ? await updateControle(editingId, input)
        : await createControle(input)

      if (!saved?.id) {
        throw new Error(editingId ? 'Échec de la mise à jour.' : 'Échec de la création.')
      }

      // Upload optionnel via workflow Power Automate + persistance de l'URL.
      if (attachment && saved.id) {
        try {
          const uploadedUrl = await uploadPlanControleAttachment(saved.id, attachment, 'Visite')
          if (uploadedUrl) {
            await appendPlanControleAttachmentUrls(saved.id, [uploadedUrl])
          }
        } catch (uploadErr) {
          const detail = uploadErr instanceof Error ? uploadErr.message : String(uploadErr)
          console.error('Échec upload pièce jointe Plan de Contrôle', uploadErr)
          setFormError(`Le contrôle a été enregistré mais la pièce jointe a échoué : ${detail}`)
          await refresh()
          return  // garde la modale ouverte
        }
      }

      await refresh()
      closeForm()
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Échec de l\'enregistrement du contrôle.')
    } finally {
      setSaving(false)
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
      <div className="filters-bar">
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
        <div className="filter-actions">
          <button type="button" className="btn-search-filters" onClick={applyFilters}>
            Rechercher
          </button>
          <button type="button" className="btn-reset-filters" onClick={resetFilters}>
            Réinitialiser
          </button>
        </div>
      </div>

      {/* ─── Tableau ────────────────────────────────────────────────── */}
      {loading ? (
        <p className="loading-text" style={{ marginTop: 16 }}>Chargement des contrôles...</p>
      ) : filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {controles.length === 0
            ? 'Aucun contrôle planifié pour le moment.'
            : 'Aucun contrôle ne correspond aux critères.'}
        </p>
      ) : (
        <div className="table-wrapper" style={{ marginTop: 16 }}>
          <table className="data-table anomalies-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Catégorie</th>
                <th>Libellé</th>
                <th>Fréquence</th>
                <th>Responsable</th>
                <th>Année</th>
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
                  <td>{c.annee}</td>
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
        <ControleDetailModal
          entry={selected}
          onClose={() => setSelected(null)}
          onEdit={canManage ? () => openEdit(selected) : undefined}
        />
      )}

      {/* ─── Modale création ───────────────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
            <div className="modal-header">
              <h2>{editingId ? 'Modifier le contrôle' : 'Nouveau contrôle'}</h2>
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
                  {/* Champ Personne SharePoint → sélecteur Office 365 */}
                  <UserPicker
                    id="ctrl-resp"
                    selectedName={form.responsableName}
                    selectedEmail={form.responsableEmail}
                    onSelect={(name, email) => {
                      updateForm('responsableName', name)
                      updateForm('responsableEmail', email)
                    }}
                    onClear={() => {
                      updateForm('responsableName', '')
                      updateForm('responsableEmail', '')
                    }}
                    disabled={saving}
                  />
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

              {/* Pièce jointe optionnelle — uploadée après création via Power
                  Automate (workflow PLAN_CONTROLE_ATTACHMENT_API_URL). Attache
                  le fichier à l'item SharePoint nouvellement créé. */}
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="ctrl-attachment">Pièce jointe (optionnel)</label>
                <input
                  id="ctrl-attachment"
                  type="file"
                  onChange={e => setAttachment(e.target.files?.[0] ?? null)}
                  disabled={saving}
                />
                {attachment && (
                  <span className="selected-email">📎 {attachment.name}</span>
                )}
              </div>

              {formError && (
                <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0' }} role="alert">
                  {formError}
                </p>
              )}

              <div className="modal-actions" style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn-cta btn-cta-detail" onClick={closeForm} disabled={saving}>
                  Annuler
                </button>
                <button type="button" className="btn-cta btn-cta-affect" onClick={submitForm} disabled={saving}>
                  {saving
                    ? (editingId ? 'Enregistrement...' : 'Création...')
                    : (editingId ? 'Enregistrer' : 'Créer le contrôle')}
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
 * MODALE DÉTAIL (lecture seule + bouton Modifier + pièces jointes)
 *
 * Le bouton Modifier n'apparaît que si `onEdit` est fourni — décision prise
 * côté parent en fonction du rôle (canManage).
 * ══════════════════════════════════════════════════════════════════════════ */

interface ControleDetailModalProps {
  entry: ControleEntry
  onClose: () => void
  /** Callback pour basculer en mode édition (undefined = bouton masqué). */
  onEdit?: () => void
}

function ControleDetailModal({ entry, onClose, onEdit }: ControleDetailModalProps) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const attachments = entry.attachments ?? []

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(640px, 100%)' }}>
        <div className="modal-header" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <h2 style={{ flex: 1 }}>Détail du contrôle</h2>
          {onEdit && (
            <button type="button" className="btn-cta btn-cta-detail" onClick={onEdit}>
              ✎ Modifier
            </button>
          )}
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
          </dl>

          {/* Pièces jointes — icônes selon extension, lien externe vers le fichier */}
          <div className="detail-attachments" style={{ marginTop: 16 }}>
            <h3 style={{ marginBottom: 8 }}>Pièces jointes</h3>
            {attachments.length === 0 ? (
              <p className="loading-text">Aucune pièce jointe.</p>
            ) : (
              <ul className="detail-attachments-list">
                {attachments.map((att, i) => {
                  const iconType = getAttachmentIconType(att.name)
                  return (
                    <li key={i} className="detail-attachment-item">
                      <span aria-hidden="true" style={{ fontSize: 24, width: 56, textAlign: 'center', flexShrink: 0 }}>
                        {getAttachmentIcon(iconType)}
                      </span>
                      <a href={att.url} target="_blank" rel="noreferrer" style={{ color: '#1d4ed8', textDecoration: 'none', fontWeight: 600, wordBreak: 'break-all' }}>
                        {att.name}
                      </a>
                    </li>
                  )
                })}
              </ul>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
