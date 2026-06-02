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
 * Persistance des PAC : liste SharePoint DCPO_LISTE_PLAN_ACTION_CORRECTIF
 * via src/lib/pacService.ts. Le référentiel des directions reste local
 * (constante, pas de liste SharePoint dédiée).
 *
 * Source des champs : feuille "PAC DCPO" du fichier Excel
 * Tableau_de_bord_KPI_DCPO_2026.xlsx fourni par l'utilisateur.
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import {
  createPAC,
  updatePAC,
  appendPacAttachmentUrls,
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
import { UserPicker } from '../components/UserPicker'
import { uploadPacAttachment, getAttachmentIcon, getAttachmentIconType } from '../lib/ticketAttachments'


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
  responsableName: '',
  responsableEmail: '',
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


export default function PlanActionCorrectif({ userEmail, userRole }: PlanActionCorrectifProps) {
  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * ════════════════════════════════════════════════════════════════════════ */

  /** Liste des PAC (chargée depuis SharePoint). */
  const [pacs, setPacs] = useState<Pac[]>([])
  /** Liste des directions actives (référentiel local, synchrone). */
  const [directions] = useState<Direction[]>(() => listDirections(true))
  const [loading, setLoading] = useState(true)
  /** Filtres en cours de saisie (binding inputs). */
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** Filtres effectivement appliqués (snapshot au clic Rechercher). */
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** PAC sélectionné pour la modale de détail (null = fermée). */
  const [selected, setSelected] = useState<Pac | null>(null)
  /** Affichage du formulaire (création ou édition selon `editingId`). */
  const [showForm, setShowForm] = useState(false)
  /**
   * ID du PAC en cours d'édition (null = mode création).
   * Une seule modale sert aux deux modes — le label et l'action submit
   * s'adaptent en fonction de cet ID.
   */
  const [editingId, setEditingId] = useState<string | null>(null)
  /** Snapshot des champs en cours de saisie dans le formulaire. */
  const [form, setForm] = useState(EMPTY_FORM)
  /** Erreur de soumission affichée sous le formulaire. */
  const [formError, setFormError] = useState<string | null>(null)
  /** True pendant la sauvegarde d'un nouveau PAC. */
  const [saving, setSaving] = useState(false)
  /**
   * Pièce jointe optionnelle attachée au PAC à la création.
   * Uploadée via Power Automate (uploadPacAttachment) APRÈS la création de
   * l'item SP, car le workflow a besoin de l'ID du record pour attacher.
   */
  const [attachment, setAttachment] = useState<File | null>(null)

  /** Permission "créer / éditer un PAC". */
  const canManage = !!userRole && MANAGER_ROLES.includes(userRole)


  /* ════════════════════════════════════════════════════════════════════════
   * CHARGEMENT INITIAL + REFRESH (SharePoint)
   * ════════════════════════════════════════════════════════════════════════ */

  // Chargement initial depuis SharePoint. Le fetch async dans un effet est
  // autorisé (le setState a lieu dans un callback après await).
  useEffect(() => {
    let cancelled = false
    listPACs()
      .then(data => { if (!cancelled) setPacs(data) })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [])

  /** Recharge la liste depuis SharePoint après création. */
  const refresh = async () => {
    setLoading(true)
    try {
      setPacs(await listPACs())
    } finally {
      setLoading(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION + STATS
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * PAC après application des filtres (côté client, pas d'OData ici).
   *
   * Restriction de visibilité par rôle :
   *   - Controleur → ne voit QUE les PAC dont il est le responsable de mise
   *     en œuvre (responsableEmail = userEmail, case-insensitive)
   *   - Manager (Chef_Departement / Directeur) → voit tout
   *
   * Le filtre rôle est appliqué AVANT les filtres de recherche pour que
   * les compteurs/stats reflètent uniquement ce que l'utilisateur a réellement
   * le droit de voir.
   */
  const filtered = useMemo(() => {
    const myEmail = userEmail?.toLowerCase()
    const restrictToMine = userRole === 'Controleur'
    return pacs.filter(p => {
      // ─── Garde de visibilité par rôle ─────────────────────────────────
      if (restrictToMine) {
        if (!myEmail || p.responsableEmail.toLowerCase() !== myEmail) return false
      }
      // ─── Filtres de la barre de recherche ────────────────────────────
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
  }, [pacs, appliedFilters, userRole, userEmail])

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

  /** Ouvre le formulaire en mode CRÉATION (champs vides). */
  const openForm = () => {
    setEditingId(null)
    setForm(EMPTY_FORM)
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
  }

  /**
   * Ouvre le formulaire en mode ÉDITION (champs pré-remplis depuis un PAC).
   * Le picker UserPicker prend en charge un nom/email pré-rempli via le
   * composant UserPickerEdit (props `initialName` / `initialEmail`).
   */
  const openEdit = (pac: Pac) => {
    setEditingId(pac.id)
    setForm({
      sourcePac: pac.sourcePac,
      dateCreation: pac.dateCreation,
      intitule: pac.intitule,
      descriptionProbleme: pac.descriptionProbleme,
      causeImmediate: pac.causeImmediate,
      causeRacine: pac.causeRacine,
      actionsCorrectives: pac.actionsCorrectives,
      directionsConcernees: pac.directionsConcernees,
      echeance: pac.echeance,
      kpi: pac.kpi,
      annee: pac.annee,
      responsableName: pac.responsable,
      responsableEmail: pac.responsableEmail,
      statut: pac.statut,
      observations: pac.observations ?? '',
    })
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
    // On ferme la modale détail pour que la modale édition soit lisible.
    setSelected(null)
  }

  const closeForm = () => {
    setShowForm(false)
    setEditingId(null)
    setAttachment(null)
    setFormError(null)
  }

  /**
   * Crée OU met à jour le PAC en SharePoint, puis enchaîne (si applicable)
   * l'upload Power Automate de la pièce jointe + persistance de son URL
   * dans le champ `urlPiecesJointes` via `appendPacAttachmentUrls`.
   *
   * Comportement en cas d'échec d'upload :
   *   - Le PAC est déjà créé/modifié en SP → pas de rollback
   *   - On affiche un avertissement et on garde la modale ouverte
   */
  const submitForm = async () => {
    setFormError(null)
    setSaving(true)
    try {
      const input = {
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
        responsableName: form.responsableName,
        responsableEmail: form.responsableEmail,
        statut: form.statut,
        observations: form.observations.trim() || undefined,
      }

      // Création OU édition selon editingId
      const saved = editingId
        ? await updatePAC(editingId, input)
        : await createPAC(input)

      if (!saved?.id) {
        throw new Error(editingId ? 'Échec de la mise à jour.' : 'Échec de la création.')
      }

      // Upload optionnel de la pièce jointe via le workflow Power Automate dédié.
      // L'URL renvoyée est ensuite concaténée dans le champ urlPiecesJointes
      // (multi-URLs séparées par " | ", idem anomalies).
      if (attachment && saved.id) {
        try {
          const uploadedUrl = await uploadPacAttachment(saved.id, attachment, 'Visite')
          if (uploadedUrl) {
            await appendPacAttachmentUrls(saved.id, [uploadedUrl])
          }
        } catch (uploadErr) {
          const detail = uploadErr instanceof Error ? uploadErr.message : String(uploadErr)
          console.error('Échec upload pièce jointe PAC', uploadErr)
          setFormError(`Le PAC a été enregistré mais la pièce jointe a échoué : ${detail}`)
          await refresh()
          return  // on garde la modale ouverte
        }
      }

      await refresh()
      closeForm()
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Échec de l\'enregistrement du PAC.')
    } finally {
      setSaving(false)
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
      <div className="filters-bar">
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
        <div className="filter-actions">
          <button type="button" className="btn-search-filters" onClick={applyFilters}>
            Rechercher
          </button>
          <button type="button" className="btn-reset-filters" onClick={resetFilters}>
            Réinitialiser
          </button>
        </div>
      </div>

      {/* ─── Tableau des PAC ───────────────────────────────────────── */}
      {loading ? (
        <p className="loading-text" style={{ marginTop: 16 }}>Chargement des PAC...</p>
      ) : filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {pacs.length === 0
            ? 'Aucun PAC enregistré pour le moment.'
            : 'Aucun PAC ne correspond aux critères.'}
        </p>
      ) : (
        <div className="table-wrapper" style={{ marginTop: 16 }}>
          <table className="data-table anomalies-table">
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
        <PacDetailModal
          pac={selected}
          onClose={() => setSelected(null)}
          onEdit={canManage ? () => openEdit(selected) : undefined}
        />
      )}

      {/* ─── Modale de création ────────────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
            <div className="modal-header">
              <h2>{editingId ? 'Modifier le PAC' : "Nouveau Plan d'Action Correctif"}</h2>
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
                  {/* Champ Personne SharePoint → sélecteur Office 365 */}
                  <UserPicker
                    id="pac-resp"
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

              {/* Pièce jointe optionnelle — uploadée après création via Power
                  Automate (workflow PAC_ATTACHMENT_API_URL). Attache le fichier
                  à l'item SharePoint nouvellement créé. */}
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="pac-attachment">Pièce jointe (optionnel)</label>
                <input
                  id="pac-attachment"
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
                    : (editingId ? 'Enregistrer' : 'Créer le PAC')}
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
 * Affiche tous les champs du PAC + ses pièces jointes. Si l'utilisateur a
 * les droits manager (cf. canManage côté parent), le prop `onEdit` est
 * fourni et un bouton "Modifier" est affiché en header pour ouvrir le
 * formulaire en mode édition.
 * ══════════════════════════════════════════════════════════════════════════ */

interface PacDetailModalProps {
  pac: Pac
  onClose: () => void
  /** Callback pour basculer en mode édition (undefined = bouton masqué). */
  onEdit?: () => void
}

function PacDetailModal({ pac, onClose, onEdit }: PacDetailModalProps) {
  // Escape ferme la modale (cohérent avec les autres modales du projet)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const attachments = pac.attachments ?? []

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
        <div className="modal-header" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <h2 style={{ flex: 1 }}>Détail du PAC</h2>
          {/* Bouton Modifier : visible uniquement si onEdit fourni (manager). */}
          {onEdit && (
            <button type="button" className="btn-cta btn-cta-detail" onClick={onEdit}>
              ✎ Modifier
            </button>
          )}
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
