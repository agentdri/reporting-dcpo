/**
 * ============================================================================
 * GESTION DES ABSENCES — Configuration (managers uniquement)
 * ============================================================================
 *
 * CRUD complet sur la liste `DCPO_LISTE_ABSENCES`. Une absence enregistrée
 * ici est automatiquement retirée du dénominateur du taux de soumission
 * des rapports quotidiens (cf. ControllerReportingList.tsx).
 *
 * Périmètre :
 *   - Accès : Chef_Departement / Directeur (contrôlé côté Dashboard.tsx
 *     via isManager avant de proposer l'onglet dans le menu Configuration)
 *   - Actions : Créer / Modifier / Supprimer une absence
 *   - Filtres : Contrôleur (nom/email), Motif, Période active
 *
 * Modèle métier : voir `src/lib/absenceService.ts`.
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import {
  listAbsences,
  createAbsence,
  updateAbsence,
  deleteAbsence,
  ABSENCE_MOTIFS,
  type Absence,
} from '../lib/absenceService'
import { UserPicker } from '../components/UserPicker'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { ExportButtons } from '../components/ExportButtons'
import { formatDateForExport } from '../lib/exporters'
import { ModalOverlay } from '../components/ModalOverlay'

/** Props reçues du Dashboard : identité du manager (pour renseigner declarePar). */
interface AbsenceManagementProps {
  userName?: string
  userEmail?: string
}

/** Filtres de la vue liste — pattern saisie/applied comme les autres pages. */
interface FilterState {
  controleur: string    // recherche partielle nom OU email
  motif: string         // '' = tous
  activeOnly: boolean   // ne montrer que les absences ACTIVES (dateFin >= aujourd'hui)
}

const EMPTY_FILTERS: FilterState = {
  controleur: '',
  motif: '',
  activeOnly: false,
}

/** État initial du formulaire (création ou édition). */
interface FormState {
  controleurName: string
  controleurEmail: string
  dateDebut: string    // YYYY-MM-DD
  dateFin: string      // YYYY-MM-DD
  motif: string
  commentaire: string
}

const EMPTY_FORM: FormState = {
  controleurName: '',
  controleurEmail: '',
  dateDebut: '',
  dateFin: '',
  motif: 'Congé annuel',
  commentaire: '',
}

/** Formatte YYYY-MM-DD → JJ/MM/AAAA (locale FR). */
function formatDateFR(iso: string): string {
  if (!iso) return '—'
  const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})/)
  return m ? `${m[3]}/${m[2]}/${m[1]}` : iso
}

/**
 * Vérifie qu'une absence est "active" à une date donnée = la date est
 * dans l'intervalle [dateDebut, dateFin] (inclusifs).
 */
function isAbsenceActive(abs: Absence, ref: Date): boolean {
  if (!abs.dateDebut || !abs.dateFin) return false
  const refIso = ref.toISOString().slice(0, 10)
  return abs.dateDebut <= refIso && refIso <= abs.dateFin
}

export default function AbsenceManagement({
  userName,
  userEmail,
}: AbsenceManagementProps) {
  /* ─── ÉTATS ────────────────────────────────────────────────────────── */

  const [absences, setAbsences] = useState<Absence[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  // Filtres — saisie vs appliqué
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)

  // Formulaire de saisie (modale)
  const [showForm, setShowForm] = useState(false)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [form, setForm] = useState<FormState>(EMPTY_FORM)
  const [saving, setSaving] = useState(false)
  const [formError, setFormError] = useState<string | null>(null)

  // Confirmation de suppression
  const [deletingId, setDeletingId] = useState<string | null>(null)


  /* ─── CHARGEMENT ───────────────────────────────────────────────────── */

  const refresh = async () => {
    setLoading(true)
    setError(null)
    try {
      setAbsences(await listAbsences())
    } catch (err) {
      console.error('AbsenceManagement.refresh error', err)
      setError('Impossible de charger les absences.')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    refresh()
  }, [])


  /* ─── DÉRIVATIONS (filtre + pagination + stats) ────────────────────── */

  const filtered = useMemo(() => {
    const search = appliedFilters.controleur.trim().toLowerCase()
    const today = new Date()
    return absences.filter(a => {
      if (search) {
        const haystack = `${a.controleurName} ${a.controleurEmail}`.toLowerCase()
        if (!haystack.includes(search)) return false
      }
      if (appliedFilters.motif && a.motif !== appliedFilters.motif) return false
      if (appliedFilters.activeOnly && !isAbsenceActive(a, today)) return false
      return true
    })
  }, [absences, appliedFilters])

  const pagination = usePagination({
    total: filtered.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  const paged = useMemo(
    () => filtered.slice(pagination.start, pagination.end),
    [filtered, pagination.start, pagination.end],
  )

  const stats = useMemo(() => {
    const today = new Date()
    const active = filtered.filter(a => isAbsenceActive(a, today)).length
    const future = filtered.filter(a => a.dateDebut > today.toISOString().slice(0, 10)).length
    return { total: filtered.length, active, future }
  }, [filtered])


  /* ─── HANDLERS ─────────────────────────────────────────────────────── */

  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const applyFilters = () => setAppliedFilters(filters)
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS)
    setAppliedFilters(EMPTY_FILTERS)
  }

  const updateForm = <K extends keyof FormState>(key: K, value: FormState[K]) => {
    setForm(prev => ({ ...prev, [key]: value }))
  }

  const openCreate = () => {
    setEditingId(null)
    setForm(EMPTY_FORM)
    setFormError(null)
    setShowForm(true)
  }

  const openEdit = (abs: Absence) => {
    setEditingId(abs.id)
    setForm({
      controleurName: abs.controleurName,
      controleurEmail: abs.controleurEmail,
      dateDebut: abs.dateDebut,
      dateFin: abs.dateFin,
      motif: abs.motif || 'Congé annuel',
      commentaire: abs.commentaire ?? '',
    })
    setFormError(null)
    setShowForm(true)
  }

  const closeForm = () => {
    if (saving) return
    setShowForm(false)
    setEditingId(null)
    setForm(EMPTY_FORM)
    setFormError(null)
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    // Validations côté client
    if (!form.controleurEmail.trim()) {
      setFormError('Sélectionnez un contrôleur.')
      return
    }
    if (!form.dateDebut || !form.dateFin) {
      setFormError('Renseignez les dates de début et de fin.')
      return
    }
    if (form.dateFin < form.dateDebut) {
      setFormError('La date de fin doit être postérieure ou égale à la date de début.')
      return
    }
    if (!ABSENCE_MOTIFS.includes(form.motif as never)) {
      setFormError('Motif invalide.')
      return
    }
    if (!userEmail) {
      setFormError('Session invalide : impossible de renseigner "declaré par".')
      return
    }

    setSaving(true)
    setFormError(null)
    try {
      if (editingId) {
        const updated = await updateAbsence(editingId, {
          controleurEmail: form.controleurEmail,
          controleurName: form.controleurName,
          dateDebut: form.dateDebut,
          dateFin: form.dateFin,
          motif: form.motif,
          commentaire: form.commentaire,
          declareParEmail: userEmail,
          declareParName: userName ?? '',
        })
        if (!updated) throw new Error('update failed')
      } else {
        const created = await createAbsence({
          controleurEmail: form.controleurEmail,
          controleurName: form.controleurName,
          dateDebut: form.dateDebut,
          dateFin: form.dateFin,
          motif: form.motif,
          commentaire: form.commentaire,
          declareParEmail: userEmail,
          declareParName: userName ?? '',
        })
        if (!created) throw new Error('create failed')
      }
      await refresh()
      closeForm()
    } catch (err) {
      console.error('AbsenceManagement.save error', err)
      setFormError('Échec de la sauvegarde.')
    } finally {
      setSaving(false)
    }
  }

  const confirmDelete = async () => {
    if (!deletingId) return
    setSaving(true)
    try {
      const ok = await deleteAbsence(deletingId)
      if (ok) {
        await refresh()
      } else {
        alert('Suppression échouée.')
      }
    } finally {
      setSaving(false)
      setDeletingId(null)
    }
  }


  /* ─── RENDU ────────────────────────────────────────────────────────── */

  return (
    <>
      {/* ─── HEADER : titre + exports + bouton "+ Nouvelle absence" ──── */}
      <div className="content-header">
        <h2>Gestion des absences</h2>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <ExportButtons
            filename="absences_controleurs"
            pdfTitle="Absences des contrôleurs"
            getHeaders={() => [
              'ID', 'Contrôleur', 'Email', 'Date début', 'Date fin', 'Motif',
              'Commentaire', 'Déclarée par', 'Créée le',
            ]}
            getRows={() => filtered.map(a => [
              a.id,
              a.controleurName,
              a.controleurEmail,
              formatDateForExport(a.dateDebut),
              formatDateForExport(a.dateFin),
              a.motif,
              a.commentaire ?? '',
              a.declareParName ?? a.declareParEmail ?? '',
              formatDateForExport(a.createdAt),
            ])}
            disabled={loading}
          />
          <button className="btn-add" type="button" onClick={openCreate} disabled={loading}>
            + Nouvelle absence
          </button>
        </div>
      </div>

      {/* ─── STATS CARDS ──────────────────────────────────────────────── */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Total absences</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.active}</span>
          <span className="stat-label">Actives aujourd'hui</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{stats.future}</span>
          <span className="stat-label">Futures</span>
        </div>
      </div>

      {/* ─── FILTRES ──────────────────────────────────────────────────── */}
      <div className="filters-bar">
        <div className="filter-field" style={{ flex: '1 1 220px' }}>
          <label>Contrôleur</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filters.controleur}
            onChange={e => updateFilter('controleur', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
          />
        </div>
        <div className="filter-field">
          <label>Motif</label>
          <select value={filters.motif} onChange={e => updateFilter('motif', e.target.value)}>
            <option value="">Tous</option>
            {ABSENCE_MOTIFS.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
        <label className="filter-toggle" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <input
            type="checkbox"
            checked={filters.activeOnly}
            onChange={e => updateFilter('activeOnly', e.target.checked)}
          />
          Actives uniquement
        </label>
        <button type="button" className="btn-search-filters" onClick={applyFilters} disabled={loading}>
          Rechercher
        </button>
        <button type="button" className="btn-reset-filters" onClick={resetFilters}>
          Réinitialiser
        </button>
      </div>

      {/* ─── TABLEAU ──────────────────────────────────────────────────── */}
      {error && <div className="bulletin-error" role="alert">{error}</div>}

      {loading ? (
        <p className="loading-text">Chargement des absences...</p>
      ) : filtered.length === 0 ? (
        <p className="loading-text">Aucune absence enregistrée pour ces filtres.</p>
      ) : (
        <div className="table-wrapper">
          <table className="data-table">
            <thead>
              <tr>
                <th>Contrôleur</th>
                <th>Email</th>
                <th>Date début</th>
                <th>Date fin</th>
                <th>Motif</th>
                <th>Commentaire</th>
                <th>Déclarée par</th>
                <th style={{ width: 160, textAlign: 'center' }}>Actions</th>
              </tr>
            </thead>
            <tbody>
              {paged.map(abs => (
                <tr key={abs.id}>
                  <td><strong>{abs.controleurName}</strong></td>
                  <td style={{ color: '#64748b', fontSize: 12 }}>{abs.controleurEmail}</td>
                  <td>{formatDateFR(abs.dateDebut)}</td>
                  <td>{formatDateFR(abs.dateFin)}</td>
                  <td>
                    <span
                      style={{
                        fontSize: 11,
                        padding: '2px 8px',
                        borderRadius: 10,
                        background: '#e0e7ff',
                        color: '#3730a3',
                      }}
                    >
                      {abs.motif}
                    </span>
                  </td>
                  <td style={{ fontSize: 12, color: '#475569' }}>
                    {abs.commentaire || '—'}
                  </td>
                  <td style={{ fontSize: 12, color: '#475569' }}>
                    {abs.declareParName ?? '—'}
                  </td>
                  <td style={{ textAlign: 'center' }}>
                    <button
                      type="button"
                      className="btn-detail"
                      onClick={() => openEdit(abs)}
                      style={{ marginRight: 6 }}
                    >
                      Modifier
                    </button>
                    <button
                      type="button"
                      className="btn-affect"
                      style={{ background: '#c0392b' }}
                      onClick={() => setDeletingId(abs.id)}
                    >
                      Supprimer
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <Pagination state={pagination} total={filtered.length} itemLabel="absences" />
        </div>
      )}

      {/* ─── MODALE CREATE / EDIT ─────────────────────────────────────── */}
      {showForm && (
        <ModalOverlay onClose={closeForm}>
          <div
            className="modal"
            onClick={e => e.stopPropagation()}
            role="dialog"
            aria-modal="true"
            style={{ maxWidth: 560 }}
          >
            <header className="modal-header">
              <h2>{editingId ? "Modifier l'absence" : 'Nouvelle absence'}</h2>
              <button className="modal-close" onClick={closeForm} disabled={saving} aria-label="Fermer">
                &times;
              </button>
            </header>
            <form onSubmit={handleSubmit} style={{ padding: 16 }}>
              {formError && (
                <div className="bulletin-error" role="alert" style={{ marginBottom: 12 }}>
                  {formError}
                </div>
              )}

              <div className="form-field" style={{ marginBottom: 12 }}>
                <label>Contrôleur *</label>
                <UserPicker
                  selectedName={form.controleurName}
                  selectedEmail={form.controleurEmail}
                  onSelect={(name, email) => {
                    updateForm('controleurName', name)
                    updateForm('controleurEmail', email)
                  }}
                  onClear={() => {
                    updateForm('controleurName', '')
                    updateForm('controleurEmail', '')
                  }}
                  placeholder="Rechercher un contrôleur..."
                  disabled={saving}
                />
              </div>

              <div className="form-grid" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12, marginBottom: 12 }}>
                <div className="form-field">
                  <label htmlFor="abs-debut">Date de début *</label>
                  <input
                    id="abs-debut"
                    type="date"
                    value={form.dateDebut}
                    onChange={e => updateForm('dateDebut', e.target.value)}
                    disabled={saving}
                    required
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="abs-fin">Date de fin *</label>
                  <input
                    id="abs-fin"
                    type="date"
                    value={form.dateFin}
                    onChange={e => updateForm('dateFin', e.target.value)}
                    disabled={saving}
                    required
                  />
                </div>
              </div>

              <div className="form-field" style={{ marginBottom: 12 }}>
                <label htmlFor="abs-motif">Motif *</label>
                <select
                  id="abs-motif"
                  value={form.motif}
                  onChange={e => updateForm('motif', e.target.value)}
                  disabled={saving}
                >
                  {ABSENCE_MOTIFS.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
              </div>

              <div className="form-field" style={{ marginBottom: 16 }}>
                <label htmlFor="abs-comment">Commentaire</label>
                <textarea
                  id="abs-comment"
                  rows={3}
                  value={form.commentaire}
                  onChange={e => updateForm('commentaire', e.target.value)}
                  disabled={saving}
                  placeholder="Précisions optionnelles (n° dossier RH, ordre de mission...)"
                />
              </div>

              <footer style={{ display: 'flex', justifyContent: 'flex-end', gap: 8 }}>
                <button
                  type="button"
                  className="btn-cta btn-cta-ghost"
                  onClick={closeForm}
                  disabled={saving}
                >
                  Annuler
                </button>
                <button
                  type="submit"
                  className="btn-cta btn-cta-primary"
                  disabled={saving}
                >
                  {saving ? 'Enregistrement…' : editingId ? 'Enregistrer' : 'Créer'}
                </button>
              </footer>
            </form>
          </div>
        </ModalOverlay>
      )}

      {/* ─── MODALE CONFIRMATION SUPPRESSION ──────────────────────────── */}
      {deletingId && (
        <ModalOverlay onClose={() => (!saving && setDeletingId(null))}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ maxWidth: 460 }}>
            <header className="modal-header">
              <h2>Supprimer l'absence ?</h2>
            </header>
            <div style={{ padding: 16 }}>
              <p>
                Cette absence sera définitivement supprimée. Les jours correspondants
                redeviendront <strong>attendus</strong> dans le calcul du taux
                de soumission du contrôleur.
              </p>
              <footer style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 16 }}>
                <button
                  type="button"
                  className="btn-cta btn-cta-ghost"
                  onClick={() => setDeletingId(null)}
                  disabled={saving}
                >
                  Annuler
                </button>
                <button
                  type="button"
                  className="btn-cta btn-cta-primary"
                  style={{ background: '#c0392b' }}
                  onClick={confirmDelete}
                  disabled={saving}
                >
                  {saving ? 'Suppression…' : 'Supprimer'}
                </button>
              </footer>
            </div>
          </div>
        </ModalOverlay>
      )}
    </>
  )
}
