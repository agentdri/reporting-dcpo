/**
 * ============================================================================
 * MODULE — GESTION DES DIRECTIONS (référentiel)
 * ============================================================================
 *
 * Page CRUD sur la liste SharePoint DCPO_LISTE_DIRECTION.
 *
 * Visible uniquement aux managers (Chef_Departement / Directeur) — restriction
 * gérée côté Dashboard.tsx au niveau du menu de navigation.
 *
 * Pourquoi une page dédiée plutôt que la gestion inline dans le PAC ?
 *   - Référentiel partagé : les directions peuvent à terme servir à d'autres
 *     listes (anomalies, plans de contrôle…)
 *   - Séparation des responsabilités : la gestion des ressources de référence
 *     vit dans le menu "Configuration"
 *
 * Pattern réutilisable : ce composant sert de modèle pour de futures pages
 * de référentiel (sigle/libellé + Title — schéma simple).
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import {
  createDirection,
  deleteDirection,
  listDirections,
  updateDirection,
  type Direction,
} from '../lib/directionService'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'


/** État vide du formulaire (création + reset). */
const EMPTY_FORM = { sigle: '', libelle: '' }


/* ──────────────────────────────────────────────────────────────────────────
 * COMPOSANT PRINCIPAL
 * ────────────────────────────────────────────────────────────────────────── */

export default function DirectionManagement() {
  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * ════════════════════════════════════════════════════════════════════════ */

  /** Liste des directions chargées depuis SharePoint. */
  const [directions, setDirections] = useState<Direction[]>([])
  const [loading, setLoading] = useState(true)

  /** Recherche libre (sigle + libellé), appliquée à la frappe (pas de bouton). */
  const [search, setSearch] = useState('')

  /** Formulaire création/édition. null editingId = mode création. */
  const [showForm, setShowForm] = useState(false)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [form, setForm] = useState(EMPTY_FORM)
  const [formError, setFormError] = useState<string | null>(null)
  const [formSaving, setFormSaving] = useState(false)

  /** Confirmation de suppression. null = pas de confirmation en cours. */
  const [pendingDelete, setPendingDelete] = useState<Direction | null>(null)
  const [deleting, setDeleting] = useState(false)


  /* ════════════════════════════════════════════════════════════════════════
   * CHARGEMENT DEPUIS SHAREPOINT
   *
   * Fetch async dans un useEffect : autorisé car le setState a lieu dans un
   * callback après await (pas dans le corps synchrone de l'effet).
   * ════════════════════════════════════════════════════════════════════════ */

  useEffect(() => {
    let cancelled = false
    listDirections()
      .then(data => { if (!cancelled) setDirections(data) })
      .finally(() => { if (!cancelled) setLoading(false) })
    return () => { cancelled = true }
  }, [])

  /** Recharge la liste après une mutation (créa / édit / suppr). */
  const refresh = async () => {
    setLoading(true)
    try {
      setDirections(await listDirections())
    } finally {
      setLoading(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION
   * ════════════════════════════════════════════════════════════════════════ */

  /** Filtrage côté client sur sigle + libellé (recherche immédiate). */
  const filtered = useMemo(() => {
    const needle = search.trim().toLowerCase()
    if (!needle) return directions
    return directions.filter(d =>
      d.sigle.toLowerCase().includes(needle) ||
      d.libelle.toLowerCase().includes(needle),
    )
  }, [directions, search])

  /**
   * Pagination — porte sur la liste filtrée. resetKey suit la recherche
   * pour ramener à la page 1 quand l'utilisateur tape.
   */
  const pagination = usePagination({
    total: filtered.length,
    resetKey: search,
  })
  const pagedDirections = useMemo(
    () => filtered.slice(pagination.start, pagination.end),
    [filtered, pagination.start, pagination.end],
  )


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS — FORMULAIRE (création + édition)
   * ════════════════════════════════════════════════════════════════════════ */

  const openCreate = () => {
    setEditingId(null)
    setForm(EMPTY_FORM)
    setFormError(null)
    setShowForm(true)
  }

  const openEdit = (d: Direction) => {
    setEditingId(d.id)
    setForm({ sigle: d.sigle, libelle: d.libelle })
    setFormError(null)
    setShowForm(true)
  }

  const closeForm = () => {
    setShowForm(false)
    setEditingId(null)
    setFormError(null)
  }

  /**
   * Crée ou met à jour selon `editingId`.
   *
   * Validation locale minimale (les services valident aussi) :
   *   - sigle non vide
   *   - libellé non vide
   *
   * Contrôle d'unicité sur le sigle : si on tente d'utiliser un sigle déjà
   * pris par une AUTRE direction (id différent), on bloque côté front pour
   * éviter d'avoir deux directions avec le même code (sinon les références
   * stockées dans les PAC deviennent ambiguës).
   */
  const submitForm = async () => {
    setFormError(null)
    const sigle = form.sigle.trim()
    const libelle = form.libelle.trim()
    if (!sigle) {
      setFormError('Le sigle est obligatoire.')
      return
    }
    if (!libelle) {
      setFormError('Le libellé est obligatoire.')
      return
    }
    // Unicité du sigle (case-insensitive)
    const conflict = directions.find(d => d.sigle.toLowerCase() === sigle.toLowerCase() && d.id !== editingId)
    if (conflict) {
      setFormError(`Le sigle "${sigle}" est déjà utilisé par : ${conflict.libelle}`)
      return
    }

    setFormSaving(true)
    try {
      if (editingId) {
        await updateDirection(editingId, { sigle, libelle })
      } else {
        await createDirection({ sigle, libelle })
      }
      await refresh()
      closeForm()
    } catch (err) {
      console.error('submitForm direction error', err)
      setFormError(err instanceof Error ? err.message : 'Échec de l\'enregistrement.')
    } finally {
      setFormSaving(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS — SUPPRESSION
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * Demande la suppression d'une direction.
   *
   * On ne contrôle PAS ici si la direction est encore référencée par des
   * PAC (le scan serait coûteux et peu fiable côté client). Le manager doit
   * vérifier en amont — c'est volontairement permissif pour les corrections
   * de typos / nettoyage.
   */
  const askDelete = (d: Direction) => {
    setPendingDelete(d)
  }

  const confirmDelete = async () => {
    if (!pendingDelete?.id) return
    setDeleting(true)
    try {
      await deleteDirection(pendingDelete.id)
      setPendingDelete(null)
      await refresh()
    } catch (err) {
      console.error('confirmDelete direction error', err)
      alert('Échec de la suppression. Réessayer.')
    } finally {
      setDeleting(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   * ════════════════════════════════════════════════════════════════════════ */

  return (
    <>
      {/* ─── Header + bouton + nouvelle ─────────────────────────────── */}
      <div className="content-header">
        <h2>Directions</h2>
        <button className="btn-add" type="button" onClick={openCreate}>
          + Nouvelle direction
        </button>
      </div>

      {/* ─── Barre de recherche ────────────────────────────────────── */}
      <div className="filters-bar">
        <div className="filter-field" style={{ flex: 1, minWidth: 220 }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="Sigle ou libellé..."
            value={search}
            onChange={e => setSearch(e.target.value)}
          />
        </div>
      </div>

      {/* ─── Tableau ───────────────────────────────────────────────── */}
      {loading ? (
        <p className="loading-text" style={{ marginTop: 16 }}>Chargement des directions...</p>
      ) : filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {directions.length === 0
            ? 'Aucune direction enregistrée. Cliquer sur "Nouvelle direction" pour en ajouter.'
            : 'Aucune direction ne correspond à la recherche.'}
        </p>
      ) : (
        // Wrapper avec scroll horizontal natif si le tableau dépasse l'écran.
        <div className="table-wrapper" style={{ marginTop: 16 }}>
          <table className="data-table anomalies-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Sigle</th>
                <th>Libellé</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pagedDirections.map((d, i) => (
                <tr key={d.id}>
                  <td>{pagination.start + i + 1}</td>
                  <td><strong>{d.sigle}</strong></td>
                  <td>{d.libelle}</td>
                  <td>
                    <div style={{ display: 'flex', gap: 6 }}>
                      <button
                        type="button"
                        className="btn-cta btn-cta-detail"
                        onClick={() => openEdit(d)}
                      >
                        Modifier
                      </button>
                      <button
                        type="button"
                        className="btn-cta btn-cta-affect"
                        onClick={() => askDelete(d)}
                      >
                        Supprimer
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <Pagination
            state={pagination}
            total={filtered.length}
            itemLabel="directions"
          />
        </div>
      )}

      {/* ─── Modale création / édition ─────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(520px, 100%)' }}>
            <div className="modal-header">
              <h2>{editingId ? 'Modifier la direction' : 'Nouvelle direction'}</h2>
              <button className="modal-close" onClick={closeForm} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="dir-sigle">Sigle *</label>
                <input
                  id="dir-sigle"
                  type="text"
                  value={form.sigle}
                  onChange={e => setForm(prev => ({ ...prev, sigle: e.target.value }))}
                  placeholder="Ex: DSI, DJC, DCE..."
                  disabled={formSaving}
                  autoFocus
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="dir-libelle">Libellé complet *</label>
                <input
                  id="dir-libelle"
                  type="text"
                  value={form.libelle}
                  onChange={e => setForm(prev => ({ ...prev, libelle: e.target.value }))}
                  placeholder="Ex: Direction des Systèmes d'Information"
                  disabled={formSaving}
                />
              </div>

              {formError && (
                <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0' }} role="alert">
                  {formError}
                </p>
              )}

              <div className="modal-actions" style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button
                  type="button"
                  className="btn-cta btn-cta-detail"
                  onClick={closeForm}
                  disabled={formSaving}
                >
                  Annuler
                </button>
                <button
                  type="button"
                  className="btn-cta btn-cta-affect"
                  onClick={submitForm}
                  disabled={formSaving}
                >
                  {formSaving ? 'Enregistrement...' : (editingId ? 'Enregistrer' : 'Créer')}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ─── Modale de confirmation suppression ────────────────────── */}
      {pendingDelete && (
        <div className="modal-overlay" onClick={() => !deleting && setPendingDelete(null)}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(440px, 100%)' }}>
            <div className="modal-header">
              <h2>Confirmer la suppression</h2>
              <button className="modal-close" onClick={() => !deleting && setPendingDelete(null)} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              <p style={{ marginTop: 0 }}>
                Êtes-vous sûr de vouloir supprimer cette direction ?
              </p>
              <dl className="detail-grid" style={{ margin: '8px 0 12px' }}>
                <dt>Sigle</dt><dd><strong>{pendingDelete.sigle}</strong></dd>
                <dt>Libellé</dt><dd>{pendingDelete.libelle}</dd>
              </dl>
              <p style={{ fontSize: 12, color: '#c0392b', margin: 0 }}>
                Cette action est définitive. Si des PAC référencent encore ce sigle,
                ils continueront d'afficher le code brut sans son libellé.
              </p>
              <div className="modal-actions" style={{ marginTop: 16, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button
                  type="button"
                  className="btn-cta btn-cta-detail"
                  onClick={() => setPendingDelete(null)}
                  disabled={deleting}
                >
                  Annuler
                </button>
                <button
                  type="button"
                  className="btn-cta btn-cta-affect"
                  onClick={confirmDelete}
                  disabled={deleting}
                >
                  {deleting ? 'Suppression...' : 'Confirmer la suppression'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  )
}
