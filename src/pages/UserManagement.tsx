/**
 * ============================================================================
 * MODULE — GESTION DES UTILISATEURS (CRUD)
 * ============================================================================
 *
 * Page d'administration des utilisateurs autorisés à se connecter à l'app.
 * Accessible UNIQUEMENT aux Chef_Departement et Directeur (cf. restriction
 * appliquée dans Dashboard.tsx via MANAGER_ROLES + garde de sécurité ici).
 *
 * Fonctionnalités :
 *   - Liste des utilisateurs (table paginée + filtres recherche/rôle)
 *   - Création d'un utilisateur (modale) — picker Office 365 qui alimente
 *     automatiquement nom + email
 *   - Modification d'un utilisateur existant (même modale, mode édition)
 *
 * Suppression : VOLONTAIREMENT NON EXPOSÉE dans l'UI. Un utilisateur révoqué
 * doit être supprimé via SharePoint directement (sinon risque de perte de
 * traçabilité — un user supprimé qui est responsable d'anomalies / PAC
 * créerait des "orphelins"). Pour révoquer un accès : changer son rôle vers
 * une valeur non-allowée, ou retirer la ligne en SP avec les bonnes ACL.
 *
 * Persistance : liste SharePoint DCPO_LISTE_USER via le service généré.
 *
 * Note sur les RÔLES :
 *   Le champ `fonction` est un Choice SharePoint dont les valeurs autorisées
 *   sont définies CÔTÉ SP (Chef_Departement / Directeur / Controleur). Ces
 *   valeurs sont figées dans la constante ROLE_OPTIONS ci-dessous. Pour
 *   ajouter un nouveau rôle : (1) l'ajouter dans la colonne Choice SP, puis
 *   (2) l'ajouter ici dans ROLE_OPTIONS + ROLE_LABELS.
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import { DCPO_LISTE_USERService } from '../generated/services/DCPO_LISTE_USERService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { DCPO_LISTE_USERRead, DCPO_LISTE_USERWrite } from '../generated/models/DCPO_LISTE_USERModel'
import type { User } from '../generated/models/Office365UsersModel'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'


/**
 * Props passées par Dashboard.tsx.
 * - userEmail : permet d'identifier l'utilisateur courant pour l'empêcher
 *               de se supprimer lui-même (garde-fou UX).
 * - userRole  : rôle EFFECTIF — sert à la garde de sécurité côté page
 *               (deuxième barrière en plus de la nav).
 */
interface UserManagementProps {
  userEmail?: string
  userRole?: string
}

/** Rôles disponibles (doivent matcher les valeurs Choice côté SharePoint). */
const ROLE_OPTIONS = ['Chef_Departement', 'Directeur', 'Controleur'] as const
type RoleOption = typeof ROLE_OPTIONS[number]

/** Libellés lisibles pour l'affichage humain des rôles. */
const ROLE_LABELS: Record<string, string> = {
  Chef_Departement: 'Chef de département',
  Directeur: 'Directeur',
  Controleur: 'Contrôleur',
}

/** Classe CSS de la pastille de rôle (réutilise les manager-pill existantes). */
function getRolePillClass(role: string): string {
  switch (role) {
    case 'Directeur': return 'manager-pill-realise'
    case 'Chef_Departement': return 'manager-pill-pending'
    case 'Controleur': return 'manager-pill-reporte'
    default: return 'manager-pill'
  }
}

/** Liste fermée des rôles ayant le droit d'accéder à ce module. */
const ADMIN_ROLES = ['Chef_Departement', 'Directeur']

/** État des filtres (vide = pas de filtre). */
interface FilterState {
  search: string
  fonction: '' | RoleOption
}
const EMPTY_FILTERS: FilterState = { search: '', fonction: '' }

/** État du formulaire (création OU édition selon `editingId`). */
interface UserFormState {
  nom: string
  email: string
  fonction: RoleOption | ''
}
const EMPTY_FORM: UserFormState = { nom: '', email: '', fonction: '' }

/**
 * Validation basique d'un email (suffisante pour l'UX — la vraie validation
 * sera faite par SharePoint à l'enregistrement).
 */
function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())
}

/** Formate une date SP (ISO) en jj/mm/aaaa. Tolérant : '—' si vide/invalide. */
function formatDate(d: string | undefined): string {
  if (!d) return '—'
  const parsed = new Date(d)
  if (Number.isNaN(parsed.getTime())) return '—'
  return parsed.toLocaleDateString('fr-FR')
}


export default function UserManagement({ userEmail, userRole }: UserManagementProps) {
  /**
   * Garde de sécurité (défense en profondeur) — appliquée dans le rendu, PAS
   * en early-return, pour ne pas violer les règles des hooks React.
   *
   * La nav du Dashboard cache déjà l'onglet aux non-managers ; cette garde
   * sert de filet en cas d'accès par un autre chemin.
   */
  const canAccess = !!userRole && ADMIN_ROLES.includes(userRole)

  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * ════════════════════════════════════════════════════════════════════════ */

  /** Liste complète des utilisateurs depuis SharePoint. */
  const [users, setUsers] = useState<DCPO_LISTE_USERRead[]>([])
  const [loading, setLoading] = useState(true)
  const [loadError, setLoadError] = useState<string | null>(null)

  /** Filtres (saisie + appliqués au clic "Rechercher"). */
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)

  /** Modale création/édition (null = fermée). */
  const [showForm, setShowForm] = useState(false)
  /** ID SP de l'utilisateur en cours d'édition (null = mode création). */
  const [editingId, setEditingId] = useState<string | null>(null)
  /** Snapshot du formulaire en cours de saisie. */
  const [form, setForm] = useState<UserFormState>(EMPTY_FORM)
  const [formError, setFormError] = useState<string | null>(null)
  const [formSaving, setFormSaving] = useState(false)

  /**
   * États du picker de personne Office 365 dans le formulaire.
   *   - peopleSearch : valeur affichée dans l'input (display name)
   *   - peopleResults : résultats Office 365 pour la dropdown
   *   - showPeopleDropdown : visibilité de la dropdown
   *
   * Le pattern est identique à celui des champs auteur/affecté dans Anomalies :
   * tape ≥ 2 caractères → appel SearchUser → click sur un résultat remplit
   * automatiquement nom + email côté form.
   */
  const [peopleSearch, setPeopleSearch] = useState('')
  const [peopleResults, setPeopleResults] = useState<User[]>([])
  const [showPeopleDropdown, setShowPeopleDropdown] = useState(false)


  /* ════════════════════════════════════════════════════════════════════════
   * CHARGEMENT INITIAL + REFRESH
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * Recharge la liste des utilisateurs depuis SharePoint.
   * Tri par nom alphabétique (plus parlant que Created desc pour de l'admin).
   */
  const refresh = async () => {
    setLoading(true)
    setLoadError(null)
    try {
      const result = await DCPO_LISTE_USERService.getAll({ orderBy: ['nom asc'] })
      setUsers(result.data ?? [])
    } catch (err) {
      console.error('UserManagement: échec chargement', err)
      setLoadError('Impossible de charger la liste des utilisateurs.')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    refresh()
  }, [])


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION + STATS
   * ════════════════════════════════════════════════════════════════════════ */

  /** Utilisateurs filtrés selon `appliedFilters` (côté client). */
  const filtered = useMemo(() => {
    const needle = appliedFilters.search.trim().toLowerCase()
    return users.filter(u => {
      if (appliedFilters.fonction && u.fonction?.Value !== appliedFilters.fonction) return false
      if (needle) {
        const haystack = `${u.nom ?? ''} ${u.Email ?? ''} ${u.Title ?? ''}`.toLowerCase()
        if (!haystack.includes(needle)) return false
      }
      return true
    })
  }, [users, appliedFilters])

  /** Stats agrégées (compteurs par rôle). */
  const stats = useMemo(() => ({
    total: filtered.length,
    directeurs: filtered.filter(u => u.fonction?.Value === 'Directeur').length,
    chefs: filtered.filter(u => u.fonction?.Value === 'Chef_Departement').length,
    controleurs: filtered.filter(u => u.fonction?.Value === 'Controleur').length,
  }), [filtered])

  const pagination = usePagination({
    total: filtered.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  const pagedUsers = useMemo(
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
   * HANDLERS — MODALE CRÉATION / ÉDITION
   * ════════════════════════════════════════════════════════════════════════ */

  /** Ouvre le formulaire en mode création (champs vides). */
  const openCreate = () => {
    setEditingId(null)
    setForm(EMPTY_FORM)
    setPeopleSearch('')
    setPeopleResults([])
    setShowPeopleDropdown(false)
    setFormError(null)
    setShowForm(true)
  }

  /** Ouvre le formulaire en mode édition (champs pré-remplis depuis l'item SP). */
  const openEdit = (user: DCPO_LISTE_USERRead) => {
    setEditingId(String(user.ID))
    setForm({
      nom: user.nom ?? '',
      email: user.Email ?? '',
      fonction: (user.fonction?.Value as RoleOption) ?? '',
    })
    // Pré-remplit le picker avec le nom du user (le tag selected-email
    // s'affichera juste en dessous tant que form.email est renseigné).
    setPeopleSearch(user.nom ?? '')
    setPeopleResults([])
    setShowPeopleDropdown(false)
    setFormError(null)
    setShowForm(true)
  }

  const closeForm = () => {
    setShowForm(false)
    setEditingId(null)
    setPeopleSearch('')
    setPeopleResults([])
    setShowPeopleDropdown(false)
    setFormError(null)
  }

  const updateForm = <K extends keyof UserFormState>(key: K, value: UserFormState[K]) => {
    setForm(prev => ({ ...prev, [key]: value }))
  }

  /**
   * Recherche Office 365 — déclenchée à partir de 2 caractères saisis.
   *
   * En cours de frappe, l'input garde sa valeur libre (peopleSearch), mais
   * form.nom / form.email NE SONT PAS modifiés tant qu'aucune personne n'est
   * sélectionnée. Cela évite qu'on valide un formulaire avec un email vide
   * juste parce que l'utilisateur a tapé un nom sans cliquer un résultat.
   */
  const searchPeople = async (term: string) => {
    setPeopleSearch(term)
    if (term.length < 2) {
      setPeopleResults([])
      setShowPeopleDropdown(false)
      return
    }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) {
        setPeopleResults(result.data)
        setShowPeopleDropdown(true)
      }
    } catch (err) {
      console.error('UserManagement: échec recherche Office 365', err)
    }
  }

  /**
   * Sélection d'une personne dans la dropdown → remplit automatiquement
   * nom (DisplayName) et email (Mail) dans le formulaire.
   */
  const selectPerson = (u: User) => {
    const name = u.DisplayName ?? u.Mail ?? ''
    const email = u.Mail ?? ''
    setForm(prev => ({ ...prev, nom: name, email }))
    setPeopleSearch(name)
    setShowPeopleDropdown(false)
    setPeopleResults([])
  }

  /**
   * Soumet le formulaire (création ou édition).
   *
   * Validation locale :
   *   - nom : non vide
   *   - email : format valide ET pas déjà utilisé par un autre utilisateur
   *   - rôle : sélectionné
   *
   * SharePoint :
   *   - Title est réutilisé depuis `nom` (colonne SP obligatoire)
   *   - `fonction` est écrit comme string (le service typage attend une string
   *     côté Write, cf. DCPO_LISTE_USERWrite)
   */
  const submitForm = async () => {
    setFormError(null)
    const nom = form.nom.trim()
    const email = form.email.trim()
    const fonction = form.fonction

    if (!nom) {
      setFormError('Le nom est obligatoire.')
      return
    }
    if (!isValidEmail(email)) {
      setFormError('Email invalide.')
      return
    }
    if (!fonction) {
      setFormError('Le rôle est obligatoire.')
      return
    }
    // Doublon email (toléré uniquement si c'est le même user en édition)
    const dup = users.find(
      u => u.Email?.toLowerCase() === email.toLowerCase() && String(u.ID) !== editingId,
    )
    if (dup) {
      setFormError('Un utilisateur avec cet email existe déjà.')
      return
    }

    setFormSaving(true)
    try {
      /**
       * Format d'écriture du champ Choice `fonction` :
       *
       * Le typage généré (DCPO_LISTE_USERWrite) déclare `fonction?: string`,
       * mais en pratique le connecteur SharePoint persiste le champ Choice
       * SEULEMENT si on lui passe l'objet `{ Value: "..." }`. Un string brut
       * est interprété comme @odata.type et la valeur est perdue → la colonne
       * apparaît vide en lecture.
       *
       * On cast via `unknown` (puis `never` au call site) pour contourner le
       * typage trop restrictif. Pattern identique à celui utilisé pour les
       * champs Personne (cf. confirmAffect dans Anomalies.tsx).
       */
      const payload = {
        Title: nom,
        nom,
        Email: email,
        fonction: { Value: fonction },
      } as unknown as Partial<Omit<DCPO_LISTE_USERWrite, 'ID'>>
      if (editingId) {
        await DCPO_LISTE_USERService.update(editingId, payload)
      } else {
        await DCPO_LISTE_USERService.create(payload as Omit<DCPO_LISTE_USERWrite, 'ID'>)
      }
      closeForm()
      await refresh()
    } catch (err) {
      console.error('UserManagement: échec save', err)
      setFormError("Échec de l'enregistrement. Réessayer.")
    } finally {
      setFormSaving(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   * ════════════════════════════════════════════════════════════════════════ */

  // Garde de sécurité : page muette pour les non-managers (cf. canAccess)
  if (!canAccess) {
    return (
      <>
        <div className="content-header">
          <h2>Gestion des utilisateurs</h2>
        </div>
        <p className="loading-text" style={{ marginTop: 16 }}>
          Accès réservé aux Chefs de département et Directeurs.
        </p>
      </>
    )
  }

  return (
    <>
      {/* ─── Header ────────────────────────────────────────────────── */}
      <div className="content-header">
        <h2>Gestion des utilisateurs</h2>
        <div style={{ display: 'flex', gap: 8 }}>
          <button className="btn-add" type="button" onClick={refresh} disabled={loading}>
            {loading ? 'Chargement...' : 'Actualiser'}
          </button>
          <button className="btn-add" type="button" onClick={openCreate}>
            + Nouvel utilisateur
          </button>
        </div>
      </div>

      {/* ─── Stats cards ───────────────────────────────────────────── */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Total</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{stats.directeurs}</span>
          <span className="stat-label">Directeurs</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.chefs}</span>
          <span className="stat-label">Chefs de département</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{stats.controleurs}</span>
          <span className="stat-label">Contrôleurs</span>
        </div>
      </div>

      {/* ─── Barre de filtres ──────────────────────────────────────── */}
      <div className="filters-bar">
        <div className="filter-field" style={{ flex: 1, minWidth: 220 }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="Nom, email..."
            value={filters.search}
            onChange={e => updateFilter('search', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
          />
        </div>
        <div className="filter-field">
          <label>Rôle</label>
          <select value={filters.fonction} onChange={e => updateFilter('fonction', e.target.value as FilterState['fonction'])}>
            <option value="">Tous</option>
            {ROLE_OPTIONS.map(r => (
              <option key={r} value={r}>{ROLE_LABELS[r]}</option>
            ))}
          </select>
        </div>
        <div className="filter-actions">
          <button type="button" className="btn-search-filters" onClick={applyFilters} disabled={loading}>
            Rechercher
          </button>
          <button type="button" className="btn-reset-filters" onClick={resetFilters}>
            Réinitialiser
          </button>
        </div>
      </div>

      {/* ─── Erreur de chargement ──────────────────────────────────── */}
      {loadError && (
        <p className="loading-text" style={{ color: '#c0392b', marginTop: 16 }} role="alert">
          {loadError}
        </p>
      )}

      {/* ─── Tableau des utilisateurs ──────────────────────────────── */}
      {loading ? (
        <p className="loading-text" style={{ marginTop: 16 }}>Chargement des utilisateurs...</p>
      ) : filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {users.length === 0
            ? 'Aucun utilisateur enregistré pour le moment.'
            : 'Aucun utilisateur ne correspond aux critères.'}
        </p>
      ) : (
        // Wrapper avec scroll horizontal natif. La table utilise
        // `anomalies-table` qui force un layout auto + width auto sur les
        // cellules → chaque colonne dimensionne au contenu, et le wrapper
        // scrolle si l'ensemble dépasse la largeur écran.
        <div className="table-wrapper" style={{ marginTop: 16 }}>
          <table className="data-table anomalies-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Nom</th>
                <th>Email</th>
                <th>Rôle</th>
                <th>Créé le</th>
                <th>Modifié le</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pagedUsers.map((u, i) => {
                const isSelf = !!(u.Email && userEmail && u.Email.toLowerCase() === userEmail.toLowerCase())
                return (
                  <tr key={u.ID}>
                    <td>{pagination.start + i + 1}</td>
                    <td>
                      <strong>{u.nom ?? '—'}</strong>
                      {isSelf && <span style={{ marginLeft: 6, fontSize: 11, color: '#666' }}>(moi)</span>}
                    </td>
                    <td>{u.Email ?? '—'}</td>
                    <td>
                      {u.fonction?.Value ? (
                        <span className={`manager-pill ${getRolePillClass(u.fonction.Value)}`}>
                          {ROLE_LABELS[u.fonction.Value] ?? u.fonction.Value}
                        </span>
                      ) : '—'}
                    </td>
                    <td>{formatDate(u.Created)}</td>
                    <td>{formatDate(u.Modified)}</td>
                    <td>
                      <button
                        type="button"
                        className="btn-cta btn-cta-detail"
                        onClick={() => openEdit(u)}
                      >
                        Modifier
                      </button>
                    </td>
                  </tr>
                )
              })}
            </tbody>
          </table>
          <Pagination
            state={pagination}
            total={filtered.length}
            itemLabel="utilisateurs"
          />
        </div>
      )}

      {/* ─── Modale création / édition ─────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(560px, 100%)' }}>
            <div className="modal-header">
              <h2>{editingId ? "Modifier l'utilisateur" : 'Nouvel utilisateur'}</h2>
              <button className="modal-close" onClick={closeForm} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              {/* Picker Office 365 : un seul champ alimente automatiquement
                  nom (DisplayName) et email (Mail) sur sélection d'un résultat.
                  Le tag .selected-email sous l'input rappelle l'email choisi. */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="user-person">Personne *</label>
                <div className="autocomplete-wrapper">
                  <input
                    id="user-person"
                    type="text"
                    value={peopleSearch}
                    placeholder="Rechercher un utilisateur Office 365..."
                    onChange={e => searchPeople(e.target.value)}
                    onBlur={() => setTimeout(() => setShowPeopleDropdown(false), 200)}
                    disabled={formSaving}
                    autoFocus
                  />
                  {form.email && <span className="selected-email">{form.email}</span>}
                  {showPeopleDropdown && peopleResults.length > 0 && (
                    <ul className="autocomplete-dropdown">
                      {peopleResults.map(u => (
                        <li key={u.Id} onClick={() => selectPerson(u)}>
                          <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="user-fonction">Rôle *</label>
                <select
                  id="user-fonction"
                  value={form.fonction}
                  onChange={e => updateForm('fonction', e.target.value as RoleOption)}
                  disabled={formSaving}
                >
                  <option value="">— Choisir un rôle —</option>
                  {ROLE_OPTIONS.map(r => (
                    <option key={r} value={r}>{ROLE_LABELS[r]}</option>
                  ))}
                </select>
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
                  {formSaving ? 'Enregistrement...' : (editingId ? 'Enregistrer' : "Créer l'utilisateur")}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

    </>
  )
}
