/**
 * ============================================================================
 * COMPOSANT — <UserPicker />
 * ============================================================================
 *
 * Sélecteur d'utilisateur Office 365 réutilisable (autocomplete).
 *
 * Utilisé pour renseigner les champs de type "Personne" des listes SharePoint
 * (ex: responsable du plan de contrôle, responsable de mise en œuvre d'un PAC).
 *
 * Comportement :
 *   - Tant qu'aucun utilisateur n'est sélectionné : champ de recherche +
 *     liste déroulante des résultats (recherche dès 2 caractères)
 *   - Une fois sélectionné : affichage du nom + email + bouton "Changer"
 *
 * Le parent reçoit le DisplayName ET l'email via onSelect (l'email sert à
 * construire le format Claims pour l'écriture SharePoint).
 * ============================================================================
 */

import { useState } from 'react'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { User } from '../generated/models/Office365UsersModel'

export interface UserPickerProps {
  /** Nom affiché de l'utilisateur actuellement sélectionné (vide si aucun). */
  selectedName: string
  /** Email de l'utilisateur sélectionné (vide si aucun). */
  selectedEmail: string
  /** Appelé quand l'utilisateur choisit une personne dans la liste. */
  onSelect: (name: string, email: string) => void
  /** Appelé quand l'utilisateur efface la sélection (bouton "Changer"). */
  onClear: () => void
  /** Placeholder du champ de recherche. */
  placeholder?: string
  /** id du champ (pour lier un <label htmlFor>). */
  id?: string
  /** Désactive les interactions (pendant une sauvegarde par ex.). */
  disabled?: boolean
}

export function UserPicker({
  selectedName,
  selectedEmail,
  onSelect,
  onClear,
  placeholder = 'Rechercher une personne...',
  id,
  disabled = false,
}: UserPickerProps) {
  const [search, setSearch] = useState('')
  const [results, setResults] = useState<User[]>([])
  const [loading, setLoading] = useState(false)

  /** Recherche Office 365 (déclenchée dès 2 caractères). */
  const doSearch = async (term: string) => {
    setSearch(term)
    if (term.trim().length < 2) {
      setResults([])
      return
    }
    setLoading(true)
    try {
      const res = await Office365UsersService.SearchUser(term, 10)
      if (res.data) setResults(res.data)
    } catch (err) {
      console.error('UserPicker: erreur recherche utilisateur', err)
    } finally {
      setLoading(false)
    }
  }

  const handleSelect = (u: User) => {
    onSelect(u.DisplayName ?? u.Mail ?? '', u.Mail ?? '')
    setSearch('')
    setResults([])
  }

  // ─── Mode "sélectionné" : récap + bouton Changer ──────────────────────
  if (selectedName || selectedEmail) {
    return (
      <div className="user-picker-selected">
        <div>
          <strong>{selectedName || selectedEmail}</strong>
          {selectedEmail && selectedName && (
            <div className="user-picker-email">{selectedEmail}</div>
          )}
        </div>
        <button
          type="button"
          className="btn-cta btn-cta-detail"
          onClick={onClear}
          disabled={disabled}
        >
          Changer
        </button>
      </div>
    )
  }

  // ─── Mode "recherche" : input + dropdown ──────────────────────────────
  return (
    <div className="autocomplete-wrapper" style={{ position: 'relative' }}>
      <input
        id={id}
        type="text"
        value={search}
        placeholder={placeholder}
        onChange={e => doSearch(e.target.value)}
        disabled={disabled}
        autoComplete="off"
      />
      {loading && <span className="user-picker-hint">Recherche…</span>}
      {results.length > 0 && (
        <ul className="autocomplete-dropdown">
          {results.map(u => (
            <li key={u.Id} onClick={() => handleSelect(u)}>
              <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
