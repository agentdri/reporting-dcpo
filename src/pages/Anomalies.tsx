import { useState, useEffect } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { DCPO_LISTE_ANORMALIERead, DCPO_LISTE_ANORMALIEWrite } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import type { User } from '../generated/models/Office365UsersModel'

interface AnomaliesProps {
  userName?: string
  userEmail?: string
}

interface FormState {
  field_0: string
  field_4: string
  field_5: string
  field_6: string
  field_7: string
  field_8: number
  field_9: string
  field_10: string
}

const EMPTY_FORM: FormState = {
  field_0: '',
  field_4: '',
  field_5: '',
  field_6: '',
  field_7: '',
  field_8: 0,
  field_9: '',
  field_10: '',
}

const FIELD_LABELS: Record<string, string> = {
  field_0: 'Date',
  field_4: 'Cause',
  field_5: 'Classification',
  field_6: 'Agence',
  field_7: 'Reseau',
  field_8: 'Montant',
  field_9: 'Date de regularisation',
  field_10: 'Statut',
}

const DATE_FIELDS = ['field_0', 'field_9']

function stripHtml(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  return doc.body.textContent?.trim() ?? ''
}

function toClaims(email: string) {
  return `i:0#.f|membership|${email}`
}

export default function Anomalies({ userName, userEmail }: AnomaliesProps) {
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [loading, setLoading] = useState(true)
  const [showForm, setShowForm] = useState(false)
  const [form, setForm] = useState(EMPTY_FORM)
  const [submitting, setSubmitting] = useState(false)

  // Detail modale
  const [detailItem, setDetailItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)

  // Affectation modale
  const [affectItemId, setAffectItemId] = useState<number | null>(null)
  const [affectSearch, setAffectSearch] = useState('')
  const [affectResults, setAffectResults] = useState<User[]>([])
  const [affectLoading, setAffectLoading] = useState(false)

  // Agences & Reseaux
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])

  // Auteur search
  const [auteurSearch, setAuteurSearch] = useState('')
  const [auteurEmail, setAuteurEmail] = useState('')
  const [auteurResults, setAuteurResults] = useState<User[]>([])
  const [showAuteurDropdown, setShowAuteurDropdown] = useState(false)

  // Personne affectee search
  const [affecteFormSearch, setAffecteFormSearch] = useState('')
  const [affecteFormEmail, setAffecteFormEmail] = useState('')
  const [affecteFormResults, setAffecteFormResults] = useState<User[]>([])
  const [showAffecteFormDropdown, setShowAffecteFormDropdown] = useState(false)

  const searchAuteur = async (term: string) => {
    setAuteurSearch(term)
    if (term.length < 2) { setAuteurResults([]); setShowAuteurDropdown(false); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) { setAuteurResults(result.data); setShowAuteurDropdown(true) }
    } catch (err) { console.error('Erreur recherche auteur', err) }
  }

  const selectAuteur = (u: User) => {
    setAuteurEmail(u.Mail ?? '')
    setAuteurSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAuteurDropdown(false)
    setAuteurResults([])
  }

  const searchAffecteForm = async (term: string) => {
    setAffecteFormSearch(term)
    if (term.length < 2) { setAffecteFormResults([]); setShowAffecteFormDropdown(false); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) { setAffecteFormResults(result.data); setShowAffecteFormDropdown(true) }
    } catch (err) { console.error('Erreur recherche personne affectee', err) }
  }

  const selectAffecteForm = (u: User) => {
    setAffecteFormEmail(u.Mail ?? '')
    setAffecteFormSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAffecteFormDropdown(false)
    setAffecteFormResults([])
  }

  const searchAffectUser = async (term: string) => {
    setAffectSearch(term)
    if (term.length < 2) { setAffectResults([]); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) setAffectResults(result.data)
    } catch (err) { console.error('Erreur recherche affectation', err) }
  }

  const confirmAffect = async (u: User) => {
    if (!affectItemId || !u.Mail) return
    setAffectLoading(true)
    try {
      await DCPO_LISTE_ANORMALIEService.update(String(affectItemId), {
        personneAffecter: {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(u.Mail),
        } as never,
      })
      closeAffectModal()
      await fetchItems()
    } catch (err) { console.error('Erreur affectation', err) }
    finally { setAffectLoading(false) }
  }

  const closeAffectModal = () => {
    setAffectItemId(null)
    setAffectSearch('')
    setAffectResults([])
  }

  const fetchItems = async () => {
    setLoading(true)
    try {
      const result = await DCPO_LISTE_ANORMALIEService.getAll()
      if (result.data) setItems(result.data)
    } catch (err) { console.error('Erreur chargement anomalies', err) }
    finally { setLoading(false) }
  }

  useEffect(() => {
    fetchItems()
    const loadLists = async () => {
      try {
        const [agencesRes, reseauxRes] = await Promise.all([
          DCPO_LISTE_AGENCESService.getAll(),
          DCPO_LISTE_RESEAUXService.getAll(),
        ])
        if (agencesRes.data) setAgences(agencesRes.data)
        if (reseauxRes.data) setReseaux(reseauxRes.data)
      } catch (err) { console.error('Erreur chargement agences/reseaux', err) }
    }
    loadLists()
  }, [])

  const handleChange = (field: string, value: string | number) => {
    setForm(prev => ({ ...prev, [field]: value }))
  }

  const resetForm = () => {
    setForm(EMPTY_FORM)
    setAuteurSearch(''); setAuteurEmail('')
    setAffecteFormSearch(''); setAffecteFormEmail('')
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setSubmitting(true)
    try {
      const payload: Record<string, unknown> = {}
      if (form.field_0) payload.field_0 = form.field_0 + 'T00:00:00Z'
      if (form.field_4) payload.field_4 = form.field_4
      if (form.field_5) payload.field_5 = form.field_5
      if (form.field_6) payload.field_6 = form.field_6
      if (form.field_7) payload.field_7 = form.field_7
      if (form.field_8) payload.field_8 = form.field_8
      if (form.field_9) payload.field_9 = form.field_9 + 'T00:00:00Z'
      if (form.field_10) payload.field_10 = form.field_10

      if (auteurEmail) {
        payload['auteur_anormalie'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(auteurEmail),
        }
      }
      if (userEmail) {
        payload['declarant_anormalie'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(userEmail),
        }
      }
      if (affecteFormEmail) {
        payload['personneAffecter'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(affecteFormEmail),
        }
      }

      const result = await DCPO_LISTE_ANORMALIEService.create(payload as Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>)
      if (!result.success) { console.error('Erreur SharePoint:', result.error); return }

      resetForm()
      setShowForm(false)
      await fetchItems()
    } catch (err) { console.error('Erreur creation:', err) }
    finally { setSubmitting(false) }
  }

  return (
    <>
      <div className="content-header">
        <h2>Liste des anomalies</h2>
        <button className="btn-add" onClick={() => setShowForm(!showForm)}>
          {showForm ? 'Annuler' : '+ Nouvelle anomalie'}
        </button>
      </div>

      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{items.length}</span>
          <span className="stat-label">Total</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{items.filter(i => i.field_10 === 'Ouvert').length}</span>
          <span className="stat-label">Ouvert</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{items.filter(i => i.field_10 === 'En cours').length}</span>
          <span className="stat-label">En cours</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{items.filter(i => i.field_10 === 'Resolu').length}</span>
          <span className="stat-label">Resolu</span>
        </div>
        <div className="stat-card clos">
          <span className="stat-value">{items.filter(i => i.field_10 === 'Clos').length}</span>
          <span className="stat-label">Clos</span>
        </div>
        <div className="stat-card montant">
          <span className="stat-value">{items.reduce((s, i) => s + (i.field_8 ?? 0), 0).toLocaleString()}</span>
          <span className="stat-label">Montant total</span>
        </div>
      </div>

      {showForm && (
        <form className="create-form" onSubmit={handleSubmit}>
          <h2>Creer une anomalie</h2>
          <div className="form-grid">
            <div className="form-field">
              <label>Declarant</label>
              <input type="text" readOnly value={userName ?? ''} />
              <span className="selected-email">{userEmail}</span>
            </div>

            <div className="form-field">
              <label>Auteur</label>
              <div className="autocomplete-wrapper">
                <input type="text" value={auteurSearch} placeholder="Rechercher un auteur..."
                  onChange={e => searchAuteur(e.target.value)}
                  onBlur={() => setTimeout(() => setShowAuteurDropdown(false), 200)} />
                {auteurEmail && <span className="selected-email">{auteurEmail}</span>}
                {showAuteurDropdown && auteurResults.length > 0 && (
                  <ul className="autocomplete-dropdown">
                    {auteurResults.map(u => (
                      <li key={u.Id} onClick={() => selectAuteur(u)}>
                        <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </div>

            <div className="form-field">
              <label>Personne affectee</label>
              <div className="autocomplete-wrapper">
                <input type="text" value={affecteFormSearch} placeholder="Rechercher une personne..."
                  onChange={e => searchAffecteForm(e.target.value)}
                  onBlur={() => setTimeout(() => setShowAffecteFormDropdown(false), 200)} />
                {affecteFormEmail && <span className="selected-email">{affecteFormEmail}</span>}
                {showAffecteFormDropdown && affecteFormResults.length > 0 && (
                  <ul className="autocomplete-dropdown">
                    {affecteFormResults.map(u => (
                      <li key={u.Id} onClick={() => selectAffecteForm(u)}>
                        <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </div>

            {Object.keys(FIELD_LABELS).map(field => (
              <div className="form-field" key={field}>
                <label>{FIELD_LABELS[field]}</label>
                {field === 'field_8' ? (
                  <input type="number" value={form[field as keyof FormState] as number}
                    onChange={e => handleChange(field, Number(e.target.value))} />
                ) : field === 'field_10' ? (
                  <select value={form[field as keyof FormState] as string}
                    onChange={e => handleChange(field, e.target.value)}>
                    <option value="">-- Choisir --</option>
                    <option value="Ouvert">Ouvert</option>
                    <option value="En cours">En cours</option>
                    <option value="Resolu">Resolu</option>
                    <option value="Clos">Clos</option>
                  </select>
                ) : field === 'field_5' ? (
                  <select value={form[field as keyof FormState] as string}
                    onChange={e => handleChange(field, e.target.value)}>
                    <option value="">-- Choisir --</option>
                    <option value="Operationnel">Operationnel</option>
                    <option value="Fraude">Fraude</option>
                    <option value="Commercial">Commercial</option>
                  </select>
                ) : field === 'field_6' ? (
                  <select value={form[field as keyof FormState] as string}
                    onChange={e => {
                      const agenceId = e.target.value
                      handleChange('field_6', agenceId)
                      const agence = agences.find(a => String(a.ID) === agenceId)
                      handleChange('field_7', agence?.field_1 ? String(agence.field_1) : '')
                    }}>
                    <option value="">-- Choisir une agence --</option>
                    {agences.map(a => (
                      <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
                    ))}
                  </select>
                ) : field === 'field_7' ? (
                  <input type="text" readOnly
                    value={reseaux.find(r => String(r.ID) === (form[field as keyof FormState] as string))?.field_1 ?? ''}
                    placeholder="Selectionner une agence" />
                ) : field === 'field_4' ? (
                  <textarea value={form[field as keyof FormState] as string}
                    onChange={e => handleChange(field, e.target.value)} rows={3} />
                ) : DATE_FIELDS.includes(field) ? (
                  <input type="date" value={form[field as keyof FormState] as string}
                    max={field === 'field_0' ? new Date().toISOString().split('T')[0] : undefined}
                    onChange={e => handleChange(field, e.target.value)} />
                ) : (
                  <input type="text" value={form[field as keyof FormState] as string}
                    onChange={e => handleChange(field, e.target.value)} />
                )}
              </div>
            ))}
          </div>
          <button className="btn-submit" type="submit" disabled={submitting}>
            {submitting ? 'Enregistrement...' : 'Enregistrer'}
          </button>
        </form>
      )}

      {loading ? (
        <p className="loading-text">Chargement des anomalies...</p>
      ) : items.length === 0 ? (
        <p className="loading-text">Aucune anomalie trouvee.</p>
      ) : (
        <div className="table-wrapper">
          <table className="data-table">
            <thead>
              <tr>
                <th>Declarant</th>
                <th>Auteur</th>
                <th>Personne affectee</th>
                {Object.values(FIELD_LABELS).map(label => (
                  <th key={label}>{label}</th>
                ))}
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {items.map(item => (
                <tr key={item.ID}>
                  <td>{item.declarant_anormalie?.DisplayName ?? '-'}</td>
                  <td>{item.auteur_anormalie?.DisplayName ?? '-'}</td>
                  <td>{item.personneAffecter?.DisplayName ?? '-'}</td>
                  <td>{item.field_0 ? new Date(item.field_0).toLocaleDateString() : '-'}</td>
                  <td>{item.field_4 ? stripHtml(item.field_4) : '-'}</td>
                  <td>{item.field_5 ?? '-'}</td>
                  <td>{agences.find(a => String(a.ID) === item.field_6)?.Title ?? item.field_6 ?? '-'}</td>
                  <td>{reseaux.find(r => String(r.ID) === item.field_7)?.field_1 ?? item.field_7 ?? '-'}</td>
                  <td>{item.field_8 ?? '-'}</td>
                  <td>{item.field_9 ? new Date(item.field_9).toLocaleDateString() : '-'}</td>
                  <td>{item.field_10 ?? '-'}</td>
                  <td>
                    <div className="actions-col">
                      <button className="btn-detail" onClick={() => setDetailItem(item)}>
                        Detail
                      </button>
                      <button className="btn-affect" onClick={() => setAffectItemId(item.ID ?? null)}>
                        Affecter
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {detailItem && (
        <div className="modal-overlay" onClick={() => setDetailItem(null)}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 560 }}>
            <div className="modal-header">
              <h2>Detail de l'anomalie</h2>
              <button className="modal-close" onClick={() => setDetailItem(null)}>&times;</button>
            </div>
            <div className="modal-body">
              <dl className="detail-grid">
                <dt>Declarant</dt><dd>{detailItem.declarant_anormalie?.DisplayName ?? '-'}</dd>
                <dt>Auteur</dt><dd>{detailItem.auteur_anormalie?.DisplayName ?? '-'}</dd>
                <dt>Personne affectee</dt><dd>{detailItem.personneAffecter?.DisplayName ?? '-'}</dd>
                <dt>Date</dt><dd>{detailItem.field_0 ? new Date(detailItem.field_0).toLocaleDateString() : '-'}</dd>
                <dt>Cause</dt><dd>{detailItem.field_4 ? stripHtml(detailItem.field_4) : '-'}</dd>
                <dt>Classification</dt><dd>{detailItem.field_5 ?? '-'}</dd>
                <dt>Agence</dt><dd>{agences.find(a => String(a.ID) === detailItem.field_6)?.Title ?? detailItem.field_6 ?? '-'}</dd>
                <dt>Reseau</dt><dd>{reseaux.find(r => String(r.ID) === detailItem.field_7)?.field_1 ?? detailItem.field_7 ?? '-'}</dd>
                <dt>Montant</dt><dd>{detailItem.field_8?.toLocaleString() ?? '-'}</dd>
                <dt>Date regularisation</dt><dd>{detailItem.field_9 ? new Date(detailItem.field_9).toLocaleDateString() : '-'}</dd>
                <dt>Statut</dt><dd>{detailItem.field_10 ?? '-'}</dd>
              </dl>
            </div>
          </div>
        </div>
      )}

      {affectItemId !== null && (
        <div className="modal-overlay" onClick={closeAffectModal}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Affecter l'anomalie</h2>
              <button className="modal-close" onClick={closeAffectModal}>&times;</button>
            </div>
            <div className="modal-body">
              <input type="text" value={affectSearch} placeholder="Rechercher un utilisateur..."
                onChange={e => searchAffectUser(e.target.value)} autoFocus />
              {affectResults.length > 0 && (
                <ul className="modal-results">
                  {affectResults.map(u => (
                    <li key={u.Id} onClick={() => confirmAffect(u)}>
                      <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                    </li>
                  ))}
                </ul>
              )}
              {affectLoading && <p className="loading-text">Affectation en cours...</p>}
            </div>
          </div>
        </div>
      )}
    </>
  )
}
