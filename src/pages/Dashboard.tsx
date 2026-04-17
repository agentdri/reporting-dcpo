import { useState, useEffect } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { DCPO_LISTE_ANORMALIERead, DCPO_LISTE_ANORMALIEWrite } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import type { User } from '../generated/models/Office365UsersModel'
import './Dashboard.css'

interface DashboardProps {
  userName?: string
  userRole?: string
  userEmail?: string
}

type Tab = 'anomalie' | 'reporting'

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

export default function Dashboard({ userName, userRole, userEmail }: DashboardProps) {
  const [activeTab, setActiveTab] = useState<Tab>('anomalie')
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [loading, setLoading] = useState(true)
  const [showForm, setShowForm] = useState(false)
  const [form, setForm] = useState(EMPTY_FORM)
  const [submitting, setSubmitting] = useState(false)

  // Context menu
  const [menuOpenId, setMenuOpenId] = useState<number | null>(null)

  // Agences & Reseaux
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])

  // Auteur search
  const [auteurSearch, setAuteurSearch] = useState('')
  const [auteurEmail, setAuteurEmail] = useState('')
  const [auteurResults, setAuteurResults] = useState<User[]>([])
  const [showAuteurDropdown, setShowAuteurDropdown] = useState(false)

  // Personne affectee search
  const [affecteSearch, setAffecteSearch] = useState('')
  const [affecteEmail, setAffecteEmail] = useState('')
  const [affecteResults, setAffecteResults] = useState<User[]>([])
  const [showAffecteDropdown, setShowAffecteDropdown] = useState(false)

  // Declarant = utilisateur connecte

  const searchAuteur = async (term: string) => {
    setAuteurSearch(term)
    if (term.length < 2) {
      setAuteurResults([])
      setShowAuteurDropdown(false)
      return
    }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) {
        setAuteurResults(result.data)
        setShowAuteurDropdown(true)
      }
    } catch (err) {
      console.error('Erreur recherche auteur', err)
    }
  }

  const selectAuteur = (u: User) => {
    setAuteurEmail(u.Mail ?? '')
    setAuteurSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAuteurDropdown(false)
    setAuteurResults([])
  }

  const searchAffecte = async (term: string) => {
    setAffecteSearch(term)
    if (term.length < 2) {
      setAffecteResults([])
      setShowAffecteDropdown(false)
      return
    }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) {
        setAffecteResults(result.data)
        setShowAffecteDropdown(true)
      }
    } catch (err) {
      console.error('Erreur recherche personne affectee', err)
    }
  }

  const selectAffecte = (u: User) => {
    setAffecteEmail(u.Mail ?? '')
    setAffecteSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAffecteDropdown(false)
    setAffecteResults([])
  }



  const fetchItems = async () => {
    setLoading(true)
    try {
      const result = await DCPO_LISTE_ANORMALIEService.getAll()
      if (result.data) {
        setItems(result.data)
      }
    } catch (err) {
      console.error('Erreur lors du chargement des anomalies', err)
    } finally {
      setLoading(false)
    }
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
      } catch (err) {
        console.error('Erreur chargement agences/reseaux', err)
      }
    }
    loadLists()
  }, [])

  const handleChange = (field: string, value: string | number) => {
    setForm(prev => ({ ...prev, [field]: value }))
  }

  const resetForm = () => {
    setForm(EMPTY_FORM)
    setAuteurSearch('')
    setAuteurEmail('')
    setAffecteSearch('')
    setAffecteEmail('')
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

      if (affecteEmail) {
        payload['personneAffecter'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(affecteEmail),
        }
      }

      console.log('Payload envoye:', JSON.stringify(payload))
      const result = await DCPO_LISTE_ANORMALIEService.create(payload as Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>)
      console.log('Resultat creation:', result)

      if (!result.success) {
        console.error('Erreur SharePoint:', result.error)
        return
      }

      resetForm()
      setShowForm(false)
      await fetchItems()
    } catch (err) {
      console.error('Erreur lors de la creation:', err)
    } finally {
      setSubmitting(false)
    }
  }

  return (
    <div className="dashboard">
      <header className="dashboard-topbar">
        <div className="topbar-left">
          <h1 className="topbar-title">ReportingDCPO</h1>
        </div>
        <div className="topbar-user">
          <span className="topbar-user-name">{userName}</span>
          <span className="topbar-user-job">{userRole}</span>
        </div>
      </header>

      <div className="dashboard-body">
      <nav className="dashboard-nav">
        <button
          className={`nav-item ${activeTab === 'anomalie' ? 'active' : ''}`}
          onClick={() => setActiveTab('anomalie')}
        >
          Anomalies
        </button>
        <button
          className={`nav-item ${activeTab === 'reporting' ? 'active' : ''}`}
          onClick={() => setActiveTab('reporting')}
        >
          Reporting
        </button>
      </nav>

      <div className="dashboard-content">
        {activeTab === 'anomalie' && (
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
            </div>

            {showForm && (
              <form className="create-form" onSubmit={handleSubmit}>
                <h2>Creer une anomalie</h2>
                <div className="form-grid">
                  {/* Declarant */}
                  <div className="form-field">
                    <label>Declarant</label>
                    <input type="text" readOnly value={userName ?? ''} />
                    <span className="selected-email">{userEmail}</span>
                  </div>

                  {/* Auteur */}
                  <div className="form-field">
                    <label>Auteur</label>
                    <div className="autocomplete-wrapper">
                      <input
                        type="text"
                        value={auteurSearch}
                        placeholder="Rechercher un auteur..."
                        onChange={e => searchAuteur(e.target.value)}
                        onBlur={() => setTimeout(() => setShowAuteurDropdown(false), 200)}
                      />
                      {auteurEmail && (
                        <span className="selected-email">{auteurEmail}</span>
                      )}
                      {showAuteurDropdown && auteurResults.length > 0 && (
                        <ul className="autocomplete-dropdown">
                          {auteurResults.map(u => (
                            <li key={u.Id} onClick={() => selectAuteur(u)}>
                              <strong>{u.DisplayName}</strong>
                              <span>{u.Mail}</span>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
                  </div>

                  {/* Personne affectee */}
                  <div className="form-field">
                    <label>Personne affectee</label>
                    <div className="autocomplete-wrapper">
                      <input
                        type="text"
                        value={affecteSearch}
                        placeholder="Rechercher une personne..."
                        onChange={e => searchAffecte(e.target.value)}
                        onBlur={() => setTimeout(() => setShowAffecteDropdown(false), 200)}
                      />
                      {affecteEmail && (
                        <span className="selected-email">{affecteEmail}</span>
                      )}
                      {showAffecteDropdown && affecteResults.length > 0 && (
                        <ul className="autocomplete-dropdown">
                          {affecteResults.map(u => (
                            <li key={u.Id} onClick={() => selectAffecte(u)}>
                              <strong>{u.DisplayName}</strong>
                              <span>{u.Mail}</span>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
                  </div>

                  {/* Other fields */}
                  {Object.keys(FIELD_LABELS).map(field => (
                    <div className="form-field" key={field}>
                      <label>{FIELD_LABELS[field]}</label>
                      {field === 'field_8' ? (
                        <input
                          type="number"
                          value={form[field as keyof FormState] as number}
                          onChange={e => handleChange(field, Number(e.target.value))}
                        />
                      ) : field === 'field_10' ? (
                        <select
                          value={form[field as keyof FormState] as string}
                          onChange={e => handleChange(field, e.target.value)}
                        >
                          <option value="">-- Choisir --</option>
                          <option value="Ouvert">Ouvert</option>
                          <option value="En cours">En cours</option>
                          <option value="Resolu">Resolu</option>
                          <option value="Clos">Clos</option>
                        </select>
                      ) : field === 'field_5' ? (
                        <select
                          value={form[field as keyof FormState] as string}
                          onChange={e => handleChange(field, e.target.value)}
                        >
                          <option value="">-- Choisir --</option>
                          <option value="Operationnel">Operationnel</option>
                          <option value="Fraude">Fraude</option>
                          <option value="Commercial">Commercial</option>
                        </select>
                      ) : field === 'field_6' ? (
                        <select
                          value={form[field as keyof FormState] as string}
                          onChange={e => {
                            const agenceId = e.target.value
                            handleChange('field_6', agenceId)
                            const agence = agences.find(a => String(a.ID) === agenceId)
                            handleChange('field_7', agence?.field_1 ? String(agence.field_1) : '')
                          }}
                        >
                          <option value="">-- Choisir une agence --</option>
                          {agences.map(a => (
                            <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
                          ))}
                        </select>
                      ) : field === 'field_7' ? (
                        <input
                          type="text"
                          readOnly
                          value={reseaux.find(r => String(r.ID) === (form[field as keyof FormState] as string))?.field_1 ?? ''}
                          placeholder="Selectionner une agence"
                        />
                      ) : field === 'field_4' ? (
                        <textarea
                          value={form[field as keyof FormState] as string}
                          onChange={e => handleChange(field, e.target.value)}
                          rows={3}
                        />
                      ) : DATE_FIELDS.includes(field) ? (
                        <input
                          type="date"
                          value={form[field as keyof FormState] as string}
                          max={field === 'field_0' ? new Date().toISOString().split('T')[0] : undefined}
                          onChange={e => handleChange(field, e.target.value)}
                        />
                      ) : (
                        <input
                          type="text"
                          value={form[field as keyof FormState] as string}
                          onChange={e => handleChange(field, e.target.value)}
                        />
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
                        <td className="action-cell">
                          <button className="btn-dots" onClick={() => setMenuOpenId(menuOpenId === item.ID ? null : (item.ID ?? null))}>
                            &#8942;
                          </button>
                          {menuOpenId === item.ID && (
                            <div className="context-menu" onMouseLeave={() => setMenuOpenId(null)}>
                              <button onClick={() => { console.log('Voir', item.ID); setMenuOpenId(null) }}>Voir</button>
                              <button onClick={() => { console.log('Modifier', item.ID); setMenuOpenId(null) }}>Modifier</button>
                              <button className="danger" onClick={() => { console.log('Supprimer', item.ID); setMenuOpenId(null) }}>Supprimer</button>
                            </div>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}

        {activeTab === 'reporting' && (
          <div className="reporting-placeholder">
            <h2>Reporting</h2>
            <p>Section reporting en cours de construction.</p>
          </div>
        )}
      </div>
      </div>
    </div>
  )
}