import { useState, useEffect } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { DCPO_LISTE_ANORMALIERead, DCPO_LISTE_ANORMALIEWrite } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { User } from '../generated/models/Office365UsersModel'
import './Dashboard.css'

interface DashboardProps {
  onBack: () => void
  userName?: string
  userJobTitle?: string
}

type Tab = 'anomalie' | 'reporting'

interface FormState {
  field_0: string
  field_3: string
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
  field_3: '',
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
  field_3: 'Unite',
  field_4: 'Cause',
  field_5: 'Classification',
  field_6: 'Agence',
  field_7: 'Reseau',
  field_8: 'Montant',
  field_9: 'Date de regularisation',
  field_10: 'Statut',
}

const DATE_FIELDS = ['field_0', 'field_9']

function toClaims(email: string) {
  return `i:0#.f|membership|${email}`
}

export default function Dashboard({ onBack, userName, userJobTitle }: DashboardProps) {
  const [activeTab, setActiveTab] = useState<Tab>('anomalie')
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [loading, setLoading] = useState(true)
  const [showForm, setShowForm] = useState(false)
  const [form, setForm] = useState(EMPTY_FORM)
  const [submitting, setSubmitting] = useState(false)

  // Auteur search
  const [auteurSearch, setAuteurSearch] = useState('')
  const [auteurEmail, setAuteurEmail] = useState('')
  const [auteurResults, setAuteurResults] = useState<User[]>([])
  const [showAuteurDropdown, setShowAuteurDropdown] = useState(false)

  // Declarant search
  const [declarantSearch, setDeclarantSearch] = useState('')
  const [declarantEmail, setDeclarantEmail] = useState('')
  const [declarantResults, setDeclarantResults] = useState<User[]>([])
  const [showDeclarantDropdown, setShowDeclarantDropdown] = useState(false)

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

  const searchDeclarant = async (term: string) => {
    setDeclarantSearch(term)
    if (term.length < 2) {
      setDeclarantResults([])
      setShowDeclarantDropdown(false)
      return
    }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) {
        setDeclarantResults(result.data)
        setShowDeclarantDropdown(true)
      }
    } catch (err) {
      console.error('Erreur recherche declarant', err)
    }
  }

  const selectDeclarant = (u: User) => {
    setDeclarantEmail(u.Mail ?? '')
    setDeclarantSearch(u.DisplayName ?? u.Mail ?? '')
    setShowDeclarantDropdown(false)
    setDeclarantResults([])
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
  }, [])

  const handleChange = (field: string, value: string | number) => {
    setForm(prev => ({ ...prev, [field]: value }))
  }

  const resetForm = () => {
    setForm(EMPTY_FORM)
    setAuteurSearch('')
    setAuteurEmail('')
    setDeclarantSearch('')
    setDeclarantEmail('')
  }

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setSubmitting(true)
    try {
      const payload: Record<string, unknown> = {}
      if (form.field_0) payload.field_0 = form.field_0 + 'T00:00:00Z'
      if (form.field_3) payload.field_3 = form.field_3
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

      if (declarantEmail) {
        payload['declarant_anormalie'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(declarantEmail),
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
          <button className="btn-back" onClick={onBack}>Retour</button>
          <h1 className="topbar-title">ReportingDCPO</h1>
        </div>
        <div className="topbar-user">
          <span className="topbar-user-name">{userName}</span>
          <span className="topbar-user-job">{userJobTitle}</span>
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

            {showForm && (
              <form className="create-form" onSubmit={handleSubmit}>
                <h2>Creer une anomalie</h2>
                <div className="form-grid">
                  {/* Declarant */}
                  <div className="form-field">
                    <label>Declarant</label>
                    <div className="autocomplete-wrapper">
                      <input
                        type="text"
                        value={declarantSearch}
                        placeholder="Rechercher un declarant..."
                        onChange={e => searchDeclarant(e.target.value)}
                        onBlur={() => setTimeout(() => setShowDeclarantDropdown(false), 200)}
                      />
                      {declarantEmail && (
                        <span className="selected-email">{declarantEmail}</span>
                      )}
                      {showDeclarantDropdown && declarantResults.length > 0 && (
                        <ul className="autocomplete-dropdown">
                          {declarantResults.map(u => (
                            <li key={u.Id} onClick={() => selectDeclarant(u)}>
                              <strong>{u.DisplayName}</strong>
                              <span>{u.Mail}</span>
                            </li>
                          ))}
                        </ul>
                      )}
                    </div>
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
                      ) : DATE_FIELDS.includes(field) ? (
                        <input
                          type="date"
                          value={form[field as keyof FormState] as string}
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
                      {Object.values(FIELD_LABELS).map(label => (
                        <th key={label}>{label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {items.map(item => (
                      <tr key={item.ID}>
                        <td>{item.declarant_anormalie?.DisplayName ?? '-'}</td>
                        <td>{item.auteur_anormalie?.DisplayName ?? '-'}</td>
                        <td>{item.field_0 ? new Date(item.field_0).toLocaleDateString() : '-'}</td>
                        <td>{item.field_3 ?? '-'}</td>
                        <td>{item.field_4 ?? '-'}</td>
                        <td>{item.field_5 ?? '-'}</td>
                        <td>{item.field_6 ?? '-'}</td>
                        <td>{item.field_7 ?? '-'}</td>
                        <td>{item.field_8 ?? '-'}</td>
                        <td>{item.field_9 ? new Date(item.field_9).toLocaleDateString() : '-'}</td>
                        <td>{item.field_10 ?? '-'}</td>
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