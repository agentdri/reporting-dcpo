import { useState, useEffect } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import type { DCPO_LISTE_ANORMALIERead, DCPO_LISTE_ANORMALIEWrite } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import './Dashboard.css'

interface DashboardProps {
  onBack: () => void
}

type Tab = 'anomalie' | 'reporting'

const EMPTY_FORM: Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'> = {
  Title: '',
  field_0: '',
  field_2: '',
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
  Title: 'Declarant',
  field_0: 'Date',
  field_2: 'Auteur',
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

export default function Dashboard({ onBack }: DashboardProps) {
  const [activeTab, setActiveTab] = useState<Tab>('anomalie')
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [loading, setLoading] = useState(true)
  const [showForm, setShowForm] = useState(false)
  const [form, setForm] = useState(EMPTY_FORM)
  const [submitting, setSubmitting] = useState(false)

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

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setSubmitting(true)
    try {
      const payload: Record<string, unknown> = {}
      if (form.Title) payload.Title = form.Title
      if (form.field_0) payload.field_0 = form.field_0 + 'T00:00:00Z'
      if (form.field_2) payload.field_2 = form.field_2
      if (form.field_3) payload.field_3 = form.field_3
      if (form.field_4) payload.field_4 = form.field_4
      if (form.field_5) payload.field_5 = form.field_5
      if (form.field_6) payload.field_6 = form.field_6
      if (form.field_7) payload.field_7 = form.field_7
      if (form.field_8) payload.field_8 = form.field_8
      if (form.field_9) payload.field_9 = form.field_9 + 'T00:00:00Z'
      if (form.field_10) payload.field_10 = form.field_10

      console.log('Payload envoye:', JSON.stringify(payload))
      const result = await DCPO_LISTE_ANORMALIEService.create(payload as Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>)
      console.log('Resultat creation:', result)

      if (!result.success) {
        console.error('Erreur SharePoint:', result.error)
        return
      }

      setForm(EMPTY_FORM)
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
                  {Object.keys(FIELD_LABELS).map(field => (
                    <div className="form-field" key={field}>
                      <label>{FIELD_LABELS[field]}</label>
                      {field === 'field_8' ? (
                        <input
                          type="number"
                          value={form[field as keyof typeof form] as number}
                          onChange={e => handleChange(field, Number(e.target.value))}
                        />
                      ) : field === 'field_10' ? (
                        <select
                          value={form[field as keyof typeof form] as string}
                          onChange={e => handleChange(field, e.target.value)}
                        >
                          <option value="">-- Choisir --</option>
                          <option value="Regularise">Ouvert</option>
                          <option value="En cours">En cours</option>
                          <option value="Non regularise">Clos</option>
                        </select>
                      ) : field === 'field_5' ? (
                        <select
                          value={form[field as keyof typeof form] as string}
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
                          value={form[field as keyof typeof form] as string}
                          onChange={e => handleChange(field, e.target.value)}
                        />
                      ) : (
                        <input
                          type="text"
                          value={form[field as keyof typeof form] as string}
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
                      {Object.values(FIELD_LABELS).map(label => (
                        <th key={label}>{label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {items.map(item => (
                      <tr key={item.ID}>
                        <td>{item.Title ?? '-'}</td>
                        <td>{item.field_0 ? new Date(item.field_0).toLocaleDateString() : '-'}</td>
                        <td>{item.field_2 ?? '-'}</td>
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