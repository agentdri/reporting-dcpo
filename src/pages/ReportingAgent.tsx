import { useEffect, useMemo, useState } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'

function stripHtml(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  return doc.body.textContent?.trim() ?? ''
}

interface AgentStats {
  email: string
  displayName: string
  total: number
  ouvert: number
  enCours: number
  resolu: number
  clos: number
  montantTotal: number
  anomalies: DCPO_LISTE_ANORMALIERead[]
}

interface FilterState {
  dateFrom: string
  dateTo: string
  agence: string
  reseau: string
  classification: string
  criticite: string
  statut: string
  agent: string
  affecte: string
}

const EMPTY_FILTERS: FilterState = {
  dateFrom: '',
  dateTo: '',
  agence: '',
  reseau: '',
  classification: '',
  criticite: '',
  statut: '',
  agent: '',
  affecte: '',
}

const CRITICITE_OPTIONS = ['Faible', 'Moyenne', 'Haute', 'Critique']
const STATUT_OPTIONS = ['Ouvert', 'En cours', 'Resolu', 'Clos']

export default function ReportingAgent() {
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])
  const [loading, setLoading] = useState(true)
  const [selectedAgent, setSelectedAgent] = useState<string | null>(null)
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)

  useEffect(() => {
    const load = async () => {
      try {
        const [anomRes, agencesRes, reseauxRes] = await Promise.all([
          DCPO_LISTE_ANORMALIEService.getAll(),
          DCPO_LISTE_AGENCESService.getAll(),
          DCPO_LISTE_RESEAUXService.getAll(),
        ])
        if (anomRes.data) setItems(anomRes.data)
        if (agencesRes.data) setAgences(agencesRes.data)
        if (reseauxRes.data) setReseaux(reseauxRes.data)
      } catch (err) {
        console.error('Erreur chargement reporting', err)
      } finally {
        setLoading(false)
      }
    }
    load()
  }, [])

  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const applyFilters = () => setAppliedFilters(filters)
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS)
    setAppliedFilters(EMPTY_FILTERS)
  }

  const filteredItems = useMemo(() => {
    const f = appliedFilters
    const fromDate = f.dateFrom ? new Date(`${f.dateFrom}T00:00:00`) : undefined
    const toDate = f.dateTo ? new Date(`${f.dateTo}T23:59:59`) : undefined
    const agentTerm = f.agent.trim().toLowerCase()
    const affecteTerm = f.affecte.trim().toLowerCase()

    return items.filter(it => {
      if (f.agence && String(it.field_6 ?? '') !== f.agence) return false
      if (f.reseau && String(it.field_7 ?? '') !== f.reseau) return false
      if (f.classification && it.field_5 !== f.classification) return false
      if (f.criticite && it.criticiteAnomalie !== f.criticite) return false
      if (f.statut && it.field_10 !== f.statut) return false

      if (fromDate || toDate) {
        const refRaw = it.field_0 ?? it.Created
        if (!refRaw) return false
        const ref = new Date(refRaw)
        if (Number.isNaN(ref.getTime())) return false
        if (fromDate && ref < fromDate) return false
        if (toDate && ref > toDate) return false
      }

      if (agentTerm) {
        const haystack = `${it.auteur_anormalie?.DisplayName ?? ''} ${it.auteur_anormalie?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(agentTerm)) return false
      }
      if (affecteTerm) {
        const haystack = `${it.personneAffecter?.DisplayName ?? ''} ${it.personneAffecter?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(affecteTerm)) return false
      }
      return true
    })
  }, [items, appliedFilters])

  const agentMap = useMemo(() => {
    const map = new Map<string, AgentStats>()
    filteredItems.forEach(item => {
      const email = item.auteur_anormalie?.Email ?? ''
      const name = item.auteur_anormalie?.DisplayName ?? 'Inconnu'
      if (!email) return

      if (!map.has(email)) {
        map.set(email, {
          email,
          displayName: name,
          total: 0,
          ouvert: 0,
          enCours: 0,
          resolu: 0,
          clos: 0,
          montantTotal: 0,
          anomalies: [],
        })
      }
      const stats = map.get(email)!
      stats.total++
      stats.montantTotal += item.field_8 ?? 0
      stats.anomalies.push(item)
      if (item.field_10 === 'Ouvert') stats.ouvert++
      else if (item.field_10 === 'En cours') stats.enCours++
      else if (item.field_10 === 'Resolu') stats.resolu++
      else if (item.field_10 === 'Clos') stats.clos++
    })
    return map
  }, [filteredItems])

  const agents = useMemo(
    () => Array.from(agentMap.values()).sort((a, b) => b.total - a.total),
    [agentMap],
  )
  const selectedStats = selectedAgent ? agentMap.get(selectedAgent) : null

  if (loading) {
    return <p className="loading-text">Chargement du reporting...</p>
  }

  return (
    <>
      <div className="content-header">
        <h2>Reporting par Agent</h2>
        {selectedAgent && (
          <button className="btn-add" onClick={() => setSelectedAgent(null)}>
            Retour à la liste
          </button>
        )}
      </div>

      {!selectedAgent && (
        <div className="filters-bar">
          <div className="filter-field">
            <label>Date du</label>
            <input type="date" value={filters.dateFrom} onChange={e => updateFilter('dateFrom', e.target.value)} />
          </div>
          <div className="filter-field">
            <label>Date au</label>
            <input type="date" value={filters.dateTo} onChange={e => updateFilter('dateTo', e.target.value)} />
          </div>
          <div className="filter-field">
            <label>Agence</label>
            <select value={filters.agence} onChange={e => updateFilter('agence', e.target.value)}>
              <option value="">Toutes</option>
              {agences.map(a => <option key={a.ID} value={String(a.ID)}>{a.Title}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Réseau</label>
            <select value={filters.reseau} onChange={e => updateFilter('reseau', e.target.value)}>
              <option value="">Tous</option>
              {reseaux.map(r => <option key={r.ID} value={String(r.ID)}>{r.field_1 ?? r.Title}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Classification</label>
            <select value={filters.classification} onChange={e => updateFilter('classification', e.target.value)}>
              <option value="">Toutes</option>
              <option value="Operationnel">Opérationnel</option>
              <option value="Fraude">Fraude</option>
              <option value="Commercial">Commercial</option>
            </select>
          </div>
          <div className="filter-field">
            <label>Criticité</label>
            <select value={filters.criticite} onChange={e => updateFilter('criticite', e.target.value)}>
              <option value="">Toutes</option>
              {CRITICITE_OPTIONS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Statut</label>
            <select value={filters.statut} onChange={e => updateFilter('statut', e.target.value)}>
              <option value="">Tous</option>
              {STATUT_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Agent (auteur)</label>
            <input
              type="text"
              placeholder="Nom ou email..."
              value={filters.agent}
              onChange={e => updateFilter('agent', e.target.value)}
            />
          </div>
          <div className="filter-field">
            <label>Personne affectée</label>
            <input
              type="text"
              placeholder="Nom ou email..."
              value={filters.affecte}
              onChange={e => updateFilter('affecte', e.target.value)}
            />
          </div>
          <button type="button" className="btn-search-filters" onClick={applyFilters} disabled={loading}>
            Rechercher
          </button>
          <button type="button" className="btn-reset-filters" onClick={resetFilters}>Réinitialiser</button>
        </div>
      )}

      {!selectedAgent ? (
        <>
          <div className="stats-cards">
            <div className="stat-card total">
              <span className="stat-value">{agents.length}</span>
              <span className="stat-label">Agents</span>
            </div>
            <div className="stat-card ouvert">
              <span className="stat-value">{filteredItems.length}</span>
              <span className="stat-label">Total anomalies</span>
            </div>
            <div className="stat-card en-cours">
              <span className="stat-value">{filteredItems.reduce((s, i) => s + (i.field_8 ?? 0), 0).toLocaleString()}</span>
              <span className="stat-label">Montant total</span>
            </div>
          </div>

          {agents.length === 0 ? (
            <p className="loading-text">Aucun agent ne correspond aux critères.</p>
          ) : (
            <div className="table-wrapper">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>Agent</th>
                    <th>Email</th>
                    <th>Total</th>
                    <th>Ouvert</th>
                    <th>En cours</th>
                    <th>Resolu</th>
                    <th>Clos</th>
                    <th>Montant total</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {agents.map(agent => (
                    <tr key={agent.email}>
                      <td><strong>{agent.displayName}</strong></td>
                      <td>{agent.email}</td>
                      <td><strong>{agent.total}</strong></td>
                      <td>{agent.ouvert}</td>
                      <td>{agent.enCours}</td>
                      <td>{agent.resolu}</td>
                      <td>{agent.clos}</td>
                      <td>{agent.montantTotal.toLocaleString()}</td>
                      <td>
                        <button className="btn-affect" onClick={() => setSelectedAgent(agent.email)}>
                          Detail
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </>
      ) : selectedStats && (
        <>
          <div className="agent-detail-header">
            <h3>{selectedStats.displayName}</h3>
            <span className="agent-detail-email">{selectedStats.email}</span>
          </div>

          <div className="stats-cards">
            <div className="stat-card total">
              <span className="stat-value">{selectedStats.total}</span>
              <span className="stat-label">Total</span>
            </div>
            <div className="stat-card ouvert">
              <span className="stat-value">{selectedStats.ouvert}</span>
              <span className="stat-label">Ouvert</span>
            </div>
            <div className="stat-card en-cours">
              <span className="stat-value">{selectedStats.enCours}</span>
              <span className="stat-label">En cours</span>
            </div>
            <div className="stat-card resolu">
              <span className="stat-value">{selectedStats.resolu}</span>
              <span className="stat-label">Resolu</span>
            </div>
            <div className="stat-card clos">
              <span className="stat-value">{selectedStats.clos}</span>
              <span className="stat-label">Clos</span>
            </div>
          </div>

          <div className="table-wrapper">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Declarant</th>
                  <th>Cause</th>
                  <th>Classification</th>
                  <th>Agence</th>
                  <th>Reseau</th>
                  <th>Montant</th>
                  <th>Date regularisation</th>
                  <th>Statut</th>
                  <th>Personne affectee</th>
                </tr>
              </thead>
              <tbody>
                {selectedStats.anomalies.map(item => (
                  <tr key={item.ID}>
                    <td>{item.field_0 ? new Date(item.field_0).toLocaleDateString() : '-'}</td>
                    <td>{item.declarant_anormalie?.DisplayName ?? '-'}</td>
                    <td>{item.field_4 ? stripHtml(item.field_4) : '-'}</td>
                    <td>{item.field_5 ?? '-'}</td>
                    <td>{agences.find(a => String(a.ID) === item.field_6)?.Title ?? item.field_6 ?? '-'}</td>
                    <td>{reseaux.find(r => String(r.ID) === item.field_7)?.field_1 ?? item.field_7 ?? '-'}</td>
                    <td>{item.field_8?.toLocaleString() ?? '-'}</td>
                    <td>{item.field_9 ? new Date(item.field_9).toLocaleDateString() : '-'}</td>
                    <td>{item.field_10 ?? '-'}</td>
                    <td>{item.personneAffecter?.DisplayName ?? '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </>
  )
}
