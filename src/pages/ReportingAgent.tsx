import { useState, useEffect } from 'react'
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

export default function ReportingAgent() {
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])
  const [loading, setLoading] = useState(true)
  const [selectedAgent, setSelectedAgent] = useState<string | null>(null)

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

  const agentMap = new Map<string, AgentStats>()
  items.forEach(item => {
    const email = item.auteur_anormalie?.Email ?? ''
    const name = item.auteur_anormalie?.DisplayName ?? 'Inconnu'
    if (!email) return

    if (!agentMap.has(email)) {
      agentMap.set(email, {
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
    const stats = agentMap.get(email)!
    stats.total++
    stats.montantTotal += item.field_8 ?? 0
    stats.anomalies.push(item)
    if (item.field_10 === 'Ouvert') stats.ouvert++
    else if (item.field_10 === 'En cours') stats.enCours++
    else if (item.field_10 === 'Resolu') stats.resolu++
    else if (item.field_10 === 'Clos') stats.clos++
  })

  const agents = Array.from(agentMap.values()).sort((a, b) => b.total - a.total)
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
            Retour a la liste
          </button>
        )}
      </div>

      {!selectedAgent ? (
        <>
          <div className="stats-cards">
            <div className="stat-card total">
              <span className="stat-value">{agents.length}</span>
              <span className="stat-label">Agents</span>
            </div>
            <div className="stat-card ouvert">
              <span className="stat-value">{items.length}</span>
              <span className="stat-label">Total anomalies</span>
            </div>
            <div className="stat-card en-cours">
              <span className="stat-value">{items.reduce((s, i) => s + (i.field_8 ?? 0), 0).toLocaleString()}</span>
              <span className="stat-label">Montant total</span>
            </div>
          </div>

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