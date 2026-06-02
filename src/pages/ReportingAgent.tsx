/**
 * ============================================================================
 * REPORTING PAR AGENT — STATISTIQUES PAR AUTEUR D'ANOMALIE
 * ============================================================================
 *
 * Vue agrégée des anomalies regroupées par AGENT (= auteur_anormalie).
 *
 * Use case : un manager veut voir combien d'anomalies chaque agent
 * (commercial, opérationnel) a déclarées, leur répartition par statut,
 * leur montant cumulé.
 *
 * Workflow utilisateur :
 *   1. Vue principale : tableau des agents avec leurs stats
 *      (Total, Ouvert, En cours, Resolu, Clos, Montant)
 *   2. Clic "Detail" → bascule sur la vue détail de l'agent sélectionné
 *      avec la liste de TOUTES ses anomalies
 *   3. Bouton "Retour à la liste" pour revenir à la vue principale
 *
 * Filtres (avec bouton Rechercher pour application différée) :
 *   - Période (date)
 *   - Agence, Réseau (selects)
 *   - Classification, Criticité (selects)
 *   - Statut (select)
 *   - Agent, Personne affectée (texte)
 *
 * Filtrage 100 % côté client (les anomalies sont chargées une fois au montage).
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { formatMontantCompact } from '../lib/formatters'

/** Nettoie un texte HTML pour ne garder que le contenu textuel (DOMParser). */
function stripHtml(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  return doc.body.textContent?.trim() ?? ''
}

/**
 * Statistiques agrégées d'un agent.
 *
 * Identification :
 *   - email : clé unique du regroupement (case-insensitive en pratique)
 *   - displayName : pour l'affichage UX
 *
 * Compteurs :
 *   - total : toutes anomalies confondues
 *   - ouvert / enCours / resolu / clos : par statut
 *   - montantTotal : somme des field_8
 *
 * Détail :
 *   - anomalies : la liste brute, utilisée dans la vue détail (drill-down)
 */
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

/** Critères de filtrage du module — tous appliqués côté client. */
interface FilterState {
  dateFrom: string
  dateTo: string
  agence: string
  reseau: string
  classification: string
  criticite: string
  statut: string
  agent: string    // texte : nom OU email
  affecte: string  // texte : nom OU email
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

/**
 * Filtres de la VUE DÉTAIL (anomalies d'un agent en particulier).
 *
 * Identiques aux filtres principaux mais SANS le champ "agent" (on est déjà
 * sur un agent donné) et AVEC une recherche libre sur cause/déclarant.
 * Appliqués côté client sur selectedStats.anomalies.
 */
interface DetailFilterState {
  dateFrom: string
  dateTo: string
  agence: string
  reseau: string
  classification: string
  criticite: string
  statut: string
  affecte: string
  search: string
}

const EMPTY_DETAIL_FILTERS: DetailFilterState = {
  dateFrom: '',
  dateTo: '',
  agence: '',
  reseau: '',
  classification: '',
  criticite: '',
  statut: '',
  affecte: '',
  search: '',
}

/** Listes fermées pour les selects de filtres. */
const CRITICITE_OPTIONS = ['Faible', 'Moyenne', 'Haute', 'Critique']
const STATUT_OPTIONS = ['Ouvert', 'En cours', 'Resolu', 'Clos']

export default function ReportingAgent() {
  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS
   * ────────────────────────────────────────────────────────────────────── */

  /** Liste brute des anomalies (chargées une fois au montage). */
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])
  const [loading, setLoading] = useState(true)
  /**
   * Agent sélectionné pour la vue détail.
   *   - null  : vue principale (tableau de tous les agents)
   *   - email : vue détail de cet agent (toutes ses anomalies)
   */
  const [selectedAgent, setSelectedAgent] = useState<string | null>(null)
  /** Filtres en cours de saisie (binding inputs). */
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** Filtres effectivement appliqués (snapshot au clic Rechercher). */
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** Filtres de la vue détail (saisie + appliqués), propres à un agent. */
  const [detailFilters, setDetailFilters] = useState<DetailFilterState>(EMPTY_DETAIL_FILTERS)
  const [appliedDetailFilters, setAppliedDetailFilters] = useState<DetailFilterState>(EMPTY_DETAIL_FILTERS)

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

  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS DES FILTRES
   * ────────────────────────────────────────────────────────────────────── */

  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  /** Snapshot saisie → applied (déclenche le recalcul de filteredItems). */
  const applyFilters = () => setAppliedFilters(filters)
  /** Reset complet (vide saisie + applied). */
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS)
    setAppliedFilters(EMPTY_FILTERS)
  }

  /* ─── Handlers des filtres de la VUE DÉTAIL ─────────────────────────── */
  const updateDetailFilter = <K extends keyof DetailFilterState>(key: K, value: DetailFilterState[K]) => {
    setDetailFilters(prev => ({ ...prev, [key]: value }))
  }
  const applyDetailFilters = () => setAppliedDetailFilters(detailFilters)
  const resetDetailFilters = () => {
    setDetailFilters(EMPTY_DETAIL_FILTERS)
    setAppliedDetailFilters(EMPTY_DETAIL_FILTERS)
  }

  /**
   * Ouvre la vue détail d'un agent en repartant de filtres détail vierges
   * (sinon les filtres d'un agent précédemment consulté persisteraient).
   */
  const openAgentDetail = (email: string) => {
    setDetailFilters(EMPTY_DETAIL_FILTERS)
    setAppliedDetailFilters(EMPTY_DETAIL_FILTERS)
    setSelectedAgent(email)
  }


  /* ──────────────────────────────────────────────────────────────────────
   * CALCULS DÉRIVÉS — PIPELINE items → filteredItems → agentMap → agents
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Étape 1 — applique tous les filtres aux anomalies brutes.
   *
   * Optimisations :
   *   - Pré-calcul des termes lowercase / dates avant le filter()
   *   - Early return false dès qu'un critère ne matche pas
   *   - Ordre : filtres simples d'abord, recherches texte en dernier
   */
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

  /**
   * Étape 2 — agrégation des items filtrés par AGENT (auteur_anormalie).
   *
   * Algorithme :
   *   - Map<email, AgentStats> alimentée en un seul passage
   *   - Skip les items sans email auteur (orphelins)
   *   - Pour chaque item : créer le bucket si absent, accumuler les compteurs
   *
   * Map (vs objet) : permet une recherche O(1) par email + itération facile
   * via .values() pour le .sort() final.
   */
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

  /**
   * Étape 3 — tableau des agents trié par volume décroissant.
   * Le top contributeur en termes d'anomalies déclarées apparaît en premier.
   */
  const agents = useMemo(
    () => Array.from(agentMap.values()).sort((a, b) => b.total - a.total),
    [agentMap],
  )
  /** Stats de l'agent sélectionné (vue détail), null en vue principale. */
  const selectedStats = selectedAgent ? agentMap.get(selectedAgent) : null

  /* ──────────────────────────────────────────────────────────────────────
   * PAGINATION — 2 paginators distincts
   *   1. Pour le tableau "Liste des agents" (vue principale)
   *   2. Pour le tableau "Détail des anomalies de l'agent sélectionné"
   * Les deux ont leurs propres resetKey indépendants.
   * ────────────────────────────────────────────────────────────────────── */
  const agentsPagination = usePagination({
    total: agents.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  const pagedAgents = useMemo(
    () => agents.slice(agentsPagination.start, agentsPagination.end),
    [agents, agentsPagination.start, agentsPagination.end],
  )

  /**
   * Anomalies de l'agent sélectionné APRÈS application des filtres détail.
   * Part de selectedStats.anomalies (déjà filtré par les filtres principaux)
   * et applique en plus les critères de la vue détail.
   */
  const detailFilteredAnomalies = useMemo(() => {
    if (!selectedStats) return []
    const f = appliedDetailFilters
    const fromDate = f.dateFrom ? new Date(`${f.dateFrom}T00:00:00`) : undefined
    const toDate = f.dateTo ? new Date(`${f.dateTo}T23:59:59`) : undefined
    const affecteTerm = f.affecte.trim().toLowerCase()
    const searchTerm = f.search.trim().toLowerCase()

    return selectedStats.anomalies.filter(it => {
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
      if (affecteTerm) {
        const haystack = `${it.personneAffecter?.DisplayName ?? ''} ${it.personneAffecter?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(affecteTerm)) return false
      }
      if (searchTerm) {
        const haystack = `${it.declarant_anormalie?.DisplayName ?? ''} ${it.field_4 ? stripHtml(it.field_4) : ''}`.toLowerCase()
        if (!haystack.includes(searchTerm)) return false
      }
      return true
    })
  }, [selectedStats, appliedDetailFilters])

  /**
   * Stats recalculées sur la liste filtrée détail (pour que cartes + tableau
   * + bulletin imprimable restent cohérents avec les filtres détail).
   */
  const detailStats = useMemo(() => {
    const a = detailFilteredAnomalies
    return {
      total: a.length,
      ouvert: a.filter(i => i.field_10 === 'Ouvert').length,
      enCours: a.filter(i => i.field_10 === 'En cours').length,
      resolu: a.filter(i => i.field_10 === 'Resolu').length,
      clos: a.filter(i => i.field_10 === 'Clos').length,
      montantTotal: a.reduce((s, i) => s + (i.field_8 ?? 0), 0),
    }
  }, [detailFilteredAnomalies])

  const detailPagination = usePagination({
    total: detailFilteredAnomalies.length,
    // Reset quand on change d'agent OU que les filtres (principaux/détail) changent
    resetKey: `${selectedAgent ?? ''}|${JSON.stringify(appliedFilters)}|${JSON.stringify(appliedDetailFilters)}`,
  })
  const pagedDetailAnomalies = useMemo(
    () => detailFilteredAnomalies.slice(detailPagination.start, detailPagination.end),
    [detailFilteredAnomalies, detailPagination.start, detailPagination.end],
  )

  // Affichage simple pendant le chargement initial
  if (loading) {
    return <p className="loading-text">Chargement du reporting...</p>
  }

  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   *
   * Layout conditionnel :
   *   - selectedAgent === null → Vue PRINCIPALE (filtres + tableau agents)
   *   - selectedAgent !== null → Vue DÉTAIL (header agent + ses anomalies)
   * ════════════════════════════════════════════════════════════════════════ */
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
              {/* Format compact (millions au-delà d'1 M) pour éviter le débordement. */}
              <span className="stat-value">{formatMontantCompact(filteredItems.reduce((s, i) => s + (i.field_8 ?? 0), 0))}</span>
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
                  {pagedAgents.map(agent => (
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
                        <button className="btn-affect" onClick={() => openAgentDetail(agent.email)}>
                          Detail
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <Pagination
                state={agentsPagination}
                total={agents.length}
                itemLabel="agents"
              />
            </div>
          )}
        </>
      ) : selectedStats && (
        <>
          <div className="agent-detail-header">
            <div>
              <h3>{selectedStats.displayName}</h3>
              <span className="agent-detail-email">{selectedStats.email}</span>
            </div>
            {/* Bouton d'impression : génère le bulletin complet de l'agent
                (toutes ses anomalies). Le rendu imprimé s'appuie sur la section
                .agent-print-only + les règles @media print (cf. Dashboard.css),
                sur le même principe que l'impression des bulletins d'anomalies. */}
            <button
              type="button"
              className="btn-affect no-print"
              onClick={() => window.print()}
              title="Imprimer le bulletin complet de l'agent"
            >
              🖨️ Imprimer le bulletin
            </button>
          </div>

          {/* ─── Barre de filtres propre à l'agent (no-print) ─────────── */}
          <div className="filters-bar no-print">
            <div className="filter-field">
              <label>Date du</label>
              <input type="date" value={detailFilters.dateFrom} onChange={e => updateDetailFilter('dateFrom', e.target.value)} />
            </div>
            <div className="filter-field">
              <label>Date au</label>
              <input type="date" value={detailFilters.dateTo} onChange={e => updateDetailFilter('dateTo', e.target.value)} />
            </div>
            <div className="filter-field">
              <label>Agence</label>
              <select value={detailFilters.agence} onChange={e => updateDetailFilter('agence', e.target.value)}>
                <option value="">Toutes</option>
                {agences.map(a => <option key={a.ID} value={String(a.ID)}>{a.Title}</option>)}
              </select>
            </div>
            <div className="filter-field">
              <label>Réseau</label>
              <select value={detailFilters.reseau} onChange={e => updateDetailFilter('reseau', e.target.value)}>
                <option value="">Tous</option>
                {reseaux.map(r => <option key={r.ID} value={String(r.ID)}>{r.field_1 ?? r.Title}</option>)}
              </select>
            </div>
            <div className="filter-field">
              <label>Classification</label>
              <select value={detailFilters.classification} onChange={e => updateDetailFilter('classification', e.target.value)}>
                <option value="">Toutes</option>
                <option value="Operationnel">Opérationnel</option>
                <option value="Fraude">Fraude</option>
                <option value="Commercial">Commercial</option>
              </select>
            </div>
            <div className="filter-field">
              <label>Criticité</label>
              <select value={detailFilters.criticite} onChange={e => updateDetailFilter('criticite', e.target.value)}>
                <option value="">Toutes</option>
                {CRITICITE_OPTIONS.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div className="filter-field">
              <label>Statut</label>
              <select value={detailFilters.statut} onChange={e => updateDetailFilter('statut', e.target.value)}>
                <option value="">Tous</option>
                {STATUT_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
            <div className="filter-field">
              <label>Personne affectée</label>
              <input
                type="text"
                placeholder="Nom ou email..."
                value={detailFilters.affecte}
                onChange={e => updateDetailFilter('affecte', e.target.value)}
              />
            </div>
            <div className="filter-field" style={{ flex: 1, minWidth: 180 }}>
              <label>Recherche (cause / déclarant)</label>
              <input
                type="text"
                placeholder="Mot-clé..."
                value={detailFilters.search}
                onChange={e => updateDetailFilter('search', e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') applyDetailFilters() }}
              />
            </div>
            <div className="filter-actions">
              <button type="button" className="btn-search-filters" onClick={applyDetailFilters}>
                Rechercher
              </button>
              <button type="button" className="btn-reset-filters" onClick={resetDetailFilters}>
                Réinitialiser
              </button>
            </div>
          </div>

          <div className="stats-cards">
            <div className="stat-card total">
              <span className="stat-value">{detailStats.total}</span>
              <span className="stat-label">Total</span>
            </div>
            <div className="stat-card ouvert">
              <span className="stat-value">{detailStats.ouvert}</span>
              <span className="stat-label">Ouvert</span>
            </div>
            <div className="stat-card en-cours">
              <span className="stat-value">{detailStats.enCours}</span>
              <span className="stat-label">En cours</span>
            </div>
            <div className="stat-card resolu">
              <span className="stat-value">{detailStats.resolu}</span>
              <span className="stat-label">Resolu</span>
            </div>
            <div className="stat-card clos">
              <span className="stat-value">{detailStats.clos}</span>
              <span className="stat-label">Clos</span>
            </div>
          </div>

          {detailFilteredAnomalies.length === 0 ? (
            <p className="loading-text">Aucune anomalie ne correspond aux critères.</p>
          ) : (
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
                {pagedDetailAnomalies.map(item => (
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
            <Pagination
              state={detailPagination}
              total={detailFilteredAnomalies.length}
              itemLabel="anomalies"
            />
          </div>
          )}

          {/* ─── BULLETIN IMPRIMABLE ──────────────────────────────────────
              Masqué à l'écran (.agent-print-only → display:none), affiché
              uniquement à l'impression via @media print. Contient : identité
              agent + synthèse + anomalies (NON paginées, contrairement au
              tableau écran). Reflète les filtres détail appliqués → sans
              filtre = bulletin complet de l'agent. */}
          <div className="agent-print-only">
            <div className="agent-print-header">
              <h1>Bulletin d'anomalies</h1>
              <p>
                Agent : <strong>{selectedStats.displayName}</strong>
                {selectedStats.email ? ` (${selectedStats.email})` : ''}
              </p>
              <p>Édité le {new Date().toLocaleDateString('fr-FR')}</p>
            </div>

            <h2 className="agent-print-section-title">Synthèse</h2>
            <table className="agent-print-table agent-print-summary">
              <tbody>
                <tr><th>Total anomalies</th><td>{detailStats.total}</td></tr>
                <tr><th>Ouvert</th><td>{detailStats.ouvert}</td></tr>
                <tr><th>En cours</th><td>{detailStats.enCours}</td></tr>
                <tr><th>Résolu</th><td>{detailStats.resolu}</td></tr>
                <tr><th>Clos</th><td>{detailStats.clos}</td></tr>
                <tr><th>Montant total</th><td>{detailStats.montantTotal.toLocaleString('fr-FR')}</td></tr>
              </tbody>
            </table>

            <h2 className="agent-print-section-title">
              Détail des anomalies ({detailStats.total})
            </h2>
            <table className="agent-print-table">
              <thead>
                <tr>
                  <th>Date</th>
                  <th>Déclarant</th>
                  <th>Cause</th>
                  <th>Classification</th>
                  <th>Agence</th>
                  <th>Réseau</th>
                  <th>Montant</th>
                  <th>Date régul.</th>
                  <th>Statut</th>
                  <th>Personne affectée</th>
                </tr>
              </thead>
              <tbody>
                {detailFilteredAnomalies.map(item => (
                  <tr key={item.ID}>
                    <td>{item.field_0 ? new Date(item.field_0).toLocaleDateString('fr-FR') : '-'}</td>
                    <td>{item.declarant_anormalie?.DisplayName ?? '-'}</td>
                    <td>{item.field_4 ? stripHtml(item.field_4) : '-'}</td>
                    <td>{item.field_5 ?? '-'}</td>
                    <td>{agences.find(a => String(a.ID) === item.field_6)?.Title ?? item.field_6 ?? '-'}</td>
                    <td>{reseaux.find(r => String(r.ID) === item.field_7)?.field_1 ?? item.field_7 ?? '-'}</td>
                    <td>{item.field_8?.toLocaleString('fr-FR') ?? '-'}</td>
                    <td>{item.field_9 ? new Date(item.field_9).toLocaleDateString('fr-FR') : '-'}</td>
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
