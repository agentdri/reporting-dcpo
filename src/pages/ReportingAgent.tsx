/**
 * ============================================================================
 * REPORTING PAR AGENT — STATISTIQUES PAR AUTEUR D'ANOMALIE
 * ============================================================================
 *
 * Vue agrégée des anomalies regroupées par AGENT (= personneAffecter,
 * c'est-à-dire le contrôleur à qui l'anomalie a été assignée).
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
import {
  listControles,
  listAllEvaluations,
  getExpectedEvaluationsPerYear,
  type ControleEntry,
} from '../lib/planControleService'
import { listPACs, type Pac } from '../lib/pacService'

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


/* ──────────────────────────────────────────────────────────────────────────
 * BULLETIN — FICHE DE NOTATION DU CONTROLEUR
 *
 * Structure inspirée du format métier Excel (cf. ticket utilisateur 2026-06-05) :
 *   Plusieurs SECTIONS thématiques, chacune contenant des lignes critère
 *   avec colonnes : Prévu / Réalisé / Écart, et une ligne TAUX DE CONFORMITE
 *   en bas qui agrège les totaux.
 *
 * Sections couvertes ici :
 *   1. Surveillance des anomalies  (par classification)
 *   2. Surveillance des contrôles  (1 ligne par plan de contrôle assigné)
 *   3. Surveillance des PAC        (1 ligne par PAC assigné)
 *
 * Sémantique des colonnes :
 *   - Prévu     : objectif quantitatif (anomalies déclarées ; nb évaluations
 *                 attendues sur l'année selon la fréquence ; 1 par PAC)
 *   - Réalisé   : résultat effectif (anomalies clos+resolu ; évaluations
 *                 effectivement faites ; 1 si PAC Exécutée)
 *   - Écart     : Réalisé − Prévu (négatif = retard, 0 = à l'objectif)
 *
 * Le score % d'une section = (TOTAL Réalisé / TOTAL Prévu) × 100, capé à 100.
 * ────────────────────────────────────────────────────────────────────────── */

/** Une ligne du bulletin (critère + métriques). */
interface BulletinRow {
  critere: string
  prevu: number
  realise: number
  ecart: number
}

/** Une section complète du bulletin. */
interface BulletinSection {
  titre: string
  rows: BulletinRow[]
  totalPrevu: number
  totalRealise: number
  totalEcart: number
  /** Score % (0-100) — ratio Réalisé/Prévu, capé. 0 si rien à faire. */
  score: number
}

/** Calcule un score % à partir des totaux de section (capé à 100). */
function computeSectionScore(totalPrevu: number, totalRealise: number): number {
  if (totalPrevu <= 0) return 0
  return Math.max(0, Math.min(100, Math.round((totalRealise / totalPrevu) * 100)))
}

/** Construit une section à partir d'une liste de lignes (calcul des totaux). */
function buildSection(titre: string, rows: BulletinRow[]): BulletinSection {
  const totalPrevu = rows.reduce((s, r) => s + r.prevu, 0)
  const totalRealise = rows.reduce((s, r) => s + r.realise, 0)
  const totalEcart = totalRealise - totalPrevu
  return {
    titre,
    rows,
    totalPrevu,
    totalRealise,
    totalEcart,
    score: computeSectionScore(totalPrevu, totalRealise),
  }
}

/**
 * Construit la section "Surveillance des anomalies" pour un agent.
 *
 * Critères = classifications (Operationnel / Fraude / Commercial) + une
 * ligne "Autres / non classées" si l'agent a des anomalies sans classif.
 *
 *   Prévu   = anomalies déclarées par l'agent
 *   Réalisé = celles qui sont Resolu ou Clos
 *   Écart   = Réalisé − Prévu (négatif = anomalies en cours)
 */
function buildAnomaliesSection(anomalies: DCPO_LISTE_ANORMALIERead[]): BulletinSection {
  const classifications = ['Operationnel', 'Fraude', 'Commercial']
  const rows: BulletinRow[] = []
  for (const classif of classifications) {
    const subset = anomalies.filter(a => a.field_5 === classif)
    if (subset.length === 0) continue
    const realise = subset.filter(a => a.field_10 === 'Resolu' || a.field_10 === 'Clos').length
    rows.push({
      critere: classif,
      prevu: subset.length,
      realise,
      ecart: realise - subset.length,
    })
  }
  // Anomalies sans classification reconnue → "Autres"
  const autres = anomalies.filter(a => !classifications.includes(a.field_5 ?? ''))
  if (autres.length > 0) {
    const realise = autres.filter(a => a.field_10 === 'Resolu' || a.field_10 === 'Clos').length
    rows.push({
      critere: 'Autres / non classées',
      prevu: autres.length,
      realise,
      ecart: realise - autres.length,
    })
  }
  return buildSection('SURVEILLANCE DES ANOMALIES', rows)
}

/**
 * Construit la section "Surveillance des plans de contrôle" pour un agent.
 *
 * Critères = chaque contrôle dont l'agent est responsable.
 *   Prévu   = nb évaluations attendues sur l'année (selon la fréquence)
 *   Réalisé = nb évaluations effectivement faites (compteur préchargé)
 *   Écart   = Réalisé − Prévu
 */
function buildControlesSection(
  agentEmail: string,
  controles: ControleEntry[],
  evaluationCounts: Map<number, number>,
): BulletinSection {
  const me = agentEmail.toLowerCase()
  const mine = controles.filter(c => c.responsableEmail.toLowerCase() === me)
  const rows: BulletinRow[] = mine.map(c => {
    const prevu = getExpectedEvaluationsPerYear(c.frequence)
    const realise = evaluationCounts.get(Number(c.id)) ?? 0
    return {
      critere: c.libelle,
      prevu,
      realise,
      ecart: realise - prevu,
    }
  })
  return buildSection('SURVEILLANCE DES PLANS DE CONTRÔLE', rows)
}

/**
 * Construit la section "Surveillance des plans d'action correctif" pour un agent.
 *
 * Critères = chaque PAC dont l'agent est responsable de mise en œuvre.
 *   Prévu   = 1 (chaque PAC est une tâche à exécuter)
 *   Réalisé = 1 si statut === 'Exécutée', sinon 0
 *   Écart   = Réalisé − Prévu
 */
function buildPACsSection(agentEmail: string, pacs: Pac[]): BulletinSection {
  const me = agentEmail.toLowerCase()
  const mine = pacs.filter(p => p.responsableEmail.toLowerCase() === me)
  const rows: BulletinRow[] = mine.map(p => {
    const realise = p.statut === 'Exécutée' ? 1 : 0
    return {
      critere: p.intitule,
      prevu: 1,
      realise,
      ecart: realise - 1,
    }
  })
  return buildSection('SURVEILLANCE DES PLANS D\'ACTION CORRECTIF', rows)
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
   * Données auxiliaires utilisées pour construire le bulletin de notation
   * agent (cf. buildAnomaliesSection / buildControlesSection / buildPACsSection).
   *
   * Chargées en parallèle au montage — pas bloquantes pour l'affichage du
   * tableau principal (qui ne dépend que de `items`).
   */
  const [allControles, setAllControles] = useState<ControleEntry[]>([])
  const [allPacs, setAllPacs] = useState<Pac[]>([])
  const [evaluationCounts, setEvaluationCounts] = useState<Map<number, number>>(new Map())
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
        // Chargement parallèle : 6 sources (anomalies, agences, réseaux,
        // contrôles, PAC, évaluations). Tout pour rendre le bulletin agent
        // complet sans appels supplémentaires lors de l'impression.
        const [anomRes, agencesRes, reseauxRes, controlesRes, pacsRes, evalsRes] = await Promise.all([
          DCPO_LISTE_ANORMALIEService.getAll(),
          DCPO_LISTE_AGENCESService.getAll(),
          DCPO_LISTE_RESEAUXService.getAll(),
          listControles(),
          listPACs(),
          listAllEvaluations(),
        ])
        if (anomRes.data) setItems(anomRes.data)
        if (agencesRes.data) setAgences(agencesRes.data)
        if (reseauxRes.data) setReseaux(reseauxRes.data)
        setAllControles(controlesRes)
        setAllPacs(pacsRes)
        // Agrégation des évaluations en compteur par planControleId
        const map = new Map<number, number>()
        for (const ev of evalsRes) {
          if (!ev.planControleId) continue
          map.set(ev.planControleId, (map.get(ev.planControleId) ?? 0) + 1)
        }
        setEvaluationCounts(map)
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
   * Étape 2 — agrégation des items filtrés par AGENT.
   *
   * /!\ Sémantique métier : l'"agent" ici est le CONTRÔLEUR AFFECTÉ à
   * l'anomalie (champ `personneAffecter`), pas l'auteur qui l'a déclarée.
   * Cette vue mesure la CHARGE de chaque contrôleur — pas qui a signalé quoi.
   *
   * Algorithme :
   *   - Map<email, AgentStats> alimentée en un seul passage
   *   - Skip les items sans personne affectée (anomalies non assignées)
   *   - Pour chaque item : créer le bucket si absent, accumuler les compteurs
   *
   * Map (vs objet) : permet une recherche O(1) par email + itération facile
   * via .values() pour le .sort() final.
   */
  const agentMap = useMemo(() => {
    const map = new Map<string, AgentStats>()
    filteredItems.forEach(item => {
      const email = item.personneAffecter?.Email ?? ''
      const name = item.personneAffecter?.DisplayName ?? 'Inconnu'
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

  /**
   * Bulletin agent — 3 sections + score global.
   *
   * Construit uniquement quand un agent est sélectionné (sinon null) pour
   * éviter le calcul inutile sur la vue principale.
   *
   * Sections :
   *   1. Anomalies (par classification, basé sur la liste filtrée détail)
   *   2. Plans de contrôle (1 ligne par contrôle dont l'agent est responsable)
   *   3. PAC (1 ligne par PAC dont l'agent est responsable de mise en œuvre)
   *
   * Score global = moyenne pondérée des scores de section (poids = totalPrevu).
   */
  const bulletinSections = useMemo(() => {
    if (!selectedStats) return null
    const anomaliesSection = buildAnomaliesSection(detailFilteredAnomalies)
    const controlesSection = buildControlesSection(selectedStats.email, allControles, evaluationCounts)
    const pacsSection = buildPACsSection(selectedStats.email, allPacs)

    const totalPrevu = anomaliesSection.totalPrevu + controlesSection.totalPrevu + pacsSection.totalPrevu
    const totalRealise = anomaliesSection.totalRealise + controlesSection.totalRealise + pacsSection.totalRealise
    const scoreGlobal = computeSectionScore(totalPrevu, totalRealise)

    return { anomaliesSection, controlesSection, pacsSection, totalPrevu, totalRealise, scoreGlobal }
  }, [selectedStats, detailFilteredAnomalies, allControles, evaluationCounts, allPacs])

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
                    <th>Personne affectée</th>
                    <th>Email</th>
                    <th>Total</th>
                    <th>Ouvert</th>
                    <th>En cours</th>
                    <th>Resolu</th>
                    <th>Clos</th>
                    <th>Montant total</th>
                  </tr>
                </thead>
                <tbody>
                  {pagedAgents.map(agent => (
                    <tr
                      key={agent.email}
                      className="row-clickable"
                      onClick={() => openAgentDetail(agent.email)}
                      role="button"
                      tabIndex={0}
                      onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openAgentDetail(agent.email) } }}
                    >
                      <td><strong>{agent.displayName}</strong></td>
                      <td>{agent.email}</td>
                      <td><strong>{agent.total}</strong></td>
                      <td>{agent.ouvert}</td>
                      <td>{agent.enCours}</td>
                      <td>{agent.resolu}</td>
                      <td>{agent.clos}</td>
                      <td>{agent.montantTotal.toLocaleString()}</td>
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

          {/* ─── BULLETIN IMPRIMABLE — Fiche de notation du contrôleur ───
              Masqué à l'écran (.agent-print-only → display:none), affiché
              uniquement à l'impression via @media print.

              Format inspiré du fichier Excel métier : 3 sections de
              surveillance (Anomalies / Plans de Contrôle / PAC) avec lignes
              critère + colonnes Prévu/Réalisé/Écart + ligne "TAUX DE
              CONFORMITE" par section + score global en bas. */}
          <div className="agent-print-only">
            <div className="agent-print-header">
              <h1>FICHE DE NOTATION DU CONTROLEUR</h1>
              <p style={{ textAlign: 'center', color: '#c0392b' }}>
                Édité le {new Date().toLocaleDateString('fr-FR', { day: '2-digit', month: 'long', year: 'numeric' })}
              </p>
              <p>
                <strong>{selectedStats.displayName}</strong>
                {selectedStats.email ? ` — ${selectedStats.email}` : ''}
              </p>
            </div>

            {bulletinSections && (
              <>
                {/* Une section par domaine (anomalies, contrôles, PAC).
                    Affichée même vide pour matérialiser visuellement le
                    périmètre couvert par le bulletin. */}
                {[
                  bulletinSections.anomaliesSection,
                  bulletinSections.controlesSection,
                  bulletinSections.pacsSection,
                ].map(section => (
                  <BulletinSectionTable key={section.titre} section={section} />
                ))}

                {/* Score global = ratio Réalisé / Prévu sur l'ensemble des
                    sections (capé à 100). C'est le "TAUX DE CONFORMITE
                    GLOBAL" du contrôleur. */}
                <h2 className="agent-print-section-title">SCORE GLOBAL</h2>
                <table className="agent-print-table agent-print-summary">
                  <tbody>
                    <tr><th>Total prévu</th><td>{bulletinSections.totalPrevu}</td></tr>
                    <tr><th>Total réalisé</th><td>{bulletinSections.totalRealise}</td></tr>
                    <tr><th>Écart</th><td>{bulletinSections.totalRealise - bulletinSections.totalPrevu}</td></tr>
                    <tr><th>Score</th><td><strong>{bulletinSections.scoreGlobal}%</strong></td></tr>
                  </tbody>
                </table>
              </>
            )}
          </div>
        </>
      )}
    </>
  )
}


/* ══════════════════════════════════════════════════════════════════════════
 * COMPOSANT — Tableau d'une section du bulletin
 *
 * Rend une section dans le format métier :
 *   - Titre de section (bandeau)
 *   - En-tête : Critères | Prévu | Réalisé | Écart
 *   - Lignes critère
 *   - Ligne TAUX DE CONFORMITE (totaux)
 *   - Score % à droite du titre
 *
 * Affiche un message neutre si la section n'a aucune ligne (cas d'un agent
 * sans contrôle assigné par exemple).
 * ══════════════════════════════════════════════════════════════════════════ */
function BulletinSectionTable({ section }: { section: BulletinSection }) {
  return (
    <>
      <h2 className="agent-print-section-title bulletin-section-title">
        <span>{section.titre}</span>
        <span className="bulletin-score">Score : {section.score}%</span>
      </h2>
      {section.rows.length === 0 ? (
        <p style={{ fontSize: 11, color: '#666', margin: '4px 0 8px' }}>
          Aucun élément pour cette section.
        </p>
      ) : (
        <table className="agent-print-table bulletin-section-table">
          <thead>
            <tr>
              <th>CRITERES DE NOTATION</th>
              <th style={{ width: 70, textAlign: 'center' }}>Prévu</th>
              <th style={{ width: 70, textAlign: 'center' }}>Réalisé</th>
              <th style={{ width: 70, textAlign: 'center' }}>Écart</th>
            </tr>
          </thead>
          <tbody>
            {section.rows.map((r, idx) => (
              <tr key={idx}>
                <td>{r.critere}</td>
                <td style={{ textAlign: 'center' }}>{r.prevu}</td>
                <td style={{ textAlign: 'center' }}>{r.realise}</td>
                <td style={{ textAlign: 'center', color: r.ecart < 0 ? '#c0392b' : '#10b981' }}>
                  {r.ecart > 0 ? `+${r.ecart}` : r.ecart}
                </td>
              </tr>
            ))}
            {/* Ligne totale "TAUX DE CONFORMITE" — fond contrasté, gras */}
            <tr className="bulletin-section-total">
              <th>TAUX DE CONFORMITE</th>
              <th style={{ textAlign: 'center' }}>{section.totalPrevu}</th>
              <th style={{ textAlign: 'center' }}>{section.totalRealise}</th>
              <th style={{ textAlign: 'center', color: section.totalEcart < 0 ? '#c0392b' : '#10b981' }}>
                {section.totalEcart > 0 ? `+${section.totalEcart}` : section.totalEcart}
              </th>
            </tr>
          </tbody>
        </table>
      )}
    </>
  )
}
