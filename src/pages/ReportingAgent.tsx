/**
 * ============================================================================
 * REPORTING PAR AGENT — STATISTIQUES PAR PERSONNE (AFFECTÉE OU AUTEUR)
 * ============================================================================
 *
 * Vue agrégée des anomalies regroupées par "agent" — la sémantique d'"agent"
 * est pilotée par le sélecteur `groupingMode` :
 *
 *   - 'affecte' (DÉFAUT) : agent = personne AFFECTÉE (personneAffecter)
 *     → mesure la CHARGE des contrôleurs ("qui gère quoi")
 *
 *   - 'auteur' : agent = AUTEUR de l'anomalie (auteur_anormalie)
 *     → mesure la PRODUCTION de signalements ("qui a déclaré quoi")
 *
 * Le bulletin imprimable s'adapte automatiquement au mode (titre différent :
 * "Fiche de notation du contrôleur" vs "Fiche de notation de l'auteur").
 *
 * ARCHITECTURE EN DEUX VUES
 * -------------------------
 *   1. Vue principale (selectedAgent === null) :
 *      - Sélecteur de mode de regroupement
 *      - Stats cards globales (Agents, Total anomalies, Montant total)
 *      - Barre de filtres (saisie/applied : application différée au clic)
 *      - Tableau des agents : colonnes Personne, Email, Total/Ouvert/EnCours/
 *        Resolu/Clos, Montant, Domaines d'activité (chips bleues),
 *        Types d'action/sanction (chips ambrées)
 *
 *   2. Vue détail (selectedAgent !== null) :
 *      - Header : nom + email de l'agent + bouton Imprimer le bulletin
 *      - Barre de filtres DÉTAIL (filtres propres à la vue agent)
 *      - Stats cards spécifiques à l'agent (recalculées sur la liste filtrée)
 *      - Tableau des anomalies de l'agent (paginé, scrollable)
 *      - Bouton Retour à la liste
 *
 * BULLETIN IMPRIMABLE (.agent-print-only)
 * ---------------------------------------
 * Caché à l'écran, visible uniquement à l'impression. Format "Fiche de
 * notation" inspiré du tableau de bord DCPO 2026 :
 *   - Header : titre + nom agent + date d'édition
 *   - Sections (cf. BulletinSection) :
 *       • Anomalies (par classification)
 *       • Plans de Contrôle (par contrôle assigné)
 *       • Plans d'Action Correctif (par PAC assigné)
 *     Chaque section affiche un tableau Critère / Prévu / Réalisé / Écart,
 *     une ligne TAUX DE CONFORMITE (totaux) et un score % en haut.
 *   - Score global = moyenne pondérée des scores de section
 *
 * DONNÉES CHARGÉES AU MONTAGE
 * ---------------------------
 *   - Anomalies (DCPO_LISTE_ANORMALIE)
 *   - Agences + Réseaux (référentiels pour résolution des libellés)
 *   - Plans de Contrôle (pour la section bulletin)
 *   - Toutes les évaluations de contrôle (pour le compteur réalisé)
 *   - PACs (pour la section bulletin)
 *
 * Filtrage 100 % CÔTÉ CLIENT (pas d'OData) — tout est en mémoire après le
 * fetch initial. Adapté pour < 5 000 anomalies (limite SP par défaut).
 *
 * PERMISSIONS
 * -----------
 * Tous les rôles autorisés peuvent voir cette page (cf. ALLOWED_ROLES dans
 * App.tsx). Il n'y a pas de filtrage rôle-dépendant côté serveur — un
 * contrôleur voit donc les stats de TOUS les autres agents. Si on veut
 * restreindre, ajouter un filtre OData `personneAffecter/Email eq <self>`
 * dans fetchAnomalies (cf. pattern dans Anomalies.tsx).
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { loadAgences, loadReseaux } from '../lib/spReferenceRows'
import { getAllPages } from '../lib/sharePointPaging'
import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { ExportButtons } from '../components/ExportButtons'
import { formatMontantCompact, formatDateOnlyFR } from '../lib/formatters'
import {
  listControles,
  listAllEvaluations,
  getExpectedEvaluationsPerYear,
  type ControleEntry,
} from '../lib/planControleService'
import { listPACs, type Pac } from '../lib/pacService'
import { buildConsolidatedBulletin } from '../lib/anomalyBulletin'
import { BulletinModal } from './AnomalyBulletins'

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
  /**
   * Valeurs DISTINCTES de domaine d'activité rencontrées sur les anomalies
   * de cet agent. Triées par ordre alphabétique pour un affichage stable.
   */
  domaines: string[]
  /**
   * Valeurs DISTINCTES de typeSanction rencontrées sur les anomalies clôturées
   * de cet agent (le champ n'est renseigné qu'à la clôture, donc cette liste
   * peut rester vide si aucune anomalie n'est clos).
   */
  sanctions: string[]
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
  /**
   * Anomalie sélectionnée pour la modale récapitulative (déclenchée par
   * clic sur le N° de ticket dans le tableau détail). `null` = pas de modale.
   * On réutilise le composant `BulletinModal` déjà défini dans
   * AnomalyBulletins.tsx pour garder la même UX de fiche récap partout.
   */
  const [detailAnomaly, setDetailAnomaly] = useState<DCPO_LISTE_ANORMALIERead | null>(null)
  /**
   * Mode de regroupement des anomalies :
   *   - 'affecte' : par contrôleur affecté (personneAffecter) — par défaut
   *   - 'auteur'  : par auteur de l'anomalie (auteur_anormalie)
   *
   * Le changement de mode rebuild agentMap et reset la sélection courante
   * (le contexte change : "l'agent X" en mode affecte ≠ "l'agent X" en mode
   * auteur). Permet d'imprimer le bulletin selon les deux axes.
   */
  const [groupingMode, setGroupingMode] = useState<'affecte' | 'auteur'>('affecte')
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
        const [anomRows, agencesRows, reseauxRows, controlesRes, pacsRes, evalsRes] = await Promise.all([
          getAllPages<DCPO_LISTE_ANORMALIERead>(DCPO_LISTE_ANORMALIEService),
          loadAgences(),
          loadReseaux(),
          listControles(),
          listPACs(),
          listAllEvaluations(),
        ])
        setItems(anomRows)
        setAgences(agencesRows)
        setReseaux(reseauxRows)
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
   * Sémantique métier pilotée par `groupingMode` :
   *   - 'affecte' (par défaut) → contrôleur AFFECTÉ (personneAffecter) :
   *     mesure la CHARGE de chaque contrôleur, qui s'occupe de quoi.
   *   - 'auteur' → AUTEUR de l'anomalie (auteur_anormalie) : mesure la
   *     PRODUCTION de signalements par agent (qui a déclaré quoi).
   *
   * Permet d'imprimer le bulletin selon les DEUX axes (cf. sélecteur en
   * haut de page).
   */
  const agentMap = useMemo(() => {
    const map = new Map<string, AgentStats>()
    filteredItems.forEach(item => {
      // Choix de la personne de référence selon le mode actif
      const personne = groupingMode === 'auteur'
        ? item.auteur_anormalie
        : item.personneAffecter
      const email = personne?.Email ?? ''
      const name = personne?.DisplayName ?? 'Inconnu'
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
          domaines: [],
          sanctions: [],
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

      // Domaines & sanctions : on collecte les valeurs distinctes pour
      // afficher chips en colonnes dédiées. Les chaînes vides / espaces
      // sont ignorées pour ne pas polluer la liste.
      const dom = item.domaineActivite?.trim()
      if (dom && !stats.domaines.includes(dom)) stats.domaines.push(dom)
      const san = item.typeSanction?.trim()
      if (san && !stats.sanctions.includes(san)) stats.sanctions.push(san)
    })
    // Tri alphabétique stable des listes distinctes (cohérent visuellement
    // d'un agent à l'autre — sinon l'ordre dépendrait de l'ordre de
    // traitement des items).
    for (const stats of map.values()) {
      stats.domaines.sort()
      stats.sanctions.sort()
    }
    return map
  }, [filteredItems, groupingMode])

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
      <div className="content-header" style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <h2 style={{ flex: 1, margin: 0 }}>
          Reporting par {groupingMode === 'auteur' ? 'auteur' : 'agent affecté'}
        </h2>
        {/* Sélecteur de regroupement : pilote agentMap → permet d'imprimer
            le bulletin selon deux axes (charge des contrôleurs vs production
            de signalements par auteur). Masqué en vue détail pour ne pas
            laisser changer le contexte d'une fiche déjà ouverte. */}
        {!selectedAgent && (
          <div className="filter-field" style={{ marginBottom: 0 }}>
            <label htmlFor="grouping-mode" style={{ fontSize: 12, marginBottom: 4 }}>Regroupement</label>
            <select
              id="grouping-mode"
              value={groupingMode}
              onChange={e => setGroupingMode(e.target.value as 'affecte' | 'auteur')}
            >
              <option value="affecte">Par personne affectée</option>
              <option value="auteur">Par auteur</option>
            </select>
          </div>
        )}
        {!selectedAgent && (
          <ExportButtons
            filename={groupingMode === 'auteur' ? 'reporting_par_auteur' : 'reporting_par_agent'}
            pdfTitle={`Reporting par ${groupingMode === 'auteur' ? 'auteur' : 'agent affecté'}`}
            getHeaders={() => [
              'Agent', 'Email', 'Total anomalies', 'Ouvert', 'En cours',
              'Résolu', 'Clos', 'Montant total', 'Domaines',
            ]}
            getRows={() => agents.map(a => [
              a.displayName,
              a.email,
              a.total,
              a.ouvert,
              a.enCours,
              a.resolu,
              a.clos,
              a.montantTotal,
              a.domaines.join(', '),
            ])}
          />
        )}
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
                    <th>{groupingMode === 'auteur' ? 'Auteur' : 'Personne affectée'}</th>
                    <th>Email</th>
                    <th>Total</th>
                    <th>Ouvert</th>
                    <th>En cours</th>
                    <th>Resolu</th>
                    <th>Clos</th>
                    <th>Montant total</th>
                    <th style={{ minWidth: 180 }}>Domaines d'activité</th>
                    <th style={{ minWidth: 180 }}>mode de traitement</th>
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
                      {/* Domaines distincts en chips bleues */}
                      <td>
                        {agent.domaines.length === 0
                          ? <span style={{ color: '#888' }}>—</span>
                          : (
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                              {agent.domaines.map(d => (
                                <span
                                  key={d}
                                  style={{
                                    fontSize: 11,
                                    padding: '2px 6px',
                                    background: '#dbeafe',
                                    color: '#1e3a8a',
                                    border: '1px solid #93c5fd',
                                    borderRadius: 10,
                                    whiteSpace: 'nowrap',
                                  }}
                                >
                                  {d}
                                </span>
                              ))}
                            </div>
                          )
                        }
                      </td>
                      {/* Sanctions distinctes en chips ambrées */}
                      <td>
                        {agent.sanctions.length === 0
                          ? <span style={{ color: '#888' }}>—</span>
                          : (
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                              {agent.sanctions.map(s => (
                                <span
                                  key={s}
                                  style={{
                                    fontSize: 11,
                                    padding: '2px 6px',
                                    background: '#fef3c7',
                                    color: '#92400e',
                                    border: '1px solid #fcd34d',
                                    borderRadius: 10,
                                    whiteSpace: 'nowrap',
                                  }}
                                >
                                  {s}
                                </span>
                              ))}
                            </div>
                          )
                        }
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
                  {/* N° Ticket : identifiant SharePoint formaté T-{ID} —
                      cohérent avec le numéro affiché dans Anomalies.tsx et
                      les bulletins clôturés. Sticky à gauche pour rester
                      visible lors du scroll horizontal du tableau. */}
                  <th>N° Ticket</th>
                  <th>Date</th>
                  <th>Declarant</th>
                  <th>Auteur</th>
                  <th>Cause</th>
                  <th>Classification</th>
                  <th>Domaine d'activité</th>
                  <th>Agence</th>
                  <th>Reseau</th>
                  <th>Montant</th>
                  <th>Date regularisation</th>
                  <th>Statut</th>
                  <th>Mode de traitement</th>
                  <th>Personne affectee</th>
                </tr>
              </thead>
              <tbody>
                {pagedDetailAnomalies.map(item => {
                  const numero = item.ID ? `T-${item.ID}` : '—'
                  return (
                  <tr key={item.ID}>
                    <td>
                      {item.ID ? (
                        <button
                          type="button"
                          onClick={() => setDetailAnomaly(item)}
                          title="Voir le récapitulatif de l'anomalie"
                          style={{
                            background: 'none',
                            border: 'none',
                            padding: 0,
                            color: '#1d4ed8',
                            fontWeight: 700,
                            textDecoration: 'underline',
                            cursor: 'pointer',
                            font: 'inherit',
                          }}
                        >
                          {numero}
                        </button>
                      ) : (
                        <strong>{numero}</strong>
                      )}
                    </td>
                    <td>{formatDateOnlyFR(item.field_0, '-')}</td>
                    <td>{item.declarant_anormalie?.DisplayName ?? '-'}</td>
                    <td>{item.auteur_anormalie?.DisplayName ?? '-'}</td>
                    <td>{item.field_4 ? stripHtml(item.field_4) : '-'}</td>
                    <td>{item.field_5 ?? '-'}</td>
                    <td>{item.domaineActivite ?? '-'}</td>
                    <td>{agences.find(a => String(a.ID) === item.field_6)?.Title ?? item.field_6 ?? '-'}</td>
                    <td>{reseaux.find(r => String(r.ID) === item.field_7)?.field_1 ?? item.field_7 ?? '-'}</td>
                    <td>{item.field_8?.toLocaleString() ?? '-'}</td>
                    <td>{formatDateOnlyFR(item.field_9, '-')}</td>
                    <td>{item.field_10 ?? '-'}</td>
                    <td>{item.typeSanction ?? '-'}</td>
                    <td>{item.personneAffecter?.DisplayName ?? '-'}</td>
                  </tr>
                  )
                })}
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
              {/* Le titre du bulletin reflète le mode de regroupement :
                  - 'affecte' → "FICHE DE NOTATION DU CONTROLEUR" (charge)
                  - 'auteur'  → "FICHE DE NOTATION DE L'AUTEUR" (production) */}
              <h1>
                {groupingMode === 'auteur'
                  ? "FICHE DE NOTATION DE L'AUTEUR"
                  : 'FICHE DE NOTATION DU CONTROLEUR'}
              </h1>
              <p style={{ textAlign: 'center', color: 'var(--rdcpo-red)' }}>
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

      {/* ─── Modale récap anomalie (clic sur N° Ticket) ─────────────
          Réutilise le composant BulletinModal d'AnomalyBulletins. On
          construit à la volée le bulletin consolidé (mapping ID agence/
          réseau → libellés, timeline, causes parsées, etc.). Bouton
          Imprimer → window.print() comme sur la vue Anomalies résolues. */}
      {detailAnomaly && (
        <BulletinModal
          bulletin={buildConsolidatedBulletin(detailAnomaly, agences, reseaux)}
          onClose={() => setDetailAnomaly(null)}
          onPrint={() => window.print()}
        />
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
  // Libellés métier spécifiques : pour la section "Surveillance des anomalies",
  // les contrôleurs notent en termes de détection/correction plutôt que de
  // prévisionnel/réalisé. Les autres sections (plans de contrôle, PAC) gardent
  // les libellés génériques.
  const isAnomaliesSection = /ANOMALIES/i.test(section.titre)
  const prevuLabel = isAnomaliesSection ? 'Détectée' : 'Prévu'
  const realiseLabel = isAnomaliesSection ? 'Corrigée' : 'Réalisé'
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
              <th style={{ width: 70, textAlign: 'center' }}>{prevuLabel}</th>
              <th style={{ width: 70, textAlign: 'center' }}>{realiseLabel}</th>
              <th style={{ width: 70, textAlign: 'center' }}>Écart</th>
            </tr>
          </thead>
          <tbody>
            {section.rows.map((r, idx) => (
              <tr key={idx}>
                <td>{r.critere}</td>
                <td style={{ textAlign: 'center' }}>{r.prevu}</td>
                <td style={{ textAlign: 'center' }}>{r.realise}</td>
                <td style={{ textAlign: 'center', color: r.ecart < 0 ? 'var(--rdcpo-red)' : '#10b981' }}>
                  {r.ecart > 0 ? `+${r.ecart}` : r.ecart}
                </td>
              </tr>
            ))}
            {/* Ligne totale "TAUX DE CONFORMITE" — fond contrasté, gras */}
            <tr className="bulletin-section-total">
              <th>TAUX DE CONFORMITE</th>
              <th style={{ textAlign: 'center' }}>{section.totalPrevu}</th>
              <th style={{ textAlign: 'center' }}>{section.totalRealise}</th>
              <th style={{ textAlign: 'center', color: section.totalEcart < 0 ? 'var(--rdcpo-red)' : '#10b981' }}>
                {section.totalEcart > 0 ? `+${section.totalEcart}` : section.totalEcart}
              </th>
            </tr>
          </tbody>
        </table>
      )}
    </>
  )
}
