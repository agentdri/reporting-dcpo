/**
 * ============================================================================
 * MODULE 3 (volet manager) — VALIDATION DES RAPPORTS QUOTIDIENS
 * ============================================================================
 *
 * Page accessible UNIQUEMENT aux managers (Chef_Departement / Directeur).
 * Permet de valider ou invalider les rapports soumis par les contrôleurs.
 *
 * Fonctionnalités :
 *   - Liste consolidée des rapports, regroupés par (contrôleur, date)
 *   - Filtres : nom/email contrôleur, période, "en attente uniquement"
 *   - Stats : compteurs par statut (Soumis / Valider / Refuser), total heures
 *   - Modale détail : résumé du rapport + lignes + lien rapport + observations
 *     + pièces jointes + historique des décisions managers passées
 *   - Actions managers :
 *     - Valider → statut 'Valider', avec note horodatée optionnelle
 *     - Invalider → statut 'Refuser', motif obligatoire (saisi en textarea)
 *
 * Persistance :
 *   - Lecture : listReports() depuis activityService (SharePoint DCPO_ACTIVICTE_CONTROLLER)
 *   - Validation : validateReport() écrase les colonnes statutValidation et
 *     motifRejet du rapport (un seul statut + un seul motif à la fois,
 *     pas d'historique). La décision la plus récente prime.
 * ============================================================================
 */

import { Fragment, useCallback, useEffect, useMemo, useState } from 'react'
import {
  listReports,
  validateReport,
  type ActivityReport,
} from '../lib/activityService'
import { DCPO_LISTE_USERService } from '../generated/services/DCPO_LISTE_USERService'
import { getAllPages } from '../lib/sharePointPaging'
import type { DCPO_LISTE_USERRead } from '../generated/models/DCPO_LISTE_USERModel'
import { listAbsences, buildAbsentDaysMap, type Absence } from '../lib/absenceService'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { ProgressBar } from '../components/ProgressBar'
import { ExportButtons } from '../components/ExportButtons'
import { formatDateForExport } from '../lib/exporters'
import './ControllerReporting.css'
import { ModalOverlay } from '../components/ModalOverlay'

/** Identité du manager courant (passée par Dashboard). */
interface ControllerReportingListProps {
  userName?: string
  userEmail?: string
}

/** Critères de filtrage de la liste des rapports. */
interface FilterState {
  controleur: string   // recherche partielle nom OU email
  dateFrom: string     // date min du rapport (YYYY-MM-DD)
  dateTo: string       // date max du rapport
  pendingOnly: boolean // ne montrer QUE les rapports en attente de validation
}

const EMPTY_FILTERS: FilterState = {
  controleur: '',
  dateFrom: '',
  dateTo: '',
  pendingOnly: false,
}

/**
 * Clé d'un groupe dans la vue regroupée : un groupe = (contrôleur, date).
 * Permet d'afficher tous les rapports d'un contrôleur pour un même jour
 * sous un même header (généralement il n'y en a qu'un, mais le manager peut
 * avoir reçu plusieurs envois).
 */
interface GroupKey {
  controleurEmail: string
  controleurName: string
  date: string
}

/**
 * Données agrégées d'un groupe : la liste des rapports + les compteurs
 * par statut + total heures cumulées.
 *
 * Pourquoi un groupe ? Le manager raisonne par "qui a soumis quoi tel jour"
 * — l'unité utile n'est pas le rapport individuel mais la JOURNÉE-CONTRÔLEUR.
 * D'où l'agrégation par couple (controleurEmail, date) avant l'affichage.
 *
 * Compteurs :
 *   - pending  : statut "Soumis" (en attente de décision manager)
 *   - realise  : statut "Valider" (le manager a validé la journée)
 *   - reporte  : statut "Refuser" (le manager a refusé — motif obligatoire,
 *                cf. validateReport dans activityService)
 *
 * Note : on ne stocke jamais l'historique des décisions (un seul statut +
 * un seul motif à la fois). La décision la plus récente écrase la précédente.
 */
interface Group {
  key: GroupKey
  reports: ActivityReport[]
  totalHeures: number
  pending: number
  realise: number
  reporte: number
}

/** Formatte un nombre d'heures avec 2 décimales max (locale FR). */
function formatHours(value: number): string {
  return `${value.toLocaleString('fr-FR', { maximumFractionDigits: 2 })} h`
}

/**
 * Compte le nombre de jours OUVRÉS (Lun-Ven) entre deux dates INCLUSIVES.
 *
 * Hypothèse simplifiée : les jours fériés camerounais/français ne sont PAS
 * exclus (référentiel non disponible côté app). L'attendu peut donc être
 * légèrement surestimé sur les périodes contenant des jours fériés — au
 * bénéfice du manager (taux plus sévère). Si un référentiel de jours
 * fériés est ajouté un jour, cette fonction est le point d'entrée à
 * modifier.
 *
 * Retourne 0 si `to < from` (période vide/inversée).
 */
/**
 * Retourne la LISTE des jours ouvrés d'une période au format 'YYYY-MM-DD'.
 * Utilisé pour :
 *   - Le compteur d'attendus (`length`)
 *   - Le calcul des jours MANQUANTS d'un contrôleur (différence entre
 *     attendus et couverts)
 */
function listBusinessDays(from: Date, to: Date): string[] {
  if (to < from) return []
  const days: string[] = []
  const cursor = new Date(from)
  cursor.setHours(0, 0, 0, 0)
  const end = new Date(to)
  end.setHours(0, 0, 0, 0)
  while (cursor <= end) {
    const day = cursor.getDay() // 0 = Dimanche, 6 = Samedi
    if (day !== 0 && day !== 6) {
      // Format YYYY-MM-DD sans passer par toISOString() (qui convertit en UTC
      // et pourrait décaler d'un jour selon le fuseau).
      const y = cursor.getFullYear()
      const m = String(cursor.getMonth() + 1).padStart(2, '0')
      const d = String(cursor.getDate()).padStart(2, '0')
      days.push(`${y}-${m}-${d}`)
    }
    cursor.setDate(cursor.getDate() + 1)
  }
  return days
}

/** Identité minimale d'un contrôleur (pour la table de taux de soumission). */
interface Controleur {
  email: string
  name: string
}

/** Date jj/mm/aaaa, "—" si vide ou invalide. */
function formatDate(value?: string): string {
  if (!value) return '—'
  try { return new Date(value).toLocaleDateString('fr-FR') } catch { return value }
}

/** Date+heure, '' si vide ou invalide (vide = pas de fallback affiché). */
function formatDateTime(value?: string): string {
  if (!value) return ''
  try { return new Date(value).toLocaleString('fr-FR') } catch { return value }
}

export default function ControllerReportingList({ userName, userEmail }: ControllerReportingListProps) {
  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS
   * ────────────────────────────────────────────────────────────────────── */

  /** Liste des rapports — chargée via useEffect (async SharePoint). */
  const [reports, setReports] = useState<ActivityReport[]>([])
  /**
   * Liste des CONTRÔLEURS actifs (DCPO_LISTE_USER filtrés sur fonction=Controleur).
   * Sert de dénominateur pour le calcul du taux de soumission — chaque
   * contrôleur doit soumettre 1 rapport par jour ouvré. On charge cette
   * liste une seule fois au montage (les changements de rôle sont rares
   * et n'ont pas besoin d'être rechargés en continu).
   */
  const [controleurs, setControleurs] = useState<Controleur[]>([])
  /**
   * Liste des absences (DCPO_LISTE_ABSENCES). Chargée une seule fois au
   * montage, rafraîchie via `refresh()` en même temps que les rapports.
   * Sert à retirer les jours d'absence du dénominateur du taux de
   * soumission (cf. `absentDaysMap` dans submissionStats).
   */
  const [absences, setAbsences] = useState<Absence[]>([])
  /** Indicateur de chargement initial (réseau en cours). */
  const [loadingReports, setLoadingReports] = useState(true)
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** Rapport sélectionné pour la modale détail (null = pas de modale). */
  const [selected, setSelected] = useState<ActivityReport | null>(null)
  /**
   * Toggle d'affichage de la section "Taux de soumission des rapports".
   * Fermée par défaut : le manager clique sur le header pour la déplier.
   * Évite de charger visuellement la page à l'ouverture — la validation
   * reste le workflow principal, le taux est une vue analytique secondaire.
   */
  const [ratesExpanded, setRatesExpanded] = useState(false)
  /**
   * Email du contrôleur dont le détail (jours manquants) est déplié dans
   * le tableau de taux. `null` = aucun détail visible. Un seul détail à
   * la fois (accordéon exclusif) pour garder la lisibilité.
   */
  const [expandedControleurEmail, setExpandedControleurEmail] = useState<string | null>(null)

  /**
   * Recharge la liste depuis SharePoint en filtrant CÔTÉ SERVEUR sur la
   * plage de dates saisie dans les filtres. Pousser le filtre OData au
   * niveau du service évite deux écueils :
   *   1. Ramener des milliers d'items pour n'en afficher que quelques-uns
   *      (coût réseau + limite de pagination du connecteur SP).
   *   2. Le comportement pathologique du connecteur SP qui, sur certaines
   *      listes, ignore `$skip` et remonte 100 000 doublons (cf. le fix
   *      dans sharePointPaging.ts) — un filtre restrictif ramène la liste
   *      en 1 page et court-circuite complètement le problème.
   *
   * Si l'utilisateur n'a saisi aucune date, `listReports()` retombe sur
   * son défaut = année en cours (cf. defaultCurrentYearRange).
   */
  const refresh = useCallback(async () => {
    setLoadingReports(true)
    try {
      // Rapports ET absences en parallèle — les 2 alimentent le calcul de taux.
      const [reportsData, absencesData] = await Promise.all([
        listReports({
          from: filters.dateFrom || undefined,
          to: filters.dateTo || undefined,
        }),
        listAbsences(),
      ])
      setReports(reportsData)
      setAbsences(absencesData)
    } finally {
      setLoadingReports(false)
    }
  }, [filters.dateFrom, filters.dateTo])

  // Chargement initial au montage + refetch dès que la plage de dates change.
  // La stabilité du callback via useCallback+deps garantit qu'on ne refetch
  // QUE quand from/to bougent, pas sur les autres filtres (contrôleur, statut)
  // qui restent traités côté client.
  useEffect(() => {
    refresh()
  }, [refresh])

  /**
   * Chargement UNIQUE au montage de la liste des contrôleurs actifs.
   * Non rerun sur refresh() car cette liste évolue rarement (ajout/départ
   * d'un contrôleur) — l'utilisateur peut recharger la page pour la
   * rafraîchir si besoin.
   */
  useEffect(() => {
    let cancelled = false
    getAllPages<DCPO_LISTE_USERRead>(DCPO_LISTE_USERService, { orderBy: ['nom asc'] })
      .then(rows => {
        if (cancelled) return
        const list: Controleur[] = rows
          .filter((u: DCPO_LISTE_USERRead) => u.fonction?.Value === 'Controleur')
          .map((u: DCPO_LISTE_USERRead) => ({
            email: (u.Email ?? '').toLowerCase(),
            name: u.Title || u.nom || u.Email || '—',
          }))
          .filter((c: Controleur) => !!c.email)
        setControleurs(list)
      })
      .catch(err => console.error('Erreur chargement contrôleurs', err))
    return () => { cancelled = true }
  }, [])

  /**
   * Applique tous les filtres sur la liste brute.
   *
   * Ordre des tests (early return false) :
   *   1. pendingOnly (booléen le moins coûteux)
   *   2. controleur (recherche texte)
   *   3. dates (parsing + comparaison)
   *
   * useMemo : recalcul uniquement si reports OU filters change.
   */
  const filtered = useMemo(() => {
    const fromDate = filters.dateFrom ? new Date(`${filters.dateFrom}T00:00:00`) : undefined
    const toDate = filters.dateTo ? new Date(`${filters.dateTo}T23:59:59`) : undefined
    const search = filters.controleur.trim().toLowerCase()
    return reports.filter(r => {
      if (filters.pendingOnly && r.statut !== 'Soumis') return false
      if (search) {
        const haystack = `${r.controleurEmail} ${r.controleurName}`.toLowerCase()
        if (!haystack.includes(search)) return false
      }
      if (fromDate || toDate) {
        const rd = new Date(r.date)
        if (Number.isNaN(rd.getTime())) return false
        if (fromDate && rd < fromDate) return false
        if (toDate && rd > toDate) return false
      }
      return true
    })
  }, [reports, filters])

  /**
   * Regroupe les rapports filtrés par (contrôleur, date).
   *
   * Algorithme :
   *   1. Map<key, Group> où key = `${email}|${date}`
   *   2. Pour chaque rapport : créer le groupe si nécessaire, accumuler
   *   3. Trier les groupes :
   *      - D'abord par date desc (plus récent en haut)
   *      - Puis par nom de contrôleur (alpha)
   */
  const groups = useMemo<Group[]>(() => {
    const map = new Map<string, Group>()
    for (const r of filtered) {
      const k = `${r.controleurEmail}|${r.date}`
      if (!map.has(k)) {
        map.set(k, {
          key: { controleurEmail: r.controleurEmail, controleurName: r.controleurName, date: r.date },
          reports: [],
          totalHeures: 0,
          pending: 0,
          realise: 0,
          reporte: 0,
        })
      }
      const g = map.get(k)!
      g.reports.push(r)
      g.totalHeures += r.totalHeures
      if (r.statut === 'Soumis') g.pending++
      else if (r.statut === 'Valider') g.realise++
      else if (r.statut === 'Refuser') g.reporte++
    }
    return Array.from(map.values()).sort((a, b) => {
      if (a.key.date !== b.key.date) return a.key.date < b.key.date ? 1 : -1
      return a.key.controleurName.localeCompare(b.key.controleurName)
    })
  }, [filtered])

  /**
   * Pagination — porte sur les GROUPES (1 carte = 1 contrôleur+date).
   * Plus naturel qu'au niveau du rapport individuel car les managers
   * scrollent par contrôleur/jour, pas par rapport unitaire.
   */
  const pagination = usePagination({
    total: groups.length,
    resetKey: JSON.stringify(filters),
  })
  const pagedGroups = useMemo(
    () => groups.slice(pagination.start, pagination.end),
    [groups, pagination.start, pagination.end],
  )

  /**
   * Stats globales pour les cards en haut de page.
   * Recalculées sur la liste filtrée (reflète les filtres en cours).
   */
  const totals = useMemo(() => {
    return {
      total: filtered.length,
      pending: filtered.filter(r => r.statut === 'Soumis').length,
      realise: filtered.filter(r => r.statut === 'Valider').length,
      reporte: filtered.filter(r => r.statut === 'Refuser').length,
      totalHeures: filtered.reduce((s, r) => s + r.totalHeures, 0),
    }
  }, [filtered])

  /**
   * TAUX DE SOUMISSION DES RAPPORTS PAR CONTRÔLEUR
   *
   * Métrique : sur une période donnée, chaque contrôleur est censé soumettre
   * 1 rapport par jour ouvré (Lun-Ven). Le taux mesure la couverture réelle :
   *
   *   taux = (nb de jours DISTINCTS couverts par un rapport non refusé)
   *          / (nb de jours ouvrés dans la période)
   *
   * Règles :
   *   - On compte les rapports "Soumis" et "Valider" (les Refuser sont
   *     exclus — ils doivent être resoumis pour compter).
   *   - Plusieurs rapports pour la même date d'un même contrôleur = 1 date
   *     couverte (déduplication via Set<date>).
   *   - Période par défaut (aucun filtre date) = 30 derniers jours.
   *   - dateTo est capée à `today` — on ne veut pas pénaliser sur des
   *     jours ouvrés futurs.
   *   - Contrôleurs sans aucun rapport dans la période apparaissent avec
   *     taux = 0 % (rendus visibles pour signaler l'absence complète).
   *
   * Limite connue : jours fériés non exclus du dénominateur (cf. helper
   * countBusinessDays). Impact = léger sur-durcissement du taux.
   */
  const submissionStats = useMemo(() => {
    // Bornes de période : par défaut = 30 derniers jours (avec cap à aujourd'hui)
    // ⚠ `to` DOIT être en fin de journée (23:59:59.999) et non minuit — sinon
    // les rapports timestampés dans la journée du `to` sont exclus par le
    // test `rd > to` alors qu'ils devraient être inclus (cas d'un rapport
    // du 07/07 stocké '2026-07-07T00:00:00Z' → 01:00 local en UTC+1, qui
    // dépasse `to` = 07/07 00:00 local si on ne fait pas l'ajustement).
    const today = new Date()
    today.setHours(0, 0, 0, 0)
    const eodToday = new Date(today)
    eodToday.setHours(23, 59, 59, 999)
    const defaultFrom = new Date(today)
    defaultFrom.setDate(defaultFrom.getDate() - 30)

    const from = filters.dateFrom ? new Date(`${filters.dateFrom}T00:00:00`) : defaultFrom
    const rawTo = filters.dateTo ? new Date(`${filters.dateTo}T23:59:59.999`) : eodToday
    // Cap : on n'attend pas des soumissions futures (mais on garde l'EOD)
    const to = rawTo > eodToday ? eodToday : rawTo

    const expectedDaysList = listBusinessDays(from, to)
    // NB : `expectedDays` reste la "capacité BRUTE" de la période (avant
    // déduction des absences). Elle est conservée pour information mais
    // le dénominateur EFFECTIF de chaque contrôleur retire ses jours
    // d'absence (cf. plus bas `expectedEffectifSet`).
    const expectedDays = expectedDaysList.length

    // Coverage : email contrôleur → Set des dates DISTINCTES couvertes
    // par un rapport statut != 'Refuser' DANS la période [from, to].
    const coverage = new Map<string, Set<string>>()
    for (const r of reports) {
      if (r.statut === 'Refuser') continue
      const rd = new Date(r.date)
      if (Number.isNaN(rd.getTime())) continue
      if (rd < from || rd > to) continue
      const email = (r.controleurEmail ?? '').toLowerCase()
      if (!email) continue
      if (!coverage.has(email)) coverage.set(email, new Set())
      coverage.get(email)!.add(r.date.slice(0, 10))
    }

    // Absences : Map<email, Set<jour>> pour la période. Les jours listés
    // ici seront RETIRÉS du dénominateur du contrôleur concerné (congé,
    // maladie, formation, mission → pas de rapport attendu).
    const absentDaysMap = buildAbsentDaysMap(absences, from, to)

    // Fusion : on part de la liste RÉFÉRENTIEL (DCPO_LISTE_USER) pour ne
    // pas oublier les contrôleurs qui n'ont RIEN soumis. Fallback si la
    // liste des contrôleurs n'est pas encore chargée : on utilise les
    // emails distincts trouvés dans les rapports (moins rigoureux mais
    // affiche quelque chose plutôt qu'un tableau vide).
    const source: Controleur[] = controleurs.length > 0
      ? controleurs
      : Array.from(new Set(reports.map(r => (r.controleurEmail ?? '').toLowerCase()).filter(Boolean)))
          .map(email => {
            const r = reports.find(x => (x.controleurEmail ?? '').toLowerCase() === email)
            return { email, name: r?.controleurName || email }
          })

    // Cumul global : on additionne les couverts et les attendus EFFECTIFS
    // (après retrait des jours d'absence) pour un taux global significatif.
    let globalCoveredEff = 0
    let globalExpectedEff = 0

    const rows = source.map(c => {
      const coveredSet = coverage.get(c.email) ?? new Set<string>()
      const covered = coveredSet.size
      const absentSet = absentDaysMap.get(c.email) ?? new Set<string>()

      // Dénominateur EFFECTIF : jours ouvrés attendus MOINS les jours
      // d'absence du contrôleur (intersection déjà calculée par le service).
      const expectedEffectiveList = expectedDaysList.filter(d => !absentSet.has(d))
      const expectedEffective = expectedEffectiveList.length

      const taux = expectedEffective > 0
        ? Math.min(100, Math.round((covered / expectedEffective) * 100))
        : (covered > 0 ? 100 : 0) // aucun jour attendu (100 % d'absence) → 100 % si couvert, 0 sinon

      // Jours OUVRÉS attendus non couverts ET non absents = jours de rapport
      // réellement manquants pour ce contrôleur. Exclut les jours d'absence
      // pour ne pas signaler comme "manquants" des jours légitimement OFF.
      const missingDays = expectedEffectiveList.filter(d => !coveredSet.has(d))

      // Jours d'absence DANS la période analysée (pour affichage détail)
      const absentDays = Array.from(absentSet).sort()

      globalCoveredEff += covered
      globalExpectedEff += expectedEffective

      return {
        email: c.email,
        name: c.name,
        covered,
        expected: expectedEffective,
        expectedBrut: expectedDays,
        absent: absentDays.length,
        taux,
        coveredDays: Array.from(coveredSet).sort(),
        missingDays,
        absentDays,
      }
    }).sort((a, b) => a.taux - b.taux) // pires taux en premier (attire l'œil)

    // Taux global = ratio agrégé sur les dénominateurs EFFECTIFS
    const globalTaux = globalExpectedEff > 0
      ? Math.min(100, Math.round((globalCoveredEff / globalExpectedEff) * 100))
      : 0
    const globalCovered = globalCoveredEff
    const globalExpected = globalExpectedEff

    return {
      from,
      to,
      expectedDays,
      rows,
      globalTaux,
      globalCovered,
      globalExpected,
    }
  }, [reports, controleurs, absences, filters.dateFrom, filters.dateTo])

  /** Helper générique de mise à jour partielle des filtres. */
  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const resetFilters = () => setFilters(EMPTY_FILTERS)

  /**
   * Applique une décision manager (Valider ou Invalider) à un rapport.
   *
   * Règle métier : motif OBLIGATOIRE pour invalidation (Refuser).
   * Garde-fou défensif : alert + abort si motif vide.
   *
   * Après validation :
   *   - refresh() recharge la liste depuis SharePoint pour refléter le
   *     nouveau statut dans le tableau
   *   - setSelected(null) ferme automatiquement la modale (UX : le manager
   *     a pris sa décision, il n'a plus rien à faire sur cet item, on
   *     enchaîne sur le rapport suivant dans la liste)
   */
  const handleValidation = async (report: ActivityReport, decision: 'Valider' | 'Refuser', motif?: string) => {
    if (decision === 'Refuser' && !motif?.trim()) {
      alert('Un motif est requis pour invalider un reporting.')
      return
    }
    const updated = await validateReport(report.id, {
      manager: userName ?? 'Manager',
      managerEmail: userEmail,
      decision,
      motif: motif?.trim() || undefined,
    })
    if (updated) {
      await refresh()
      setSelected(null)
    }
  }

  return (
    <>
      <div className="content-header">
        <h2>Reportings soumis — Validation manager</h2>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <ExportButtons
            filename="reportings_soumis"
            pdfTitle="Reportings soumis — Validation manager"
            getHeaders={() => [
              'ID', 'Contrôleur', 'Email', 'Date activité', 'Statut',
              'Statut journée', 'Total heures', 'Temps occupé',
              'Anomalies détectées', 'Motif rejet', 'Soumis le',
            ]}
            getRows={() => filtered.map(r => [
              r.id,
              r.controleurName,
              r.controleurEmail,
              formatDateForExport(r.date),
              r.statut,
              r.statutJournee,
              r.totalHeures,
              r.tempsOccupe,
              r.anomaliesDetectees ?? '',
              r.motifRejet ?? '',
              formatDateForExport(r.submittedAt),
            ])}
            disabled={loadingReports}
          />
          <button className="btn-add" type="button" onClick={refresh} disabled={loadingReports}>
            {loadingReports ? 'Chargement...' : 'Actualiser'}
          </button>
        </div>
      </div>

      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{totals.total}</span>
          <span className="stat-label">Reportings</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{totals.pending}</span>
          <span className="stat-label">En attente</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{totals.realise}</span>
          <span className="stat-label">Validés</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{totals.reporte}</span>
          <span className="stat-label">Refusés</span>
        </div>
        <div className="stat-card montant">
          <span className="stat-value">{formatHours(totals.totalHeures)}</span>
          <span className="stat-label">Total heures</span>
        </div>
      </div>

      <div className="manager-controls">
        <div className="filter-field">
          <label>Contrôleur</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filters.controleur}
            onChange={e => updateFilter('controleur', e.target.value)}
          />
        </div>
        <div className="filter-field">
          <label>Date du</label>
          <input type="date" value={filters.dateFrom} onChange={e => updateFilter('dateFrom', e.target.value)} />
        </div>
        <div className="filter-field">
          <label>Date au</label>
          <input type="date" value={filters.dateTo} onChange={e => updateFilter('dateTo', e.target.value)} />
        </div>
        <label className="filter-toggle">
          <input
            type="checkbox"
            checked={filters.pendingOnly}
            onChange={e => updateFilter('pendingOnly', e.target.checked)}
          />
          En attente uniquement
        </label>
        <button type="button" className="btn-reset-filters" onClick={resetFilters}>Réinitialiser</button>
      </div>

      {/* ─── TAUX DE SOUMISSION DES RAPPORTS PAR CONTRÔLEUR ───────────────
          Section dédiée qui répond à la question métier : « qui soumet ses
          rapports quotidiens et à quel rythme sur la période choisie ? »
          Utilise les filtres date (dateFrom / dateTo) au-dessus ; à défaut,
          période par défaut = 30 derniers jours.

          UX : la section est REPLIÉE par défaut (le manager ne la voit qu'un
          clic sur le header) pour ne pas surcharger la vue principale
          (validation des rapports). Le taux global reste visible dans le
          header du panneau replié, pour offrir un indicateur macro sans
          exiger d'action. */}
      <section
        className="submission-rate-section"
        style={{
          marginBottom: 24,
          border: '1px solid #e2e8f0',
          borderRadius: 10,
          overflow: 'hidden',
        }}
      >
        <button
          type="button"
          onClick={() => setRatesExpanded(v => !v)}
          aria-expanded={ratesExpanded}
          aria-controls="submission-rate-body"
          style={{
            width: '100%',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexWrap: 'wrap',
            gap: 12,
            padding: '12px 16px',
            background: '#f8fafc',
            border: 'none',
            borderBottom: ratesExpanded ? '1px solid #e2e8f0' : 'none',
            cursor: 'pointer',
            textAlign: 'left',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <span
              aria-hidden="true"
              style={{
                display: 'inline-block',
                width: 14,
                fontSize: 12,
                color: '#475569',
                transform: ratesExpanded ? 'rotate(90deg)' : 'none',
                transition: 'transform 0.15s ease',
              }}
            >
              ▶
            </span>
            <h3 style={{ margin: 0, fontSize: 15 }}>
              Taux de soumission des rapports quotidiens
            </h3>
            <span
              style={{
                fontSize: 12,
                fontWeight: 700,
                padding: '2px 8px',
                borderRadius: 12,
                background: '#0f172a',
                color: '#fff',
              }}
            >
              {submissionStats.globalTaux}%
            </span>
          </div>
          <span style={{ fontSize: 12, color: '#64748b' }}>
            {submissionStats.from.toLocaleDateString('fr-FR')}
            {' → '}
            {submissionStats.to.toLocaleDateString('fr-FR')}
            {' — '}
            {submissionStats.expectedDays} j. ouvré(s)
            {controleurs.length > 0 && ` — ${controleurs.length} contrôleur(s)`}
            {' — '}
            <span style={{ textDecoration: 'underline' }}>
              {ratesExpanded ? 'Masquer le détail' : 'Afficher le détail'}
            </span>
          </span>
        </button>

        {ratesExpanded && (
          <div id="submission-rate-body" style={{ padding: 14 }}>
            {/* Bandeau taux global — vue macro instantanée */}
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 16,
                padding: '10px 14px',
                background: '#fff',
                border: '1px solid #e2e8f0',
                borderRadius: 8,
                marginBottom: 12,
              }}
            >
              <div style={{ minWidth: 180 }}>
                <div style={{ fontSize: 11, color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.5 }}>
                  Taux global
                </div>
                <div style={{ fontSize: 22, fontWeight: 700, color: '#0f172a' }}>
                  {submissionStats.globalTaux}%
                </div>
                <div style={{ fontSize: 11, color: '#64748b' }}>
                  {submissionStats.globalCovered} soumission(s) / {submissionStats.globalExpected} attendue(s)
                </div>
              </div>
              <div style={{ flex: 1 }}>
                <ProgressBar
                  value={submissionStats.globalTaux}
                  title={`Taux global : ${submissionStats.globalCovered}/${submissionStats.globalExpected}`}
                />
              </div>
            </div>

            {submissionStats.rows.length === 0 ? (
              <p className="loading-text" style={{ fontSize: 13 }}>
                Aucun contrôleur trouvé pour la période sélectionnée.
              </p>
            ) : (
              <div className="table-wrapper" style={{ maxHeight: 400, overflowY: 'auto' }}>
                <table className="manager-table">
                  <thead>
                    <tr>
                      <th style={{ width: 24 }} aria-label="Expand" />
                      <th>Contrôleur</th>
                      <th>Email</th>
                      <th style={{ textAlign: 'center', width: 100 }}>Soumis</th>
                      <th style={{ textAlign: 'center', width: 100 }}>Attendu</th>
                      <th style={{ minWidth: 180 }}>Taux</th>
                    </tr>
                  </thead>
                  <tbody>
                    {submissionStats.rows.map(row => {
                      const isOpen = expandedControleurEmail === row.email
                      const toggle = () =>
                        setExpandedControleurEmail(prev => (prev === row.email ? null : row.email))
                      return (
                        <Fragment key={row.email}>
                          <tr
                            onClick={toggle}
                            style={{ cursor: 'pointer' }}
                            title={isOpen ? 'Masquer le détail' : 'Voir le détail (manquants + absences)'}
                          >
                            <td
                              aria-hidden="true"
                              style={{
                                color: '#475569',
                                fontSize: 11,
                                transform: isOpen ? 'rotate(90deg)' : 'none',
                                transition: 'transform 0.15s ease',
                                textAlign: 'center',
                              }}
                            >
                              ▶
                            </td>
                            <td>
                              <strong>{row.name}</strong>
                              {row.absent > 0 && (
                                <span
                                  style={{
                                    marginLeft: 8,
                                    fontSize: 10,
                                    padding: '2px 6px',
                                    borderRadius: 10,
                                    background: '#e2e8f0',
                                    color: '#475569',
                                    whiteSpace: 'nowrap',
                                  }}
                                  title={`${row.absent} jour(s) d'absence retiré(s) du dénominateur`}
                                >
                                  🏖️ {row.absent}j
                                </span>
                              )}
                            </td>
                            <td style={{ color: '#64748b', fontSize: 12 }}>{row.email}</td>
                            <td style={{ textAlign: 'center' }}>{row.covered}</td>
                            <td style={{ textAlign: 'center' }}>
                              {row.expected}
                              {row.absent > 0 && (
                                <small style={{ display: 'block', fontSize: 10, color: '#94a3b8' }}>
                                  ({row.expectedBrut} − {row.absent} abs.)
                                </small>
                              )}
                            </td>
                            <td>
                              <ProgressBar
                                value={row.taux}
                                title={`${row.covered} rapport(s) soumis sur ${row.expected} attendu(s) — ${row.absent} jour(s) d'absence exclu(s)`}
                              />
                            </td>
                          </tr>
                          {isOpen && (
                            <tr key={`${row.email}-detail`}>
                              <td />
                              <td colSpan={5} style={{ background: '#fafafa', padding: '10px 14px' }}>
                                <div style={{ fontSize: 12, color: '#334155', marginBottom: 6 }}>
                                  Détail des jours de rapport{' '}
                                  <strong>manquants</strong> pour {row.name} sur la
                                  période{' '}
                                  {submissionStats.from.toLocaleDateString('fr-FR')}{' → '}
                                  {submissionStats.to.toLocaleDateString('fr-FR')} :
                                </div>
                                {row.missingDays.length === 0 ? (
                                  <div style={{ fontSize: 12, color: '#059669' }}>
                                    ✓ Aucun jour manquant — 100 % de couverture.
                                  </div>
                                ) : (
                                  <>
                                    <div
                                      style={{
                                        display: 'flex',
                                        flexWrap: 'wrap',
                                        gap: 6,
                                        marginBottom: 6,
                                      }}
                                    >
                                      {row.missingDays.map(d => (
                                        <span
                                          key={d}
                                          style={{
                                            fontSize: 11,
                                            padding: '3px 8px',
                                            borderRadius: 12,
                                            background: '#fee2e2',
                                            color: '#991b1b',
                                            border: '1px solid #fecaca',
                                            whiteSpace: 'nowrap',
                                          }}
                                          title={`Jour ouvré sans rapport soumis (${d})`}
                                        >
                                          {new Date(`${d}T00:00:00`).toLocaleDateString('fr-FR', {
                                            weekday: 'short',
                                            day: '2-digit',
                                            month: '2-digit',
                                            year: 'numeric',
                                          })}
                                        </span>
                                      ))}
                                    </div>
                                    <div style={{ fontSize: 11, color: '#64748b' }}>
                                      {row.missingDays.length} jour(s) ouvré(s) sans
                                      rapport non refusé sur les {row.expected} attendu(s).
                                    </div>
                                  </>
                                )}

                                {/* ─── Jours d'absence (chips grises) ─────
                                    Affiché seulement si le contrôleur a des
                                    absences dans la période. Ces jours ont
                                    déjà été retirés du dénominateur (cf.
                                    `expectedEffective` dans submissionStats),
                                    donc ils n'apparaissent PAS parmi les
                                    "manquants" — on les liste ici uniquement
                                    à titre informatif pour justifier visuellement
                                    l'écart entre `expectedBrut` et `expected`. */}
                                {row.absentDays.length > 0 && (
                                  <div style={{ marginTop: 12, paddingTop: 10, borderTop: '1px dashed #cbd5e1' }}>
                                    <div style={{ fontSize: 12, color: '#334155', marginBottom: 6 }}>
                                      Jours <strong>d'absence</strong> déclarés
                                      pour {row.name} sur cette période
                                      <span style={{ color: '#64748b', fontWeight: 'normal' }}>
                                        {' '}(retirés du dénominateur du taux) :
                                      </span>
                                    </div>
                                    <div
                                      style={{
                                        display: 'flex',
                                        flexWrap: 'wrap',
                                        gap: 6,
                                        marginBottom: 6,
                                      }}
                                    >
                                      {row.absentDays.map(d => (
                                        <span
                                          key={d}
                                          style={{
                                            fontSize: 11,
                                            padding: '3px 8px',
                                            borderRadius: 12,
                                            background: '#e2e8f0',
                                            color: '#475569',
                                            border: '1px solid #cbd5e1',
                                            whiteSpace: 'nowrap',
                                          }}
                                          title={`Jour d'absence (${d})`}
                                        >
                                          🏖️{' '}
                                          {new Date(`${d}T00:00:00`).toLocaleDateString('fr-FR', {
                                            weekday: 'short',
                                            day: '2-digit',
                                            month: '2-digit',
                                            year: 'numeric',
                                          })}
                                        </span>
                                      ))}
                                    </div>
                                    <div style={{ fontSize: 11, color: '#64748b' }}>
                                      {row.absentDays.length} jour(s) ouvré(s) d'absence
                                      sur la période — gérés via <em>Configuration → Gestion des absences</em>.
                                    </div>
                                  </div>
                                )}
                              </td>
                            </tr>
                          )}
                        </Fragment>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}
      </section>

      {loadingReports ? (
        <p className="loading-text">Chargement des rapports...</p>
      ) : groups.length === 0 ? (
        <div className="manager-empty">
          <p>Aucun reporting ne correspond aux critères.</p>
        </div>
      ) : (
        pagedGroups.map(group => (
          <section key={`${group.key.controleurEmail}-${group.key.date}`} className="manager-group">
            <header className="manager-group-header">
              <div className="manager-group-title">
                <strong>{group.key.controleurName}</strong>
                <span>{group.key.controleurEmail} — {formatDate(group.key.date)}</span>
              </div>
              <div className="manager-group-stats">
                <span className="manager-pill">{group.reports.length} entrée(s)</span>
                <span className="manager-pill">{formatHours(group.totalHeures)}</span>
                {group.pending > 0 && <span className="manager-pill manager-pill-pending">{group.pending} en attente</span>}
                {group.realise > 0 && <span className="manager-pill manager-pill-realise">{group.realise} validé(s)</span>}
                {group.reporte > 0 && <span className="manager-pill manager-pill-reporte">{group.reporte} refusé(s)</span>}
              </div>
            </header>
            <table className="manager-table">
              <thead>
                <tr>
                  <th>Soumis</th>
                  <th>Total heures</th>
                  <th>Temps occupé</th>
                  <th>Statut journée</th>
                  <th style={{ minWidth: 240 }}>Synthèse des actions</th>
                  <th>Statut</th>
                </tr>
              </thead>
              <tbody>
                {group.reports.map(r => (
                  <tr
                    key={r.id}
                    className="row-clickable"
                    onClick={() => setSelected(r)}
                    role="button"
                    tabIndex={0}
                    onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setSelected(r) } }}
                  >
                    <td>{formatDateTime(r.submittedAt)}</td>
                    <td>{formatHours(r.totalHeures)}</td>
                    <td>{r.tempsOccupe}%</td>
                    <td>{r.statutJournee}</td>
                    {/* Synthèse compacte : compteur + 2 premières actions
                        tronquées + bouton "Voir plus" qui ouvre le détail.
                        stopPropagation sur le bouton pour éviter le double
                        déclenchement avec le onClick de la ligne (qui ouvre
                        aussi le détail). */}
                    <td onClick={e => e.stopPropagation()} style={{ fontSize: 12 }}>
                      {r.lines.length === 0 ? (
                        <span style={{ color: '#888' }}>Aucune action</span>
                      ) : (
                        <>
                          <div style={{ fontWeight: 600, marginBottom: 4 }}>
                            {r.lines.length} action{r.lines.length > 1 ? 's' : ''}
                          </div>
                          <ul style={{ margin: 0, padding: 0, listStyle: 'none', display: 'flex', flexDirection: 'column', gap: 2 }}>
                            {r.lines.slice(0, 2).map(line => {
                              const objet = (line.objet || '').trim()
                              // Tronque à ~60 caractères pour rester compact
                              const display = objet.length > 60 ? objet.slice(0, 57) + '…' : objet
                              return (
                                <li key={line.id} style={{ color: '#444' }}>
                                  <span style={{ color: '#1d4ed8', fontWeight: 600 }}>[{line.domaine || '—'}]</span>{' '}
                                  {display || <em style={{ color: '#888' }}>(sans objet)</em>}
                                  <span style={{ color: '#888' }}> — {line.duree} min</span>
                                </li>
                              )
                            })}
                            {r.lines.length > 2 && (
                              <li style={{ color: '#888', fontStyle: 'italic' }}>
                                … +{r.lines.length - 2} autre{r.lines.length - 2 > 1 ? 's' : ''}
                              </li>
                            )}
                          </ul>
                          <button
                            type="button"
                            onClick={() => setSelected(r)}
                            style={{
                              marginTop: 6,
                              padding: '2px 8px',
                              fontSize: 11,
                              fontWeight: 600,
                              color: '#1d4ed8',
                              background: 'transparent',
                              border: '1px solid #bfdbfe',
                              borderRadius: 4,
                              cursor: 'pointer',
                            }}
                          >
                            Voir plus →
                          </button>
                        </>
                      )}
                    </td>
                    <td>
                      <span className={`manager-pill ${
                        r.statut === 'Soumis' ? 'manager-pill-pending'
                          : r.statut === 'Valider' ? 'manager-pill-realise'
                            : 'manager-pill-reporte'
                      }`}>{r.statut}</span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </section>
        ))
      )}

      {/* Barre de pagination — paginée au niveau "groupe" (1 carte = 1 couple contrôleur+date) */}
      {!loadingReports && groups.length > 0 && (
        <Pagination
          state={pagination}
          total={groups.length}
          itemLabel="groupes"
        />
      )}

      {selected && (
        <ManagerDetailModal
          report={selected}
          onClose={() => setSelected(null)}
          onValidate={(decision, motif) => handleValidation(selected, decision, motif)}
        />
      )}
    </>
  )
}

/** Props de la modale détail manager. */
interface ManagerDetailModalProps {
  report: ActivityReport
  onClose: () => void
  onValidate: (decision: 'Valider' | 'Refuser', motif?: string) => void
}

/**
 * Modale de détail d'un rapport pour le manager.
 *
 * Sections :
 *   1. Header : nom du contrôleur + bouton fermer
 *   2. Summary grid : date, total h, % occupé, statut journée, anomalies, statut
 *   3. Tableau des actions réalisées (lignes du rapport)
 *   4. Lien rapport externe (si présent)
 *   5. Observations globales (si présentes)
 *   6. Pièces jointes (si présentes)
 *   7. Historique des décisions managers (si non vide)
 *   8. Zone de décision : textarea motif + boutons Valider / Invalider
 *
 * Raccourci : Escape ferme la modale.
 *
 * Logique des boutons :
 *   - Valider : actif sauf si déjà 'Valider'
 *   - Invalider : actif sauf si déjà 'Refuser' OU motif vide
 */
function ManagerDetailModal({ report, onClose, onValidate }: ManagerDetailModalProps) {
  const [motif, setMotif] = useState('')

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <ModalOverlay onClose={onClose}>
      <div
        className="modal manager-modal"
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-labelledby="manager-modal-title"
      >
        <div className="modal-header">
          <h2 id="manager-modal-title">Reporting — {report.controleurName}</h2>
          <button className="modal-close" onClick={onClose} aria-label="Fermer">&times;</button>
        </div>
        <div className="manager-modal-body">
          <dl className="manager-summary-grid">
            <div className="manager-summary-cell">
              <dt>Date</dt><dd>{formatDate(report.date)}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Total heures</dt><dd>{formatHours(report.totalHeures)}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Temps occupé</dt><dd>{report.tempsOccupe}%</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Statut journée</dt><dd>{report.statutJournee}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>nombres d'anomalies détectées</dt><dd>{report.anomaliesDetectees ?? 0}</dd>
            </div>
            <div className="manager-summary-cell">
              <dt>Statut</dt><dd>{report.statut}</dd>
            </div>
          </dl>

          <section>
            <h3 style={{ margin: '0 0 8px', fontSize: 13, color: 'var(--rdcpo-text-strong)' }}>Actions réalisées</h3>
            <table className="manager-table">
              <thead>
                <tr>
                  <th>#</th>
                  <th>Domaine</th>
                  <th>Objet</th>
                  <th>Action</th>
                  <th>Durée</th>
                  <th>Observation</th>
                </tr>
              </thead>
              <tbody>
                {report.lines.map((line, i) => (
                  <tr key={line.id}>
                    <td>{i + 1}</td>
                    <td>{line.domaine || '—'}</td>
                    <td>{line.objet || '—'}</td>
                    <td>{line.action || '—'}</td>
                    <td>{line.duree} min</td>
                    <td>{line.observation || '—'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </section>

          {report.lienRapport && (
            <p style={{ margin: 0, fontSize: 13 }}>
              <strong>Lien rapport : </strong>
              <a href={report.lienRapport} target="_blank" rel="noreferrer">{report.lienRapport}</a>
            </p>
          )}

          {report.observationsGlobales && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: 'var(--rdcpo-text-strong)' }}>Observations globales</h3>
              <p style={{ margin: 0, fontSize: 13, whiteSpace: 'pre-wrap' }}>{report.observationsGlobales}</p>
            </section>
          )}

          {report.attachments && report.attachments.length > 0 && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: 'var(--rdcpo-text-strong)' }}>Pièces jointes</h3>
              <ul className="bulletin-attachments" style={{ listStyle: 'none', padding: 0, margin: 0 }}>
                {report.attachments.map((a, i) => (
                  <li key={i} style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '6px 0' }}>
                    <span aria-hidden="true">📎</span>
                    {a.url ? <a href={a.url} target="_blank" rel="noreferrer">{a.name}</a> : <span>{a.name}</span>}
                  </li>
                ))}
              </ul>
            </section>
          )}

          {/* Décision actuelle — un seul statut + un seul motif à la fois.
              Pas d'historique des décisions précédentes (politique métier :
              chaque nouvelle décision écrase la précédente). */}
          {report.statut !== 'Soumis' && (
            <section>
              <h3 style={{ margin: '0 0 6px', fontSize: 13, color: 'var(--rdcpo-text-strong)' }}>Décision actuelle</h3>
              <ul className="manager-history">
                <li>
                  <div>
                    <span className={`manager-history-decision ${report.statut === 'Valider' ? 'realise' : 'reporte'}`}>
                      {report.statut}
                    </span>
                    {report.updatedAt && (
                      <>
                        {' — '}
                        <span>{formatDateTime(report.updatedAt)}</span>
                      </>
                    )}
                  </div>
                  {report.motifRejet && (
                    <div style={{ marginTop: 4, color: '#444' }}>
                      <em>Motif :</em> {report.motifRejet}
                    </div>
                  )}
                </li>
              </ul>
            </section>
          )}

          <div className="manager-decision-row">
            <label htmlFor="manager-motif" style={{ fontSize: 12, fontWeight: 700, color: '#666' }}>
              Motif (obligatoire pour invalider)
            </label>
            <textarea
              id="manager-motif"
              rows={3}
              placeholder="Saisir le motif si invalidation..."
              value={motif}
              onChange={e => setMotif(e.target.value)}
            />
            <div className="manager-decision-actions">
              <button
                type="button"
                className="btn-validate"
                onClick={() => onValidate('Valider')}
                disabled={report.statut === 'Valider'}
              >
                ✓ Valider (Valider)
              </button>
              <button
                type="button"
                className="btn-invalidate"
                onClick={() => onValidate('Refuser', motif)}
                disabled={report.statut === 'Refuser' || !motif.trim()}
              >
                ✗ Invalider (Refuser)
              </button>
            </div>
          </div>
        </div>
      </div>
    </ModalOverlay>
  )
}
