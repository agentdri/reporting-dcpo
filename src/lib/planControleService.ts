/**
 * ============================================================================
 * SERVICE — PLAN DE CONTRÔLE (PLAN ANNUEL DCPO)
 * ============================================================================
 *
 * Persistance : liste SharePoint DCPO_LISTE_PLAN_CONTROLE.
 *
 * Mapping des colonnes SharePoint :
 *   ┌──────────────────┬────────────────────┬──────────────────────────────┐
 *   │ Colonne SP       │ Champ domaine      │ Notes                        │
 *   ├──────────────────┼────────────────────┼──────────────────────────────┤
 *   │ Title            │ libelle            │ libellé du contrôle          │
 *   │ field_1          │ categorie          │                              │
 *   │ field_2          │ frequence          │ Quotidienne/Hebdo/Mens/Ann.  │
 *   │ field_3 (number) │ annee              │                              │
 *   │ field_5          │ statut             │ statut workflow              │
 *   │ field_6          │ objectif           │                              │
 *   │ field_7          │ objectifChiffre    │ KPI                          │
 *   │ responsable      │ responsable*       │ champ PERSONNE (Claims)      │
 *   │ Created/Modified │ createdAt/updatedAt│ auto SharePoint              │
 *   └──────────────────┴────────────────────┴──────────────────────────────┘
 *
 * Note : la liste SharePoint n'a PAS de colonne pour le planning mensuel
 * (mois planifiés) — ce concept a été retiré du module.
 *
 * Source initiale des données : feuille "PLAN DE CONTROLE 2025" du fichier
 * Excel Tableau_de_bord_KPI_DCPO_2026-plan-de-controle.xlsx.
 * ============================================================================
 */

import { DCPO_LISTE_PLAN_CONTROLEService } from '../generated/services/DCPO_LISTE_PLAN_CONTROLEService'
import type {
  DCPO_LISTE_PLAN_CONTROLERead,
  DCPO_LISTE_PLAN_CONTROLEWrite,
} from '../generated/models/DCPO_LISTE_PLAN_CONTROLEModel'
import { DCPO_EVALUATION_PLAN_CONTROLEService } from '../generated/services/DCPO_EVALUATION_PLAN_CONTROLEService'
import type {
  DCPO_EVALUATION_PLAN_CONTROLERead,
  DCPO_EVALUATION_PLAN_CONTROLEWrite,
} from '../generated/models/DCPO_EVALUATION_PLAN_CONTROLEModel'
import { appendUrl, parseUrlList, getFileNameFromUrl } from './ticketAttachments'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE
 * ────────────────────────────────────────────────────────────────────────── */

/** Fréquence d'exécution d'un contrôle (colonne field_2). */
export type ControleFrequence = 'Quotidienne' | 'Hebdomadaire' | 'Mensuelle' | 'Annuelle'

/** Statut workflow d'un contrôle (colonne field_5). */
export type ControleStatus = 'À planifier' | 'Planifié' | 'En cours' | 'Réalisé' | 'Annulé'

/**
 * Représentation d'un contrôle du plan annuel.
 *
 * Le responsable est un champ Personne côté SharePoint : on conserve à la
 * fois son nom affiché (responsable) et son email (responsableEmail) — ce
 * dernier sert à reconstruire le format Claims lors d'une écriture.
 */
export interface ControleEntry {
  id: string
  libelle: string
  categorie: string
  objectif: string
  objectifChiffre: string
  frequence: ControleFrequence
  responsable: string       // DisplayName (affichage)
  responsableEmail: string  // email (pour écriture Person)
  annee: number
  statut: ControleStatus
  /**
   * Pièces jointes décodées depuis le champ urlPieceJointe (multi-URLs
   * séparées par " | "). Reconstruit à la lecture pour l'affichage UI.
   */
  attachments?: { name: string; url: string }[]
  createdAt: string
  updatedAt: string
}

/** Données pour créer un contrôle. */
export interface CreateControleInput {
  libelle: string
  categorie: string
  objectif: string
  objectifChiffre: string
  frequence: ControleFrequence
  responsableName: string
  responsableEmail: string
  annee: number
  statut: ControleStatus
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — RÉFÉRENTIELS (catégories, fréquences, statuts)
 *
 * Issus du fichier Excel — codés en dur car stables.
 * ────────────────────────────────────────────────────────────────────────── */

export const PLAN_CONTROLE_CATEGORIES: string[] = [
  "Contrôle préventif dans les unités d'exploitation",
  "Surveillance des opérations locales",
  "Renforcement du contrôle des activités du Comex",
  "Sécurisation des revenus de la banque",
  "Renforcement du suivi des opérations remarquables",
  "Renforcement de la surveillance des activités des directions centrales",
  "Renforcement de la surveillance des opérations (trésorerie)",
  "Évaluation contrôle des souscriptions aux produits",
  "Contrôle des opérations monétiques",
  "Contrôle des journées comptables",
  "Évaluation des contrôles des caisses",
  "Contrôle des existants",
  "Surveillance des engagements",
  "Réconciliation des comptes",
  "Révision des comptes",
  "Évaluation des activités externalisées",
  "Autre",
]

export const FREQUENCE_OPTIONS: ControleFrequence[] = [
  'Quotidienne',
  'Hebdomadaire',
  'Mensuelle',
  'Annuelle',
]

export const STATUS_OPTIONS: ControleStatus[] = [
  'À planifier',
  'Planifié',
  'En cours',
  'Réalisé',
  'Annulé',
]


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — MAPPING SHAREPOINT ↔ DOMAINE
 * ────────────────────────────────────────────────────────────────────────── */

/** Format SharePoint Claims pour un champ Personne. */
function toClaims(email: string): string {
  return `i:0#.f|membership|${email}`
}

/** Construit un ControleEntry depuis un item SharePoint. */
function fromItem(item: DCPO_LISTE_PLAN_CONTROLERead): ControleEntry {
  // Normalisation défensive de la fréquence / statut : si la valeur SP n'est
  // pas dans la liste connue, on retombe sur une valeur par défaut sûre.
  const freq = item.field_2 as ControleFrequence
  const statut = item.field_5 as ControleStatus
  // Pièces jointes : parsing du champ urlPieceJointe (multi-URLs " | ")
  const attachmentUrls = parseUrlList(item.urlPieceJointe)
  const attachments = attachmentUrls.map(url => ({
    name: getFileNameFromUrl(url),
    url,
  }))
  return {
    id: String(item.ID),
    libelle: item.Title ?? '',
    categorie: item.field_1 ?? '',
    frequence: FREQUENCE_OPTIONS.includes(freq) ? freq : 'Mensuelle',
    annee: item.field_3 ?? new Date().getFullYear(),
    statut: STATUS_OPTIONS.includes(statut) ? statut : 'À planifier',
    objectif: item.field_6 ?? '',
    objectifChiffre: item.field_7 ?? '',
    responsable: item.responsable?.DisplayName ?? '',
    responsableEmail: item.responsable?.Email ?? '',
    attachments,
    createdAt: item.Created ?? '',
    updatedAt: item.Modified ?? '',
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — API PUBLIQUE
 * ────────────────────────────────────────────────────────────────────────── */

/** Liste tous les contrôles (du plus récent au plus ancien). */
export async function listControles(): Promise<ControleEntry[]> {
  try {
    const res = await DCPO_LISTE_PLAN_CONTROLEService.getAll({ orderBy: ['Created desc'] })
    if (!res.data) return []
    return res.data.map(fromItem)
  } catch (err) {
    console.error('listControles error', err)
    return []
  }
}

/** Récupère un contrôle par ID. */
export async function getControle(id: string): Promise<ControleEntry | undefined> {
  try {
    const res = await DCPO_LISTE_PLAN_CONTROLEService.get(id)
    if (!res.data) return undefined
    return fromItem(res.data)
  } catch (err) {
    console.error('getControle error', err)
    return undefined
  }
}

/**
 * Crée un contrôle dans SharePoint.
 *
 * Validation minimale (l'UI fait le détail) :
 *   - libelle, categorie obligatoires
 *   - annee valide
 */
export async function createControle(input: CreateControleInput): Promise<ControleEntry> {
  if (!input.libelle.trim()) throw new Error('Libellé obligatoire.')
  if (!input.categorie.trim()) throw new Error('Catégorie obligatoire.')
  if (!input.annee || input.annee <= 0) throw new Error('Année invalide.')

  const payload: Record<string, unknown> = {
    Title: input.libelle.trim(),
    field_1: input.categorie,
    field_2: input.frequence,
    field_3: input.annee,
    field_5: input.statut,
    field_6: input.objectif.trim(),
    field_7: input.objectifChiffre.trim(),
  }
  // Champ Personne : on n'écrit le responsable que si un email est fourni.
  if (input.responsableEmail) {
    payload.responsable = {
      '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
      Claims: toClaims(input.responsableEmail),
    }
  }

  const res = await DCPO_LISTE_PLAN_CONTROLEService.create(
    payload as Omit<DCPO_LISTE_PLAN_CONTROLEWrite, 'ID'>,
  )
  if (!res.success || !res.data) {
    throw new Error(res.error?.message ?? 'Échec de la création du contrôle.')
  }
  return fromItem(res.data)
}

/**
 * Met à jour UNIQUEMENT le responsable d'un contrôle (affectation rapide).
 *
 * Cas d'usage : workflow "Affecter" depuis le tableau — un manager change
 * la personne responsable sans avoir à rouvrir le formulaire d'édition
 * complet.
 *
 * @param id - ID SharePoint du contrôle
 * @param email - Email de la nouvelle personne responsable (format
 *                'jdoe@afrilandfirstbank.com' — le format Claims est ajouté ici)
 * @returns Le contrôle rechargé après update, ou undefined en cas d'erreur.
 */
export async function updateControleResponsable(
  id: string,
  email: string,
): Promise<ControleEntry | undefined> {
  if (!email) return undefined
  try {
    await DCPO_LISTE_PLAN_CONTROLEService.update(
      id,
      {
        responsable: {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(email),
        },
      } as never,
    )
  } catch (err) {
    console.error('updateControleResponsable error', err)
    return undefined
  }
  return getControle(id)
}

/**
 * Met à jour un contrôle existant (champs métier — pas les pièces jointes).
 *
 * Pour ajouter une PJ, utiliser `appendPlanControleAttachmentUrls`.
 *
 * @returns Le contrôle rechargé après update, ou undefined en cas d'erreur.
 */
export async function updateControle(id: string, input: CreateControleInput): Promise<ControleEntry | undefined> {
  const payload: Record<string, unknown> = {
    Title: input.libelle.trim(),
    field_1: input.categorie,
    field_2: input.frequence,
    field_3: input.annee,
    field_5: input.statut,
    field_6: input.objectif.trim(),
    field_7: input.objectifChiffre.trim(),
  }
  if (input.responsableEmail) {
    payload.responsable = {
      '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
      Claims: toClaims(input.responsableEmail),
    }
  }
  try {
    await DCPO_LISTE_PLAN_CONTROLEService.update(
      id,
      payload as Partial<Omit<DCPO_LISTE_PLAN_CONTROLEWrite, 'ID'>>,
    )
  } catch (err) {
    console.error('updateControle error', err)
    return undefined
  }
  return getControle(id)
}

/**
 * Ajoute (concatène) une ou plusieurs URLs de PJ au champ `urlPieceJointe`
 * du contrôle, sans écraser les existantes.
 *
 * Même mécanique que appendPacAttachmentUrls : lecture, concat via appendUrl
 * (dédoublonne + sépare par " | "), réécriture.
 */
export async function appendPlanControleAttachmentUrls(controleId: string, newUrls: string[]): Promise<void> {
  const cleanUrls = newUrls.filter(u => !!u && u.trim().length > 0)
  if (cleanUrls.length === 0) return

  let existing = ''
  try {
    const res = await DCPO_LISTE_PLAN_CONTROLEService.get(controleId)
    existing = res.data?.urlPieceJointe ?? ''
  } catch (err) {
    console.error('appendPlanControleAttachmentUrls: échec lecture item', err)
  }

  const concatenated = cleanUrls.reduce(
    (acc, url) => appendUrl(acc, url),
    existing,
  )

  try {
    await DCPO_LISTE_PLAN_CONTROLEService.update(controleId, {
      urlPieceJointe: concatenated,
    } as Partial<Omit<DCPO_LISTE_PLAN_CONTROLEWrite, 'ID'>>)
  } catch (err) {
    console.error('appendPlanControleAttachmentUrls: échec update item', err)
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — ÉVALUATIONS DU PLAN DE CONTRÔLE
 *
 * Liste SharePoint DCPO_EVALUATION_PLAN_CONTROLE — une évaluation
 * correspond à l'exécution PÉRIODIQUE d'un contrôle sur une période donnée
 * (jour / semaine / mois / année selon la fréquence du contrôle parent).
 *
 * Mapping des colonnes :
 *   ┌──────────────────┬────────────────────┬──────────────────────────────┐
 *   │ Title            │ titre (libellé évaluation) │ ex: "Évaluation 2026-05" │
 *   │ observations     │ observations       │ texte libre                  │
 *   │ plan_controle_id │ planControleId     │ FK vers le contrôle parent   │
 *   │ periode          │ periode            │ format dépend de la fréq.    │
 *   └──────────────────┴────────────────────┴──────────────────────────────┘
 *
 * Format de `periode` stocké en SP, en fonction de la fréquence du contrôle :
 *   - Quotidienne  → "2026-05-15"  (input type=date)
 *   - Hebdomadaire → "2026-W22"    (input type=week, ISO week)
 *   - Mensuelle    → "2026-05"     (input type=month)
 *   - Annuelle     → "2026"        (input number / texte)
 *
 * On stocke la valeur brute du champ HTML pour un round-trip lisible côté
 * SharePoint (un manager qui ouvre l'item peut comprendre la période sans
 * formatage spécial).
 * ────────────────────────────────────────────────────────────────────────── */

/** Évaluation d'un contrôle (vue front). */
export interface ControleEvaluation {
  id: string
  titre: string
  planControleId: number
  periode: string          // format dépend de la fréquence du contrôle parent
  observations: string
  /**
   * Pièces jointes décodées depuis le champ urlPieceJointe (multi-URLs
   * séparées par " | ", idem anomalies / contrôles). Reconstruit à la
   * lecture pour faciliter l'affichage côté UI.
   */
  attachments?: { name: string; url: string }[]
  createdAt: string
  updatedAt: string
}

/** Input pour créer une évaluation. */
export interface CreateEvaluationInput {
  planControleId: number
  periode: string
  observations: string
  titre?: string  // facultatif — défaut auto-généré depuis la période
}

/** Mapping item SP → ControleEvaluation. */
function fromEvaluationItem(item: DCPO_EVALUATION_PLAN_CONTROLERead): ControleEvaluation {
  // Pièces jointes : parsing du champ urlPieceJointe (multi-URLs " | ")
  const attachmentUrls = parseUrlList(item.urlPieceJointe)
  const attachments = attachmentUrls.map(url => ({
    name: getFileNameFromUrl(url),
    url,
  }))
  return {
    id: String(item.ID),
    titre: item.Title ?? '',
    planControleId: item.plan_controle_id ?? 0,
    periode: item.periode ?? '',
    observations: item.observations ?? '',
    attachments,
    createdAt: item.Created ?? '',
    updatedAt: item.Modified ?? '',
  }
}

/**
 * Liste les évaluations d'un contrôle donné (les plus récentes d'abord).
 *
 * Filtre serveur sur `plan_controle_id` pour limiter la charge réseau.
 */
export async function listEvaluationsForControle(controleId: string | number): Promise<ControleEvaluation[]> {
  const idNum = Number(controleId)
  if (!Number.isFinite(idNum)) return []
  try {
    const res = await DCPO_EVALUATION_PLAN_CONTROLEService.getAll({
      filter: `plan_controle_id eq ${idNum}`,
      orderBy: ['Created desc'],
    })
    if (!res.data) return []
    return res.data.map(fromEvaluationItem)
  } catch (err) {
    console.error('listEvaluationsForControle error', err)
    return []
  }
}

/**
 * Liste TOUTES les évaluations (utilisé pour le calcul du taux d'évolution
 * en bloc côté Plan de Contrôle — évite N requêtes serveur).
 *
 * Filtre serveur optionnel sur l'année pour limiter la charge si la liste
 * grossit. L'année est matchée contre la chaîne `periode` (qui peut être
 * 'YYYY', 'YYYY-MM', 'YYYY-MM-DD' ou 'YYYY-Www' — tous commencent par
 * l'année).
 */
export async function listAllEvaluations(year?: number): Promise<ControleEvaluation[]> {
  try {
    const options: { filter?: string; orderBy: string[] } = { orderBy: ['Created desc'] }
    if (year && Number.isFinite(year)) {
      options.filter = `startswith(periode, '${year}')`
    }
    const res = await DCPO_EVALUATION_PLAN_CONTROLEService.getAll(options)
    if (!res.data) return []
    return res.data.map(fromEvaluationItem)
  } catch (err) {
    console.error('listAllEvaluations error', err)
    return []
  }
}

/**
 * Nombre d'évaluations attendues sur une année pour une fréquence donnée.
 * Sert de dénominateur pour calculer le taux d'évolution.
 */
export function getExpectedEvaluationsPerYear(frequence: ControleFrequence): number {
  switch (frequence) {
    case 'Quotidienne': return 365
    case 'Hebdomadaire': return 52
    case 'Mensuelle': return 12
    case 'Annuelle': return 1
  }
}

/**
 * Calcule le taux d'évolution (%) d'un contrôle = nb évaluations faites
 * divisé par nb attendu pour la fréquence, capé à 100.
 */
export function computeEvaluationProgress(frequence: ControleFrequence, count: number): number {
  const expected = getExpectedEvaluationsPerYear(frequence)
  if (expected <= 0) return 0
  const pct = (count / expected) * 100
  return Math.max(0, Math.min(100, Math.round(pct)))
}

/**
 * Crée une évaluation pour un contrôle.
 *
 * Le `Title` SP est obligatoire — si pas fourni, on génère un libellé par
 * défaut "Évaluation <periode>" pour qu'un manager voie tout de suite à quoi
 * correspond l'item dans l'UI SharePoint.
 */
export async function createEvaluation(input: CreateEvaluationInput): Promise<ControleEvaluation> {
  if (!input.planControleId) throw new Error('Contrôle parent manquant.')
  if (!input.periode.trim()) throw new Error('La période est obligatoire.')

  const title = (input.titre ?? '').trim() || `Évaluation ${input.periode}`

  const payload: Record<string, unknown> = {
    Title: title,
    plan_controle_id: input.planControleId,
    periode: input.periode.trim(),
    observations: input.observations.trim(),
  }

  const res = await DCPO_EVALUATION_PLAN_CONTROLEService.create(
    payload as Omit<DCPO_EVALUATION_PLAN_CONTROLEWrite, 'ID'>,
  )
  if (!res.success || !res.data) {
    throw new Error(res.error?.message ?? 'Échec de la création de l\'évaluation.')
  }
  return fromEvaluationItem(res.data)
}

/**
 * Ajoute (concatène) une ou plusieurs URLs de PJ au champ `urlPieceJointe`
 * d'une évaluation, sans écraser les existantes.
 *
 * Pattern identique aux autres lists : lecture → concat via appendUrl
 * (dédoublonne + sépare par " | ") → réécriture.
 */
export async function appendEvaluationAttachmentUrls(evaluationId: string, newUrls: string[]): Promise<void> {
  const cleanUrls = newUrls.filter(u => !!u && u.trim().length > 0)
  if (cleanUrls.length === 0) return

  let existing = ''
  try {
    const res = await DCPO_EVALUATION_PLAN_CONTROLEService.get(evaluationId)
    existing = res.data?.urlPieceJointe ?? ''
  } catch (err) {
    console.error('appendEvaluationAttachmentUrls: échec lecture item', err)
  }

  const concatenated = cleanUrls.reduce(
    (acc, url) => appendUrl(acc, url),
    existing,
  )

  try {
    await DCPO_EVALUATION_PLAN_CONTROLEService.update(evaluationId, {
      urlPieceJointe: concatenated,
    } as Partial<Omit<DCPO_EVALUATION_PLAN_CONTROLEWrite, 'ID'>>)
  } catch (err) {
    console.error('appendEvaluationAttachmentUrls: échec update item', err)
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 — HELPERS UI
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Type d'input HTML adapté à la fréquence d'un contrôle.
 * Permet à la modale d'évaluation de proposer le bon picker.
 */
export function getPeriodeInputType(frequence: ControleFrequence): 'date' | 'week' | 'month' | 'number' {
  switch (frequence) {
    case 'Quotidienne': return 'date'
    case 'Hebdomadaire': return 'week'
    case 'Mensuelle': return 'month'
    case 'Annuelle': return 'number'
  }
}

/**
 * Formatte un libellé de période lisible selon la fréquence.
 *   - Quotidienne  : "Jour 15/05/2026"
 *   - Hebdomadaire : "Semaine 22 - 2026"
 *   - Mensuelle    : "Mai 2026"
 *   - Annuelle     : "Année 2026"
 *
 * Fallback : renvoie la valeur brute si le parsing échoue.
 */
export function formatPeriodeLabel(periode: string, frequence?: ControleFrequence): string {
  if (!periode) return '—'
  try {
    if (frequence === 'Quotidienne' || /^\d{4}-\d{2}-\d{2}$/.test(periode)) {
      const d = new Date(periode)
      if (!Number.isNaN(d.getTime())) return `Jour ${d.toLocaleDateString('fr-FR')}`
    }
    if (frequence === 'Hebdomadaire' || /^\d{4}-W\d{2}$/i.test(periode)) {
      const m = periode.match(/^(\d{4})-W(\d{2})$/i)
      if (m) return `Semaine ${m[2]} - ${m[1]}`
    }
    if (frequence === 'Mensuelle' || /^\d{4}-\d{2}$/.test(periode)) {
      const m = periode.match(/^(\d{4})-(\d{2})$/)
      if (m) {
        const months = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre']
        const monthIdx = parseInt(m[2], 10) - 1
        if (monthIdx >= 0 && monthIdx < 12) return `${months[monthIdx]} ${m[1]}`
      }
    }
    if (frequence === 'Annuelle' || /^\d{4}$/.test(periode)) {
      return `Année ${periode}`
    }
  } catch {
    /* fallback */
  }
  return periode
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 7 — HELPERS UI (statut)
 * ────────────────────────────────────────────────────────────────────────── */

/** Mappe un statut vers une classe CSS (pastilles manager-pill). */
export function getControleStatusClass(statut: ControleStatus): string {
  switch (statut) {
    case 'À planifier': return 'manager-pill-pending'
    case 'Planifié': return 'manager-pill'
    case 'En cours': return 'manager-pill-pending'
    case 'Réalisé': return 'manager-pill-realise'
    case 'Annulé': return 'manager-pill-reporte'
    default: return 'manager-pill'
  }
}
