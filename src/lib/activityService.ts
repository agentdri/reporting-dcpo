/**
 * ============================================================================
 * SERVICE DE GESTION DES RAPPORTS QUOTIDIENS (Module 3)
 * ============================================================================
 *
 * Persistance principale : liste SharePoint DCPO_ACTIVICTE_CONTROLLER
 *
 * Mapping des champs SharePoint :
 *   ┌─────────────────────────────────┬───────────────────────────────────┐
 *   │ Colonne SharePoint              │ Champ TypeScript                  │
 *   ├─────────────────────────────────┼───────────────────────────────────┤
 *   │ ID (auto)                       │ id                                │
 *   │ Title                           │ "Rapport <date> - <contrôleur>"   │
 *   │ controlleur (Person)            │ controleurEmail + controleurName  │
 *   │ Date                            │ date (YYYY-MM-DD)                 │
 *   │ actionDeLaJournee (multi-line)  │ lines[] (sérialisées, voir plus bas)│
 *   │ nombreAnomalieDetectee          │ anomaliesDetectees                │
 *   │ observationsGlobales            │ observationsGlobales              │
 *   │ Created (auto)                  │ submittedAt                       │
 *   │ Modified (auto)                 │ updatedAt                         │
 *   └─────────────────────────────────┴───────────────────────────────────┘
 *
 * Sérialisation des lignes dans actionDeLaJournee :
 *   La liste SharePoint n'a qu'UN champ multiligne pour TOUTES les actions
 *   de la journée. On y stocke un format hybride :
 *     1. Section humaine lisible (parcourable directement dans SharePoint)
 *     2. Bloc JSON encadré par <!--LINES_JSON ... --> pour le round-trip
 *        machine fiable (parser robuste, pas de regex fragile)
 *
 * Workflow validation :
 *   La liste SharePoint actuelle ne dispose PAS des colonnes `statut` et
 *   `historiqueValidations` → ces données sont stockées en localStorage,
 *   indexées par l'ID SharePoint du rapport. Le statut par défaut au
 *   chargement est 'Soumis'.
 *
 *   Pour activer le workflow complet de validation côté SharePoint, ajouter :
 *     - statut          : Choix (Soumis / Réalisé / Reporté)
 *     - historiqueValidations : Plusieurs lignes de texte (JSON)
 *
 * Brouillons :
 *   Restent en localStorage (saisie en cours, pas de pollution serveur).
 * ============================================================================
 */

import { DCPO_ACTIVICTE_CONTROLLERService } from '../generated/services/DCPO_ACTIVICTE_CONTROLLERService'
import type {
  DCPO_ACTIVICTE_CONTROLLERRead,
  DCPO_ACTIVICTE_CONTROLLERWrite,
} from '../generated/models/DCPO_ACTIVICTE_CONTROLLERModel'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE (inchangés côté API publique)
 * ────────────────────────────────────────────────────────────────────────── */

/** Statut de workflow d'un rapport (suivi via localStorage). */
export type ActivityStatus = 'Soumis' | 'Réalisé' | 'Reporté'

/** Une ligne d'activité (action menée par le contrôleur). */
export interface ActivityLine {
  id: string
  domaine: string
  objet: string
  action: string
  duree: number
  observation?: string
}

/** Trace d'une décision manager (validation/invalidation). */
export interface ActivityValidationNote {
  date: string
  manager: string
  managerEmail?: string
  decision: 'Réalisé' | 'Reporté'
  motif?: string
}

/** Représentation enrichie d'un rapport pour le front (combine SP + métas locales). */
export interface ActivityReport {
  id: string
  controleurEmail: string
  controleurName: string
  date: string
  lines: ActivityLine[]
  totalHeures: number
  tempsOccupe: number
  statutJournee: 'Calme' | 'Normal' | 'Chargé' | 'Très chargé'
  statut: ActivityStatus
  anomaliesDetectees?: number
  lienRapport?: string
  observationsGlobales?: string
  attachments?: { name: string; url: string }[]
  validationHistory: ActivityValidationNote[]
  submittedAt: string
  updatedAt: string
}

/** Brouillon en cours de saisie (localStorage uniquement). */
export interface ActivityDraft {
  id?: string
  controleurEmail: string
  date: string
  lines: ActivityLine[]
  anomaliesDetectees?: number
  lienRapport?: string
  observationsGlobales?: string
  updatedAt: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — LOCALSTORAGE (brouillons + métadonnées de validation)
 * ────────────────────────────────────────────────────────────────────────── */

/** Préfixe pour les brouillons de saisie (par contrôleur). */
const STORAGE_KEY_DRAFT_PREFIX = 'reportingDCPO.activityDraft.v1.'

/**
 * Préfixe pour les métadonnées de validation (statut + historique manager).
 * Indexé par l'ID SharePoint du rapport.
 * Cette indirection disparaîtra quand les colonnes statut/historiqueValidations
 * seront ajoutées à la liste SharePoint.
 */
const STORAGE_KEY_VALIDATION_PREFIX = 'reportingDCPO.activityValidation.v1.'

const READ_RAW = (key: string): string | null => {
  try { return window.localStorage.getItem(key) } catch { return null }
}
const WRITE_RAW = (key: string, value: string): void => {
  try { window.localStorage.setItem(key, value) } catch { /* ignore */ }
}
const REMOVE_RAW = (key: string): void => {
  try { window.localStorage.removeItem(key) } catch { /* ignore */ }
}

/** Génère un identifiant local unique (utilisé pour les ID des lignes). */
function uid(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}

/** Forme du blob de validation stocké en localStorage. */
interface LocalValidation {
  statut: ActivityStatus
  history: ActivityValidationNote[]
}

/** Lit la validation stockée localement pour un rapport. */
function readValidation(reportId: string): LocalValidation | undefined {
  const raw = READ_RAW(STORAGE_KEY_VALIDATION_PREFIX + reportId)
  if (!raw) return undefined
  try {
    return JSON.parse(raw) as LocalValidation
  } catch {
    return undefined
  }
}

/** Écrase les méta-données de validation pour un rapport. */
function writeValidation(reportId: string, validation: LocalValidation): void {
  WRITE_RAW(STORAGE_KEY_VALIDATION_PREFIX + reportId, JSON.stringify(validation))
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — SÉRIALISATION DES LIGNES DANS actionDeLaJournee
 *
 * Format hybride :
 *   1. Section humaine lisible (le contrôleur peut la relire dans SharePoint)
 *   2. Bloc JSON marqué pour parsing robuste (round-trip sans perte)
 *
 * Exemple produit :
 *   1) [Contrôle] Visite agence Bessengue — 60 min
 *      ▸ Vérification des écritures du jour
 *      Observation : RAS
 *
 *   2) [Reporting] Rapport hebdomadaire — 90 min
 *      ▸ Compilation des chiffres
 *
 *   <!--LINES_JSON {"v":1,"lines":[{...},{...}]}-->
 * ────────────────────────────────────────────────────────────────────────── */

const JSON_BLOCK_RE = /<!--LINES_JSON\s*([\s\S]*?)\s*-->/

/**
 * Sérialise un tableau de ActivityLine vers le texte multiligne SharePoint.
 *
 * Stratégie :
 *   - Génère d'abord la version humaine (numérotée, structurée)
 *   - Append un bloc JSON minifié encadré par marqueurs HTML-comment
 *     (lisible mais ne pollue pas trop visuellement)
 */
export function serializeLines(lines: ActivityLine[]): string {
  if (!lines || lines.length === 0) return ''

  // 1. Section humaine
  const human = lines.map((l, i) => {
    const num = i + 1
    const domaine = (l.domaine || '—').trim()
    const objet = (l.objet || '—').trim()
    const duree = Number.isFinite(l.duree) ? l.duree : 0
    const lines: string[] = []
    lines.push(`${num}) [${domaine}] ${objet} — ${duree} min`)
    if (l.action?.trim()) lines.push(`   ▸ ${l.action.trim()}`)
    if (l.observation?.trim()) lines.push(`   Observation : ${l.observation.trim()}`)
    return lines.join('\n')
  }).join('\n\n')

  // 2. Bloc JSON pour round-trip
  const json = JSON.stringify({ v: 1, lines })
  return `${human}\n\n<!--LINES_JSON ${json}-->`
}

/**
 * Parse le contenu de actionDeLaJournee pour récupérer les lignes.
 *
 * Stratégie :
 *   1. Cherche d'abord le bloc <!--LINES_JSON ... --> (round-trip parfait)
 *   2. Si absent ou JSON invalide : tentative de parsing heuristique du
 *      texte humain (pour les rapports créés manuellement dans SharePoint
 *      sans passer par l'app)
 *   3. Si tout échoue : retourne []
 */
export function parseLines(raw: string | null | undefined): ActivityLine[] {
  if (!raw) return []

  // 1. Tentative bloc JSON (chemin nominal)
  const match = raw.match(JSON_BLOCK_RE)
  if (match && match[1]) {
    try {
      const parsed = JSON.parse(match[1]) as { v?: number; lines?: ActivityLine[] }
      if (parsed && Array.isArray(parsed.lines)) {
        return parsed.lines.map(l => ({
          id: l.id ?? uid(),
          domaine: l.domaine ?? '',
          objet: l.objet ?? '',
          action: l.action ?? '',
          duree: Number(l.duree) || 0,
          observation: l.observation,
        }))
      }
    } catch (err) {
      console.warn('parseLines: bloc JSON invalide, fallback texte', err)
    }
  }

  // 2. Fallback heuristique : tente d'extraire les lignes du texte humain
  // Format attendu : "N) [domaine] objet — duree min" suivi de ▸ action et Observation:
  return parseLinesHeuristic(raw)
}

/**
 * Parsing heuristique du texte humain quand le bloc JSON est absent.
 * Utilisé pour les rapports créés directement en SharePoint sans passer
 * par l'app. Best-effort, pas de garantie de fidélité.
 */
function parseLinesHeuristic(raw: string): ActivityLine[] {
  // Retire le bloc JSON s'il existe (avec contenu corrompu)
  const text = raw.replace(JSON_BLOCK_RE, '').trim()
  if (!text) return []

  const out: ActivityLine[] = []
  // Sépare en blocs sur les double-newlines
  const blocks = text.split(/\n\s*\n/)
  for (const block of blocks) {
    const lines = block.split(/\r?\n/).map(l => l.trim()).filter(Boolean)
    if (lines.length === 0) continue
    // Parse la ligne d'en-tête : "N) [domaine] objet — duree min"
    const headerMatch = lines[0].match(/^\d+\)\s*\[([^\]]+)\]\s*(.+?)\s*[—-]\s*(\d+)\s*min/i)
    if (!headerMatch) continue
    const action = lines.find(l => l.startsWith('▸'))?.replace(/^▸\s*/, '') ?? ''
    const observation = lines.find(l => l.toLowerCase().startsWith('observation'))?.replace(/^observation\s*[:：]\s*/i, '') ?? ''
    out.push({
      id: uid(),
      domaine: headerMatch[1].trim(),
      objet: headerMatch[2].trim(),
      duree: parseInt(headerMatch[3], 10) || 0,
      action,
      observation: observation || undefined,
    })
  }
  return out
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — MAPPING ITEM SHAREPOINT → ACTIVITYREPORT
 * ────────────────────────────────────────────────────────────────────────── */

/** Construit un ActivityReport complet à partir d'un item SharePoint. */
function reportFromItem(item: DCPO_ACTIVICTE_CONTROLLERRead): ActivityReport {
  const lines = parseLines(item.actionDeLaJournee)
  const totals = computeTotals(lines)
  const dateStr = item.Date ? item.Date.split('T')[0] : ''
  const validation = readValidation(String(item.ID))

  return {
    id: String(item.ID),
    controleurEmail: item.controlleur?.Email ?? '',
    controleurName: item.controlleur?.DisplayName ?? '',
    date: dateStr,
    lines,
    totalHeures: totals.totalHeures,
    tempsOccupe: totals.tempsOccupe,
    statutJournee: totals.statutJournee,
    statut: validation?.statut ?? 'Soumis',
    anomaliesDetectees: item.nombreAnomalieDetectee,
    lienRapport: undefined,
    observationsGlobales: item.observationsGlobales,
    attachments: [],
    validationHistory: validation?.history ?? [],
    submittedAt: item.Created ?? '',
    updatedAt: item.Modified ?? '',
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — API PUBLIQUE : LECTURE
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Liste tous les rapports depuis SharePoint, triés du plus récent au plus ancien.
 * Robuste : retourne [] en cas d'erreur (logs en console).
 */
export async function listReports(): Promise<ActivityReport[]> {
  try {
    const result = await DCPO_ACTIVICTE_CONTROLLERService.getAll({
      orderBy: ['Date desc'],
    })
    if (!result.data) return []
    return result.data.map(reportFromItem)
  } catch (err) {
    console.error('listReports error', err)
    return []
  }
}

/** Récupère un rapport par ID SharePoint. */
export async function getReport(id: string): Promise<ActivityReport | undefined> {
  try {
    const result = await DCPO_ACTIVICTE_CONTROLLERService.get(id)
    if (!result.data) return undefined
    return reportFromItem(result.data)
  } catch (err) {
    console.error('getReport error', err)
    return undefined
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 — CALCULS DÉRIVÉS
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Calcule les totaux dérivés à partir des lignes :
 *   - totalHeures   : somme des durées en heures (arrondi 2 décimales)
 *   - tempsOccupe   : pourcentage d'une journée 8h occupée (cap à 100%)
 *   - statutJournee : étiquette qualitative basée sur totalHeures
 *
 * Recalculé à la lecture (pas stocké en SharePoint) → toujours cohérent
 * avec le contenu réel des lignes.
 */
export function computeTotals(lines: ActivityLine[]): { totalHeures: number; tempsOccupe: number; statutJournee: ActivityReport['statutJournee'] } {
  const totalMinutes = lines.reduce((s, l) => s + (Number.isFinite(l.duree) ? Math.max(0, l.duree) : 0), 0)
  const totalHeures = Math.round((totalMinutes / 60) * 100) / 100
  const tempsOccupe = Math.min(100, Math.round((totalMinutes / (8 * 60)) * 100))

  let statutJournee: ActivityReport['statutJournee'] = 'Normal'
  if (totalHeures < 4) statutJournee = 'Calme'
  else if (totalHeures < 7) statutJournee = 'Normal'
  else if (totalHeures < 9) statutJournee = 'Chargé'
  else statutJournee = 'Très chargé'

  return { totalHeures, tempsOccupe, statutJournee }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 7 — CRÉATION D'UN RAPPORT (SHAREPOINT)
 * ────────────────────────────────────────────────────────────────────────── */

/** Données minimales pour créer un rapport. */
export interface CreateReportInput {
  controleurEmail: string
  controleurName: string
  date: string
  lines: ActivityLine[]
  anomaliesDetectees?: number
  lienRapport?: string
  observationsGlobales?: string
  attachments?: { name: string; url: string }[]
}

/** Format SharePoint Claims pour les champs Personne/Groupe. */
function toClaims(email: string): string {
  return `i:0#.f|membership|${email}`
}

/**
 * Crée un nouvel item SharePoint dans DCPO_ACTIVICTE_CONTROLLER.
 *
 * Étapes :
 *   1. Sérialise les lignes vers le format hybride humain+JSON
 *   2. Construit le payload SP (controlleur en format Claims)
 *   3. POST via le service généré
 *   4. Retourne le rapport reconstruit depuis l'item créé
 */
export async function createReport(input: CreateReportInput): Promise<ActivityReport> {
  const actionText = serializeLines(input.lines)
  const title = `Rapport ${input.date} - ${input.controleurName}`
  const dateISO = `${input.date}T00:00:00Z`

  const payload: Record<string, unknown> = {
    Title: title,
    Date: dateISO,
    actionDeLaJournee: actionText,
    nombreAnomalieDetectee: input.anomaliesDetectees ?? 0,
    observationsGlobales: input.observationsGlobales ?? '',
  }
  if (input.controleurEmail) {
    payload.controlleur = {
      '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
      Claims: toClaims(input.controleurEmail),
    }
  }

  const result = await DCPO_ACTIVICTE_CONTROLLERService.create(
    payload as Omit<DCPO_ACTIVICTE_CONTROLLERWrite, 'ID'>,
  )
  if (!result.success || !result.data) {
    throw new Error(result.error?.message ?? 'Échec de la création du rapport.')
  }
  return reportFromItem(result.data)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 8 — MISE À JOUR D'UN RAPPORT
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Patch un rapport existant en SharePoint.
 *
 * Limitation : les calculs dérivés (totalHeures, tempsOccupe, statutJournee)
 * sont recalculés à la lecture, donc pas besoin de les mettre à jour ici.
 * Si patch.lines est fourni, on ré-encode via serializeLines().
 */
export async function updateReport(id: string, patch: Partial<ActivityReport>): Promise<ActivityReport | undefined> {
  const payload: Record<string, unknown> = {}
  if (patch.lines) payload.actionDeLaJournee = serializeLines(patch.lines)
  if (patch.anomaliesDetectees !== undefined) payload.nombreAnomalieDetectee = patch.anomaliesDetectees
  if (patch.observationsGlobales !== undefined) payload.observationsGlobales = patch.observationsGlobales
  if (patch.date) payload.Date = `${patch.date}T00:00:00Z`

  if (Object.keys(payload).length > 0) {
    try {
      await DCPO_ACTIVICTE_CONTROLLERService.update(id, payload as never)
    } catch (err) {
      console.error('updateReport error', err)
      return undefined
    }
  }
  return getReport(id)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 9 — VALIDATION MANAGER (en localStorage)
 *
 * Le statut et l'historique de validation sont stockés en localStorage
 * (indexés par l'ID SharePoint) tant que les colonnes correspondantes ne
 * sont pas ajoutées à la liste SharePoint.
 *
 * Migration future : remplacer ces deux fonctions par des appels à
 * SP_SERVICE.update() sur les futures colonnes statut + historiqueValidations.
 * ────────────────────────────────────────────────────────────────────────── */

export interface ValidationInput {
  manager: string
  managerEmail?: string
  decision: 'Réalisé' | 'Reporté'
  motif?: string
}

/**
 * Applique une décision manager. Append-only sur l'historique.
 * Retourne le rapport rechargé (avec la validation à jour).
 */
export async function validateReport(id: string, input: ValidationInput): Promise<ActivityReport | undefined> {
  const existing = readValidation(id) ?? { statut: 'Soumis', history: [] }
  const note: ActivityValidationNote = {
    date: new Date().toISOString(),
    manager: input.manager,
    managerEmail: input.managerEmail,
    decision: input.decision,
    motif: input.motif,
  }
  writeValidation(id, {
    statut: input.decision,
    history: [...existing.history, note],
  })
  return getReport(id)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 10 — BROUILLONS (localStorage)
 * ────────────────────────────────────────────────────────────────────────── */

export function loadDraft(controleurEmail: string): ActivityDraft | undefined {
  const raw = READ_RAW(STORAGE_KEY_DRAFT_PREFIX + controleurEmail.toLowerCase())
  if (!raw) return undefined
  try {
    return JSON.parse(raw) as ActivityDraft
  } catch {
    return undefined
  }
}

export function saveDraft(draft: ActivityDraft): void {
  WRITE_RAW(STORAGE_KEY_DRAFT_PREFIX + draft.controleurEmail.toLowerCase(), JSON.stringify({
    ...draft,
    updatedAt: new Date().toISOString(),
  }))
}

export function clearDraft(controleurEmail: string): void {
  REMOVE_RAW(STORAGE_KEY_DRAFT_PREFIX + controleurEmail.toLowerCase())
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 11 — HELPERS UI (inchangés)
 * ────────────────────────────────────────────────────────────────────────── */

export function makeEmptyLine(): ActivityLine {
  return { id: uid(), domaine: '', objet: '', action: '', duree: 30, observation: '' }
}

export const ACTIVITY_DOMAINES = [
  'Contrôle',
  'Reporting',
  'Investigation',
  'Réunion',
  'Formation',
  'Audit',
  'Suivi anomalie',
  'Administratif',
  'Autre',
]

export const ACTIVITY_LINE_LIMITS = {
  actionMinLength: 3,
  actionMaxLength: 500,
  dureeMin: 5,
  dureeMax: 600,
}

export interface LineValidationError {
  index: number
  field: keyof ActivityLine
  message: string
}

export function validateLines(lines: ActivityLine[]): LineValidationError[] {
  const errors: LineValidationError[] = []
  if (lines.length === 0) {
    errors.push({ index: -1, field: 'action', message: 'Au moins une ligne est requise.' })
    return errors
  }
  lines.forEach((line, index) => {
    if (!line.domaine) errors.push({ index, field: 'domaine', message: 'Domaine requis.' })
    if (!line.objet?.trim()) errors.push({ index, field: 'objet', message: 'Objet requis.' })
    const actionLen = line.action?.trim().length ?? 0
    if (actionLen < ACTIVITY_LINE_LIMITS.actionMinLength) {
      errors.push({ index, field: 'action', message: `Action: ${ACTIVITY_LINE_LIMITS.actionMinLength} caractères minimum.` })
    } else if (actionLen > ACTIVITY_LINE_LIMITS.actionMaxLength) {
      errors.push({ index, field: 'action', message: `Action: ${ACTIVITY_LINE_LIMITS.actionMaxLength} caractères maximum.` })
    }
    const duree = Number(line.duree)
    if (!Number.isFinite(duree) || duree < ACTIVITY_LINE_LIMITS.dureeMin) {
      errors.push({ index, field: 'duree', message: `Durée: ${ACTIVITY_LINE_LIMITS.dureeMin} min minimum.` })
    } else if (duree > ACTIVITY_LINE_LIMITS.dureeMax) {
      errors.push({ index, field: 'duree', message: `Durée: ${ACTIVITY_LINE_LIMITS.dureeMax} min maximum.` })
    }
  })
  return errors
}
