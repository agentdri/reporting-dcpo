/**
 * ============================================================================
 * SERVICE — PLAN D'ACTION CORRECTIF (PAC)
 * ============================================================================
 *
 * Persistance des PAC : liste SharePoint DCPO_LISTE_PLAN_ACTION_CORRECTIF.
 *
 * Mapping des colonnes SharePoint :
 *   ┌──────────────────────────┬────────────────────────┬───────────────────┐
 *   │ Colonne SP               │ Champ domaine          │ Notes             │
 *   ├──────────────────────────┼────────────────────────┼───────────────────┤
 *   │ Title                    │ intitule               │                   │
 *   │ field_1                  │ sourcePac              │                   │
 *   │ field_2                  │ dateCreation           │ texte (YYYY-MM-DD) │
 *   │ field_3                  │ descriptionProbleme    │                   │
 *   │ field_4                  │ causeImmediate         │                   │
 *   │ field_5                  │ causeRacine            │                   │
 *   │ field_6                  │ actionsCorrectives     │                   │
 *   │ field_7                  │ directionsConcernees   │ codes joints ';'  │
 *   │ field_8                  │ echeance               │ texte (YYYY-MM-DD) │
 *   │ field_9 (number)         │ annee                  │                   │
 *   │ field_11                 │ statut                 │                   │
 *   │ field_12                 │ kpi                    │                   │
 *   │ field_13                 │ observations           │                   │
 *   │ responsableMiseEnOeuvre  │ responsable*           │ champ PERSONNE     │
 *   └──────────────────────────┴────────────────────────┴───────────────────┘
 *
 * Référentiel des DIRECTIONS : conservé en localStorage / constante (aucune
 * liste SharePoint dédiée n'a été créée pour les directions). Seuls les PAC
 * eux-mêmes sont persistés en SharePoint.
 * ============================================================================
 */

import { DCPO_LISTE_PLAN_ACTION_CORRECTIFService } from '../generated/services/DCPO_LISTE_PLAN_ACTION_CORRECTIFService'
import type {
  DCPO_LISTE_PLAN_ACTION_CORRECTIFRead,
  DCPO_LISTE_PLAN_ACTION_CORRECTIFWrite,
} from '../generated/models/DCPO_LISTE_PLAN_ACTION_CORRECTIFModel'
import { appendUrl, parseUrlList, getFileNameFromUrl } from './ticketAttachments'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE
 * ────────────────────────────────────────────────────────────────────────── */

/** Statut workflow d'un PAC (colonne field_11). */
export type PacStatus = 'Exécutée' | 'En cours' | 'Non Exécutée'

/** Direction concernée (référentiel local, pas de liste SharePoint dédiée). */
export interface Direction {
  code: string
  libelle: string
  actif: boolean
}

/**
 * Représentation d'un Plan d'Action Correctif.
 *
 * Le responsable est un champ Personne côté SharePoint : on garde son nom
 * affiché (responsable) + son email (responsableEmail, pour l'écriture).
 */
export interface Pac {
  id: string
  sourcePac: string
  dateCreation: string         // YYYY-MM-DD
  intitule: string
  descriptionProbleme: string
  causeImmediate: string
  causeRacine: string
  actionsCorrectives: string
  directionsConcernees: string[]
  echeance: string             // YYYY-MM-DD
  kpi: string
  annee: number
  responsable: string          // DisplayName (affichage)
  responsableEmail: string     // email (écriture Person)
  statut: PacStatus
  observations?: string
  /**
   * Pièces jointes décodées depuis le champ urlPiecesJointes (multi-URLs
   * séparées par " | ", cf. helper appendUrl/parseUrlList). Reconstruit à
   * la lecture pour faciliter l'affichage côté UI.
   */
  attachments?: { name: string; url: string }[]
  createdAt: string
  updatedAt: string
}

/** Données pour créer un PAC. */
export interface CreatePacInput {
  sourcePac: string
  dateCreation: string
  intitule: string
  descriptionProbleme: string
  causeImmediate: string
  causeRacine: string
  actionsCorrectives: string
  directionsConcernees: string[]
  echeance: string
  kpi: string
  annee: number
  responsableName: string
  responsableEmail: string
  statut: PacStatus
  observations?: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — RÉFÉRENTIEL DES DIRECTIONS (localStorage / constante)
 *
 * Pas de liste SharePoint dédiée → on garde les directions côté client.
 * Pour préconfigurer, éditer DEFAULT_DIRECTIONS. Pour modifier dynamiquement,
 * utiliser upsertDirection / deactivateDirection (persistés en localStorage).
 * ────────────────────────────────────────────────────────────────────────── */

const STORAGE_KEY_DIRECTIONS = 'reportingDCPO.directions.v1'

const DEFAULT_DIRECTIONS: Direction[] = [
  { code: 'DJC', libelle: 'Direction Juridique et Conformité', actif: true },
  { code: 'DCE', libelle: 'Direction Commercial Entreprises', actif: true },
  { code: 'DSI', libelle: 'Direction des Systèmes d\'Information', actif: true },
]

const readDirections = (): Direction[] => {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY_DIRECTIONS)
    return raw ? (JSON.parse(raw) as Direction[]) : []
  } catch {
    return []
  }
}

const writeDirections = (dirs: Direction[]): void => {
  try {
    window.localStorage.setItem(STORAGE_KEY_DIRECTIONS, JSON.stringify(dirs))
  } catch {
    /* ignore */
  }
}

/** Liste les directions (initialise avec DEFAULT_DIRECTIONS au premier accès). */
export function listDirections(activeOnly: boolean = true): Direction[] {
  const all = readDirections()
  if (all.length === 0) {
    writeDirections(DEFAULT_DIRECTIONS)
    return activeOnly ? DEFAULT_DIRECTIONS.filter(d => d.actif) : DEFAULT_DIRECTIONS
  }
  const filtered = activeOnly ? all.filter(d => d.actif) : all
  return [...filtered].sort((a, b) => a.code.localeCompare(b.code))
}

/** Crée ou met à jour une direction (clé = code). */
export function upsertDirection(direction: Direction): void {
  const all = readDirections()
  const idx = all.findIndex(d => d.code === direction.code)
  if (idx >= 0) all[idx] = direction
  else all.push(direction)
  writeDirections(all)
}

/** Désactive une direction (soft delete). */
export function deactivateDirection(code: string): void {
  const all = readDirections()
  const idx = all.findIndex(d => d.code === code)
  if (idx >= 0) {
    all[idx] = { ...all[idx], actif: false }
    writeDirections(all)
  }
}

/** Résout le libellé d'un code de direction (fallback = code). */
export function getDirectionLabel(code: string): string {
  const all = readDirections()
  const source = all.length > 0 ? all : DEFAULT_DIRECTIONS
  return source.find(d => d.code === code)?.libelle ?? code
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — MAPPING SHAREPOINT ↔ DOMAINE (PAC)
 * ────────────────────────────────────────────────────────────────────────── */

/** Format SharePoint Claims pour un champ Personne. */
function toClaims(email: string): string {
  return `i:0#.f|membership|${email}`
}

/** Sépare/normalise les codes directions stockés en texte ("DSI;DJC"). */
function parseDirections(raw: string | undefined): string[] {
  if (!raw) return []
  return raw.split(';').map(s => s.trim()).filter(Boolean)
}

/** Normalise une date SP (peut être ISO ou déjà YYYY-MM-DD) → YYYY-MM-DD. */
function toDateOnly(raw: string | undefined): string {
  if (!raw) return ''
  return raw.split('T')[0]
}

/** Construit un Pac depuis un item SharePoint. */
function fromItem(item: DCPO_LISTE_PLAN_ACTION_CORRECTIFRead): Pac {
  const statut = item.field_11 as PacStatus
  // Pièces jointes : le champ urlPiecesJointes (multi-URLs concaténées par
  // " | ") est parsé puis transformé en { name, url }[] pour l'affichage UI.
  const attachmentUrls = parseUrlList(item.urlPiecesJointes)
  const attachments = attachmentUrls.map(url => ({
    name: getFileNameFromUrl(url),
    url,
  }))
  return {
    id: String(item.ID),
    intitule: item.Title ?? '',
    sourcePac: item.field_1 ?? '',
    dateCreation: toDateOnly(item.field_2),
    descriptionProbleme: item.field_3 ?? '',
    causeImmediate: item.field_4 ?? '',
    causeRacine: item.field_5 ?? '',
    actionsCorrectives: item.field_6 ?? '',
    directionsConcernees: parseDirections(item.field_7),
    echeance: toDateOnly(item.field_8),
    annee: item.field_9 ?? new Date().getFullYear(),
    statut: PAC_STATUS_OPTIONS.includes(statut) ? statut : 'En cours',
    kpi: item.field_12 ?? '',
    observations: item.field_13 || undefined,
    responsable: item.responsableMiseEnOeuvre?.DisplayName ?? '',
    responsableEmail: item.responsableMiseEnOeuvre?.Email ?? '',
    attachments,
    createdAt: item.Created ?? '',
    updatedAt: item.Modified ?? '',
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — API PUBLIQUE (PAC)
 * ────────────────────────────────────────────────────────────────────────── */

/** Liste tous les PAC (du plus récent au plus ancien). */
export async function listPACs(): Promise<Pac[]> {
  try {
    const res = await DCPO_LISTE_PLAN_ACTION_CORRECTIFService.getAll({ orderBy: ['Created desc'] })
    if (!res.data) return []
    return res.data.map(fromItem)
  } catch (err) {
    console.error('listPACs error', err)
    return []
  }
}

/** Récupère un PAC par ID. */
export async function getPAC(id: string): Promise<Pac | undefined> {
  try {
    const res = await DCPO_LISTE_PLAN_ACTION_CORRECTIFService.get(id)
    if (!res.data) return undefined
    return fromItem(res.data)
  } catch (err) {
    console.error('getPAC error', err)
    return undefined
  }
}

/**
 * Crée un PAC dans SharePoint.
 *
 * Validation minimale :
 *   - intitule obligatoire
 *   - au moins une direction concernée
 */
export async function createPAC(input: CreatePacInput): Promise<Pac> {
  if (!input.intitule.trim()) throw new Error('Intitulé obligatoire.')
  if (input.directionsConcernees.length === 0) {
    throw new Error('Au moins une direction concernée est requise.')
  }

  const payload: Record<string, unknown> = {
    Title: input.intitule.trim(),
    field_1: input.sourcePac.trim(),
    field_2: input.dateCreation,
    field_3: input.descriptionProbleme.trim(),
    field_4: input.causeImmediate.trim(),
    field_5: input.causeRacine.trim(),
    field_6: input.actionsCorrectives.trim(),
    field_7: input.directionsConcernees.join(';'),
    field_8: input.echeance,
    field_9: input.annee,
    field_11: input.statut,
    field_12: input.kpi.trim(),
    field_13: input.observations?.trim() ?? '',
  }
  // Champ Personne : responsable de mise en œuvre.
  if (input.responsableEmail) {
    payload.responsableMiseEnOeuvre = {
      '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
      Claims: toClaims(input.responsableEmail),
    }
  }

  const res = await DCPO_LISTE_PLAN_ACTION_CORRECTIFService.create(
    payload as Omit<DCPO_LISTE_PLAN_ACTION_CORRECTIFWrite, 'ID'>,
  )
  if (!res.success || !res.data) {
    throw new Error(res.error?.message ?? 'Échec de la création du PAC.')
  }
  return fromItem(res.data)
}

/**
 * Met à jour un PAC existant (champs métier — pas les pièces jointes).
 *
 * Pour la modification des pièces jointes, utiliser `appendPacAttachmentUrls`
 * (ajout) — la suppression de PJ n'est pas exposée ici.
 *
 * @returns Le PAC rechargé après update, ou undefined en cas d'erreur réseau.
 */
export async function updatePAC(id: string, input: CreatePacInput): Promise<Pac | undefined> {
  const payload: Record<string, unknown> = {
    Title: input.intitule.trim(),
    field_1: input.sourcePac.trim(),
    field_2: input.dateCreation,
    field_3: input.descriptionProbleme.trim(),
    field_4: input.causeImmediate.trim(),
    field_5: input.causeRacine.trim(),
    field_6: input.actionsCorrectives.trim(),
    field_7: input.directionsConcernees.join(';'),
    field_8: input.echeance,
    field_9: input.annee,
    field_11: input.statut,
    field_12: input.kpi.trim(),
    field_13: input.observations?.trim() ?? '',
  }
  if (input.responsableEmail) {
    payload.responsableMiseEnOeuvre = {
      '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
      Claims: toClaims(input.responsableEmail),
    }
  }
  try {
    await DCPO_LISTE_PLAN_ACTION_CORRECTIFService.update(
      id,
      payload as Partial<Omit<DCPO_LISTE_PLAN_ACTION_CORRECTIFWrite, 'ID'>>,
    )
  } catch (err) {
    console.error('updatePAC error', err)
    return undefined
  }
  return getPAC(id)
}

/**
 * Ajoute (concatène) une ou plusieurs URLs de pièces jointes au champ
 * `urlPiecesJointes` du PAC, sans écraser celles existantes.
 *
 * Même pattern que pour les anomalies (cf. appendUrl) : les URLs sont
 * concaténées par " | " dans le champ texte SP. Le helper appendUrl gère
 * le dédoublonnage et la mise en forme.
 *
 * Étapes :
 *   1. Lire l'item courant pour récupérer la valeur existante
 *   2. Pour chaque nouvelle URL, l'ajouter via appendUrl (dédoublonnage auto)
 *   3. Mettre à jour SharePoint avec la chaîne concaténée
 *
 * @param pacId ID SharePoint du PAC cible
 * @param newUrls URLs à ajouter (typiquement celles retournées par l'upload
 *                Power Automate)
 */
export async function appendPacAttachmentUrls(pacId: string, newUrls: string[]): Promise<void> {
  const cleanUrls = newUrls.filter(u => !!u && u.trim().length > 0)
  if (cleanUrls.length === 0) return

  // 1. Lire l'item pour préserver les URLs déjà stockées.
  let existing = ''
  try {
    const res = await DCPO_LISTE_PLAN_ACTION_CORRECTIFService.get(pacId)
    existing = res.data?.urlPiecesJointes ?? ''
  } catch (err) {
    console.error('appendPacAttachmentUrls: échec lecture item', err)
    // On continue avec une chaîne vide — pire cas : on perd l'ancien, mais
    // c'est mieux que de ne pas écrire les nouvelles.
  }

  // 2. Concaténer chaque URL via appendUrl (dédoublonne et formate).
  const concatenated = cleanUrls.reduce(
    (acc, url) => appendUrl(acc, url),
    existing,
  )

  // 3. Persister la nouvelle valeur en SharePoint.
  try {
    await DCPO_LISTE_PLAN_ACTION_CORRECTIFService.update(pacId, {
      urlPiecesJointes: concatenated,
    } as Partial<Omit<DCPO_LISTE_PLAN_ACTION_CORRECTIFWrite, 'ID'>>)
  } catch (err) {
    console.error('appendPacAttachmentUrls: échec update item', err)
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — HELPERS UI
 * ────────────────────────────────────────────────────────────────────────── */

export const PAC_STATUS_OPTIONS: PacStatus[] = ['En cours', 'Exécutée', 'Non Exécutée']

/** Mappe un statut PAC vers une classe CSS (pastilles manager-pill). */
export function getPacStatusClass(statut: PacStatus): string {
  switch (statut) {
    case 'Exécutée': return 'manager-pill-realise'
    case 'En cours': return 'manager-pill-pending'
    case 'Non Exécutée': return 'manager-pill-reporte'
    default: return 'manager-pill'
  }
}
