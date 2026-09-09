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
import { DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService } from '../generated/services/DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService'
import type {
  DCPO_EVALUATION_PLAN_ACTION_CORRECTIFRead,
  DCPO_EVALUATION_PLAN_ACTION_CORRECTIFWrite,
} from '../generated/models/DCPO_EVALUATION_PLAN_ACTION_CORRECTIFModel'
import type {
  DCPO_LISTE_PLAN_ACTION_CORRECTIFRead,
  DCPO_LISTE_PLAN_ACTION_CORRECTIFWrite,
} from '../generated/models/DCPO_LISTE_PLAN_ACTION_CORRECTIFModel'
import { appendUrl, parseUrlList, getFileNameFromUrl } from './ticketAttachments'
import { repairMojibakeDeep } from './textEncoding'
import { getAllPages } from './sharePointPaging'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE
 * ────────────────────────────────────────────────────────────────────────── */

/** Statut workflow d'un PAC (colonne field_11). */
export type PacStatus = 'Exécutée' | 'En cours' | 'Non Exécutée'

// Note : le type `Direction` et la gestion CRUD du référentiel des directions
// sont désormais dans src/lib/directionService.ts (liste SharePoint dédiée
// DCPO_LISTE_DIRECTION). Les PAC stockent toujours des sigles dans
// directionsConcernees ; la résolution des libellés se fait côté composant
// avec la liste pré-chargée.

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
 * SECTION 2 — MAPPING SHAREPOINT ↔ DOMAINE (PAC)
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
  // Réparation des accents mal encodés (cf. textEncoding.ts)
  item = repairMojibakeDeep(item)
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
    const items = await getAllPages<DCPO_LISTE_PLAN_ACTION_CORRECTIFRead>(
      DCPO_LISTE_PLAN_ACTION_CORRECTIFService,
      { orderBy: ['Created desc'] },
    )
    return items.map(fromItem)
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
    return fromItem(repairMojibakeDeep(res.data))
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
/**
 * Met à jour UNIQUEMENT le responsable d'un PAC (affectation rapide).
 *
 * Cas d'usage : workflow "Affectation" depuis le tableau — un manager change
 * la personne responsable de mise en œuvre sans avoir à rouvrir le formulaire
 * d'édition complet.
 *
 * @param id - ID SharePoint du PAC
 * @param email - Email de la nouvelle personne responsable (format
 *                'jdoe@afrilandfirstbank.com' — le format Claims est ajouté ici)
 * @returns Le PAC rechargé après update, ou undefined en cas d'erreur.
 */
export async function updatePacResponsable(
  id: string,
  email: string,
): Promise<Pac | undefined> {
  if (!email) return undefined
  try {
    await DCPO_LISTE_PLAN_ACTION_CORRECTIFService.update(
      id,
      {
        responsableMiseEnOeuvre: {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(email),
        },
      } as never,
    )
  } catch (err) {
    console.error('updatePacResponsable error', err)
    return undefined
  }
  return getPAC(id)
}

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

/**
 * Indique si un PAC est encore évaluable.
 *
 * Règle métier : dès qu'un PAC est en statut "Exécutée" ou "Non Exécutée",
 * il est CLÔTURÉ — l'agent responsable ne peut plus consigner d'évaluation.
 * Les évaluations passées restent consultables dans l'historique.
 *
 * Seul le statut "En cours" reste évaluable (suivi périodique).
 */
export function isPacEvaluable(statut: PacStatus): boolean {
  return statut === 'En cours'
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 — ÉVALUATIONS DU PLAN D'ACTION CORRECTIF
 *
 * Liste SharePoint DCPO_EVALUATION_PLAN_ACTION_CORRECTIF.
 *
 * Mapping :
 *   ┌─────────────────────────────┬────────────────────┐
 *   │ Title                       │ titre              │
 *   │ observations                │ observations       │
 *   │ plan_action_correctif_id    │ pacId (string SP)  │
 *   │ urlPieceJointe              │ attachments (URLs concaténées)
 *   └─────────────────────────────┴────────────────────┘
 *
 * Différence notable avec les évaluations de Plan de Contrôle : PAS de
 * champ période (un PAC est ponctuel, pas récurrent — on évalue son
 * avancement à un moment T, sans notion de jour/semaine/mois/année).
 * ────────────────────────────────────────────────────────────────────────── */

/** Évaluation d'un PAC (vue front). */
export interface PacEvaluation {
  id: string
  titre: string
  pacId: string
  observations: string
  /** Pièces jointes décodées depuis urlPieceJointe (concat " | "). */
  attachments: { name: string; url: string }[]
  createdAt: string
  updatedAt: string
}

/** Input pour créer une évaluation PAC. */
export interface CreatePacEvaluationInput {
  pacId: string
  observations: string
  titre?: string
}

/** Mapping item SP → PacEvaluation. */
function fromPacEvaluationItem(item: DCPO_EVALUATION_PLAN_ACTION_CORRECTIFRead): PacEvaluation {
  // Réparation des accents mal encodés (cf. textEncoding.ts)
  item = repairMojibakeDeep(item)
  const urls = parseUrlList(item.urlPieceJointe)
  return {
    id: String(item.ID),
    titre: item.Title ?? '',
    pacId: item.plan_action_correctif_id ?? '',
    observations: item.observations ?? '',
    attachments: urls.map(url => ({ name: getFileNameFromUrl(url), url })),
    createdAt: item.Created ?? '',
    updatedAt: item.Modified ?? '',
  }
}

/**
 * Liste les évaluations d'un PAC donné (les plus récentes d'abord).
 *
 * Filtre serveur OData sur `plan_action_correctif_id`. La colonne SP est
 * typée `string` → on encode l'ID entre quotes pour la comparaison.
 */
export async function listEvaluationsForPac(pacId: string | number): Promise<PacEvaluation[]> {
  const idStr = String(pacId)
  if (!idStr) return []
  try {
    const items = await getAllPages<DCPO_EVALUATION_PLAN_ACTION_CORRECTIFRead>(
      DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService,
      {
        filter: `plan_action_correctif_id eq '${idStr}'`,
        orderBy: ['Created desc'],
      },
    )
    return items.map(fromPacEvaluationItem)
  } catch (err) {
    console.error('listEvaluationsForPac error', err)
    return []
  }
}

/** Liste TOUTES les évaluations PAC (pour agrégation côté Reporting Agent). */
export async function listAllPacEvaluations(): Promise<PacEvaluation[]> {
  try {
    const items = await getAllPages<DCPO_EVALUATION_PLAN_ACTION_CORRECTIFRead>(
      DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService,
      { orderBy: ['Created desc'] },
    )
    return items.map(fromPacEvaluationItem)
  } catch (err) {
    console.error('listAllPacEvaluations error', err)
    return []
  }
}

/**
 * Crée une évaluation pour un PAC.
 *
 * Le `Title` SP est obligatoire — si pas fourni, on génère un libellé par
 * défaut avec la date du jour.
 */
export async function createPacEvaluation(input: CreatePacEvaluationInput): Promise<PacEvaluation> {
  if (!input.pacId) throw new Error('PAC parent manquant.')
  if (!input.observations.trim()) throw new Error('Les observations sont obligatoires.')

  const today = new Date().toISOString().slice(0, 10)
  const title = (input.titre ?? '').trim() || `Évaluation PAC ${input.pacId} - ${today}`

  const payload: Record<string, unknown> = {
    Title: title,
    plan_action_correctif_id: input.pacId,
    observations: input.observations.trim(),
  }

  const res = await DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService.create(
    payload as Omit<DCPO_EVALUATION_PLAN_ACTION_CORRECTIFWrite, 'ID'>,
  )
  if (!res.success || !res.data) {
    throw new Error(res.error?.message ?? "Échec de la création de l'évaluation.")
  }
  return fromPacEvaluationItem(res.data)
}

/**
 * Ajoute (concatène) une ou plusieurs URLs de pièces jointes au champ
 * `urlPieceJointe` d'une évaluation PAC, sans écraser celles existantes.
 *
 * Workflow type :
 *   1. Créer l'évaluation via createPacEvaluation (récupère l'ID)
 *   2. Uploader le fichier via uploadPacEvaluationAttachment (récupère l'URL)
 *   3. Appeler ce helper avec l'ID + l'URL pour la persister
 *
 * Cohérent avec les patterns appendXxxAttachmentUrls existants (anomalies,
 * PAC, plan de contrôle, évaluations plan de contrôle).
 */
export async function appendPacEvaluationAttachmentUrls(
  evaluationId: string,
  newUrls: string[],
): Promise<void> {
  const cleanUrls = newUrls.filter(u => !!u && u.trim().length > 0)
  if (cleanUrls.length === 0) return

  let existing = ''
  try {
    const res = await DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService.get(evaluationId)
    existing = res.data?.urlPieceJointe ?? ''
  } catch (err) {
    console.error('appendPacEvaluationAttachmentUrls: échec lecture item', err)
  }

  const concatenated = cleanUrls.reduce(
    (acc, url) => appendUrl(acc, url),
    existing,
  )

  try {
    await DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService.update(evaluationId, {
      urlPieceJointe: concatenated,
    } as Partial<Omit<DCPO_EVALUATION_PLAN_ACTION_CORRECTIFWrite, 'ID'>>)
  } catch (err) {
    console.error('appendPacEvaluationAttachmentUrls: échec update item', err)
  }
}
