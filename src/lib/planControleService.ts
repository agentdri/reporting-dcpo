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
 * SECTION 5 — HELPERS UI
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
