/**
 * ============================================================================
 * ABSENCES — Gestion des périodes d'absence des contrôleurs
 * ============================================================================
 *
 * Une absence = une période (dateDebut → dateFin inclusifs) sur laquelle un
 * contrôleur n'est PAS censé soumettre son rapport quotidien. Les jours
 * ouvrés couverts par une absence sont retirés du dénominateur du taux
 * de soumission (cf. ControllerReportingList).
 *
 * Motifs possibles (référentiel Choice côté SP) :
 *   - Congé annuel
 *   - Maladie
 *   - Formation
 *   - Mission
 *   - Autre
 *
 * Persistance : liste SharePoint `DCPO_LISTE_ABSENCES` avec les colonnes :
 *   - controleur (Person)
 *   - dateDebut (date ISO)
 *   - dateFin (date ISO)
 *   - motif (choice)
 *   - commentaire (string, optionnel)
 *   - declarePar (Person, auto)
 *
 * Cas particulier "intersection avec la période analysée" :
 *   Si absence 10/07 → 20/07 et période analysée 15/07 → 30/07,
 *   seuls les jours 15, 16, 17, 18, 19, 20 (intersection) sont
 *   retirés du dénominateur. Géré par `getAbsentDaysForControleur`.
 * ============================================================================
 */

import { DCPO_LISTE_ABSENCESService } from '../generated/services/DCPO_LISTE_ABSENCESService'
import type {
  DCPO_LISTE_ABSENCESRead,
  DCPO_LISTE_ABSENCESWrite,
} from '../generated/models/DCPO_LISTE_ABSENCESModel'
import { getAllPages } from './sharePointPaging'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES MÉTIER
 * ────────────────────────────────────────────────────────────────────────── */

/** Motifs autorisés — DOIT rester synchro avec le référentiel Choice SP. */
export const ABSENCE_MOTIFS = [
  'Congé annuel',
  'Maladie',
  'Formation',
  'Mission',
  'Autre',
] as const

export type AbsenceMotif = typeof ABSENCE_MOTIFS[number]

/**
 * Représentation métier d'une absence — abstrait le format SP brut.
 * Les dates sont au format `YYYY-MM-DD` (calendrier, pas de timezone).
 */
export interface Absence {
  id: string
  controleurEmail: string
  controleurName: string
  dateDebut: string       // YYYY-MM-DD
  dateFin: string         // YYYY-MM-DD
  motif: string
  commentaire?: string
  declareParName?: string
  declareParEmail?: string
  createdAt?: string
}

/** Payload de création (declarePar est renseigné automatiquement par saveAbsence). */
export interface CreateAbsenceInput {
  controleurEmail: string
  controleurName: string
  dateDebut: string       // YYYY-MM-DD
  dateFin: string         // YYYY-MM-DD
  motif: string
  commentaire?: string
  declareParEmail: string
  declareParName: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — HELPERS D'ENCODAGE (SP ↔ Absence)
 * ────────────────────────────────────────────────────────────────────────── */

/** Format Person SP standard. */
function toClaims(email: string): string {
  return `i:0#.f|membership|${email.toLowerCase()}`
}

/** Extrait la portion YYYY-MM-DD d'une date ISO SP. Retourne '' si vide. */
function toDateOnly(value?: string | null): string {
  if (!value) return ''
  const m = value.match(/^(\d{4}-\d{2}-\d{2})/)
  return m ? m[1] : ''
}

/** Mappe un item SP brut en Absence métier. */
function fromItem(item: DCPO_LISTE_ABSENCESRead): Absence {
  return {
    id: String(item.ID ?? ''),
    controleurEmail: item.controleur?.Email ?? '',
    controleurName: item.controleur?.DisplayName ?? item.controleur?.Email ?? '',
    dateDebut: toDateOnly(item.dateDebut),
    dateFin: toDateOnly(item.dateFin),
    motif: item.motif ?? '',
    commentaire: item.commentaire ?? undefined,
    declareParName: item.declarePar?.DisplayName ?? undefined,
    declareParEmail: item.declarePar?.Email ?? undefined,
    createdAt: item.Created ?? undefined,
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — API PUBLIQUE (CRUD)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Liste toutes les absences (triées de la plus récente à la plus ancienne
 * par date de début). Utilise getAllPages pour rapatrier toutes les pages.
 */
export async function listAbsences(): Promise<Absence[]> {
  try {
    const items = await getAllPages<DCPO_LISTE_ABSENCESRead>(
      DCPO_LISTE_ABSENCESService,
      { orderBy: ['dateDebut desc'] },
    )
    return items.map(fromItem)
  } catch (err) {
    console.error('listAbsences error', err)
    return []
  }
}

/**
 * Crée une nouvelle absence. Renseigne automatiquement `declarePar` avec
 * l'utilisateur passé en paramètre. Retourne l'absence créée (avec ID SP).
 */
export async function createAbsence(input: CreateAbsenceInput): Promise<Absence | null> {
  try {
    const payload: Partial<DCPO_LISTE_ABSENCESWrite> = {
      Title: `${input.motif} — ${input.controleurName} — ${input.dateDebut} → ${input.dateFin}`,
      dateDebut: `${input.dateDebut}T00:00:00Z`,
      dateFin: `${input.dateFin}T00:00:00Z`,
      motif: input.motif,
      commentaire: input.commentaire?.trim() || undefined,
      controleur: {
        '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
        Claims: toClaims(input.controleurEmail),
      } as never,
      declarePar: {
        '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
        Claims: toClaims(input.declareParEmail),
      } as never,
    }
    const result = await DCPO_LISTE_ABSENCESService.create(
      payload as Omit<DCPO_LISTE_ABSENCESWrite, 'ID'>,
    )
    if (!result.success || !result.data) return null
    return fromItem(result.data)
  } catch (err) {
    console.error('createAbsence error', err)
    return null
  }
}

/**
 * Met à jour une absence existante. `controleur` et `declarePar` ne sont
 * modifiés que s'ils sont fournis (utile pour changer uniquement les dates
 * ou le motif sans repasser tous les champs).
 */
export async function updateAbsence(
  id: string,
  changes: Partial<CreateAbsenceInput>,
): Promise<Absence | null> {
  try {
    const payload: Partial<DCPO_LISTE_ABSENCESWrite> = {}
    if (changes.dateDebut) payload.dateDebut = `${changes.dateDebut}T00:00:00Z`
    if (changes.dateFin) payload.dateFin = `${changes.dateFin}T00:00:00Z`
    if (changes.motif) payload.motif = changes.motif
    if (changes.commentaire !== undefined) {
      payload.commentaire = changes.commentaire.trim() || undefined
    }
    if (changes.controleurEmail) {
      payload.controleur = {
        '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
        Claims: toClaims(changes.controleurEmail),
      } as never
    }
    // Rafraîchir aussi le Title pour rester lisible côté SP
    if (changes.motif && changes.controleurName && changes.dateDebut && changes.dateFin) {
      payload.Title = `${changes.motif} — ${changes.controleurName} — ${changes.dateDebut} → ${changes.dateFin}`
    }
    const result = await DCPO_LISTE_ABSENCESService.update(id, payload)
    if (!result.success || !result.data) return null
    return fromItem(result.data)
  } catch (err) {
    console.error('updateAbsence error', err)
    return null
  }
}

/** Supprime une absence par ID. Retourne true si succès. */
export async function deleteAbsence(id: string): Promise<boolean> {
  try {
    await DCPO_LISTE_ABSENCESService.delete(id)
    return true
  } catch (err) {
    console.error('deleteAbsence error', err)
    return false
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — HELPERS DE CALCUL (intersection période × absence)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Retourne l'ensemble des jours ouvrés (Lun-Ven) où un contrôleur est
 * absent, DANS LA PÉRIODE ANALYSÉE [from, to].
 *
 * Algorithme :
 *   1. Filtrer les absences du contrôleur qui chevauchent la période
 *   2. Pour chaque absence, calculer l'INTERSECTION avec la période
 *   3. Énumérer les jours ouvrés (Lun-Ven) de chaque intersection
 *   4. Retourner un Set<string> (dates YYYY-MM-DD, dédoublonnées)
 *
 * Exemple : absence 10-20 juillet, période analysée 15-30 juillet
 * → intersection = 15-20 juillet
 * → jours ouvrés retournés = {15, 16, 17, 18, 19, 20} filtrés Lun-Ven
 *
 * @param absences Liste complète des absences (toutes contrôleurs confondus)
 * @param controleurEmail Email du contrôleur (case-insensitive)
 * @param from Borne basse de la période analysée (Date locale, 00:00)
 * @param to Borne haute de la période analysée (Date locale, EOD)
 */
export function getAbsentDaysForControleur(
  absences: Absence[],
  controleurEmail: string,
  from: Date,
  to: Date,
): Set<string> {
  const email = controleurEmail.toLowerCase()
  const result = new Set<string>()

  // Bornes normalisées à minuit local pour comparer avec les dates
  // reconstruites depuis YYYY-MM-DD (également locales).
  const fromMs = new Date(from.getFullYear(), from.getMonth(), from.getDate()).getTime()
  const toMs = new Date(to.getFullYear(), to.getMonth(), to.getDate()).getTime()
  if (toMs < fromMs) return result

  for (const abs of absences) {
    if (abs.controleurEmail.toLowerCase() !== email) continue
    if (!abs.dateDebut || !abs.dateFin) continue

    // Parse YYYY-MM-DD en Date LOCALE (évite décalages timezone)
    const [dy, dm, dd] = abs.dateDebut.split('-').map(Number)
    const [fy, fm, fd] = abs.dateFin.split('-').map(Number)
    const debutMs = new Date(dy, dm - 1, dd).getTime()
    const finMs = new Date(fy, fm - 1, fd).getTime()
    if (finMs < debutMs) continue // absence mal saisie → skip

    // Intersection avec la période analysée
    const startMs = Math.max(debutMs, fromMs)
    const endMs = Math.min(finMs, toMs)
    if (endMs < startMs) continue // pas de chevauchement

    // Énumération des jours ouvrés dans l'intersection
    const cursor = new Date(startMs)
    const end = new Date(endMs)
    while (cursor.getTime() <= end.getTime()) {
      const day = cursor.getDay() // 0 = Dimanche, 6 = Samedi
      if (day !== 0 && day !== 6) {
        const y = cursor.getFullYear()
        const m = String(cursor.getMonth() + 1).padStart(2, '0')
        const d = String(cursor.getDate()).padStart(2, '0')
        result.add(`${y}-${m}-${d}`)
      }
      cursor.setDate(cursor.getDate() + 1)
    }
  }

  return result
}

/**
 * Construit une Map<emailLowerCase, Set<jour>> pour TOUS les contrôleurs
 * d'un coup — plus efficace que d'appeler getAbsentDaysForControleur
 * dans une boucle .map() côté React (parcourt les absences une seule fois).
 */
export function buildAbsentDaysMap(
  absences: Absence[],
  from: Date,
  to: Date,
): Map<string, Set<string>> {
  const map = new Map<string, Set<string>>()
  const emails = new Set(absences.map(a => a.controleurEmail.toLowerCase()).filter(Boolean))
  for (const email of emails) {
    map.set(email, getAbsentDaysForControleur(absences, email, from, to))
  }
  return map
}
