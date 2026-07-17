/**
 * ============================================================================
 * SERVICE — RÉFÉRENTIEL DES DIRECTIONS
 * ============================================================================
 *
 * Persistance : liste SharePoint DCPO_LISTE_DIRECTION.
 *
 * Mapping des colonnes :
 *   ┌───────────┬───────────────┬─────────────────────────────────────────────┐
 *   │ Title     │ title (=sigle)│ Obligatoire SP — alimenté avec le sigle    │
 *   │ libelle   │ libelle       │ Libellé complet (ex: 'Direction des SI')   │
 *   │ sigle     │ sigle         │ Code court / acronyme (ex: 'DSI')           │
 *   └───────────┴───────────────┴─────────────────────────────────────────────┘
 *
 * Clé métier : `sigle` (DSI, DJC, DCE...). C'est le code court qui sert de
 * référence dans les autres listes (ex: PAC.directionsConcernees stocke des
 * sigles concaténés par ';').
 *
 * On synchronise Title ← sigle à l'écriture pour que la colonne par défaut
 * SharePoint affiche le sigle dans l'UI native (sinon Title resterait vide
 * et la ligne serait peu lisible côté SP).
 * ============================================================================
 */

import { DCPO_LISTE_DIRECTIONService } from '../generated/services/DCPO_LISTE_DIRECTIONService'
import type {
  DCPO_LISTE_DIRECTIONRead,
  DCPO_LISTE_DIRECTIONWrite,
} from '../generated/models/DCPO_LISTE_DIRECTIONModel'
import { getAllPages } from './sharePointPaging'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Représentation d'une direction côté front.
 *   - id : ID SharePoint (string pour homogénéiser avec le SDK)
 *   - title : copie du sigle (visibilité SP native)
 *   - libelle : libellé complet pour l'affichage
 *   - sigle : code court unique, sert de clé métier dans les références
 */
export interface Direction {
  id: string
  title: string
  libelle: string
  sigle: string
}

/** Données minimales pour créer une direction. */
export interface CreateDirectionInput {
  libelle: string
  sigle: string
}

/** Données partielles pour patcher une direction existante. */
export interface UpdateDirectionInput {
  libelle?: string
  sigle?: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — MAPPING SHAREPOINT ↔ DOMAINE
 * ────────────────────────────────────────────────────────────────────────── */

/** Construit une Direction depuis un item SharePoint. */
function fromItem(item: DCPO_LISTE_DIRECTIONRead): Direction {
  return {
    id: String(item.ID),
    title: item.Title ?? '',
    libelle: item.libelle ?? '',
    sigle: item.sigle ?? '',
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — API PUBLIQUE
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Liste les directions, triées par sigle (lecture humaine).
 *
 * On délègue le tri au client (pas d'orderBy serveur sur le sigle car il
 * n'est pas indexé et déclencherait un warning sur les listes étendues).
 */
export async function listDirections(): Promise<Direction[]> {
  try {
    const items = await getAllPages<DCPO_LISTE_DIRECTIONRead>(DCPO_LISTE_DIRECTIONService)
    return items.map(fromItem).sort((a, b) => a.sigle.localeCompare(b.sigle))
  } catch (err) {
    console.error('listDirections error', err)
    return []
  }
}

/** Récupère une direction par son ID SharePoint. */
export async function getDirection(id: string): Promise<Direction | undefined> {
  try {
    const res = await DCPO_LISTE_DIRECTIONService.get(id)
    if (!res.data) return undefined
    return fromItem(res.data)
  } catch (err) {
    console.error('getDirection error', err)
    return undefined
  }
}

/**
 * Crée une nouvelle direction.
 *
 * Validation minimale :
 *   - sigle obligatoire (clé métier)
 *   - libellé obligatoire
 *
 * Le champ Title SP est forcé à la valeur du sigle pour garder l'item lisible
 * dans l'UI native SharePoint (sinon Title serait vide).
 */
export async function createDirection(input: CreateDirectionInput): Promise<Direction> {
  const sigle = input.sigle.trim()
  const libelle = input.libelle.trim()
  if (!sigle) throw new Error('Le sigle est obligatoire.')
  if (!libelle) throw new Error('Le libellé est obligatoire.')

  const payload: Record<string, unknown> = {
    Title: sigle,
    sigle,
    libelle,
  }
  const res = await DCPO_LISTE_DIRECTIONService.create(
    payload as Omit<DCPO_LISTE_DIRECTIONWrite, 'ID'>,
  )
  if (!res.success || !res.data) {
    throw new Error(res.error?.message ?? 'Échec de la création de la direction.')
  }
  return fromItem(res.data)
}

/**
 * Met à jour une direction existante (sigle et/ou libellé).
 *
 * Si le sigle change, Title est aussi mis à jour pour rester synchrone.
 */
export async function updateDirection(id: string, patch: UpdateDirectionInput): Promise<Direction | undefined> {
  const payload: Record<string, unknown> = {}
  if (patch.sigle !== undefined) {
    const s = patch.sigle.trim()
    if (!s) throw new Error('Le sigle ne peut pas être vide.')
    payload.sigle = s
    // Title synchro avec sigle (cf. createDirection)
    payload.Title = s
  }
  if (patch.libelle !== undefined) {
    const l = patch.libelle.trim()
    if (!l) throw new Error('Le libellé ne peut pas être vide.')
    payload.libelle = l
  }
  if (Object.keys(payload).length === 0) return getDirection(id)

  try {
    await DCPO_LISTE_DIRECTIONService.update(
      id,
      payload as Partial<Omit<DCPO_LISTE_DIRECTIONWrite, 'ID'>>,
    )
  } catch (err) {
    console.error('updateDirection error', err)
    return undefined
  }
  return getDirection(id)
}

/** Supprime définitivement une direction. */
export async function deleteDirection(id: string): Promise<void> {
  await DCPO_LISTE_DIRECTIONService.delete(id)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — HELPERS UI
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Résout le libellé d'une direction à partir de son sigle, depuis une liste
 * pré-chargée. Fallback : renvoie le sigle lui-même si la direction n'est
 * pas trouvée (cas d'un sigle obsolète référencé par un ancien PAC).
 *
 * Note : on accepte la liste en paramètre plutôt qu'un appel SP à chaque
 * lookup, pour éviter N+1 requêtes côté front.
 */
export function getDirectionLabelFromList(directions: Direction[], sigle: string): string {
  return directions.find(d => d.sigle === sigle)?.libelle ?? sigle
}
