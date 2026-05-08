/**
 * ============================================================================
 * HELPERS DE LECTURE DES LISTES DE RÉFÉRENCE SHAREPOINT
 * ============================================================================
 *
 * Encapsule les appels aux services générés pour les listes "référentielles" :
 *   - DCPO_LISTE_AGENCES : agences bancaires (id, libellé, réseau parent)
 *   - DCPO_LISTE_RESEAUX : réseaux (régions, ex. OUEST, CENTRE)
 *   - DCPO_LISTE_USER    : annuaire local (email + rôle métier)
 *
 * Pourquoi cette couche ?
 *   - Centralise la gestion d'erreur (try/catch + log + retour [])
 *     → les pages consommatrices n'ont plus à se soucier des cas d'échec
 *   - Fournit des helpers de résolution ID → label pour l'affichage
 *   - Simplifie le code des pages : un seul `loadAllReferences()` au lieu
 *     de 3 appels parallèles avec gestion d'erreurs séparée
 * ============================================================================
 */

import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import { DCPO_LISTE_USERService } from '../generated/services/DCPO_LISTE_USERService'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import type { DCPO_LISTE_USERRead } from '../generated/models/DCPO_LISTE_USERModel'

/**
 * Bundle de tous les référentiels chargés en une fois.
 * Utilisé par les pages qui ont besoin de plusieurs listes simultanément
 * (typiquement Anomalies, Bulletins, Reporting Agent).
 */
export interface ReferenceRows {
  agences: DCPO_LISTE_AGENCESRead[]
  reseaux: DCPO_LISTE_RESEAUXRead[]
  users: DCPO_LISTE_USERRead[]
}

/**
 * Charge toutes les agences depuis SharePoint.
 *
 * Robustesse :
 *   - try/catch : en cas d'erreur réseau / SDK / permissions, retourne []
 *   - le ?? [] sécurise le cas où result.data serait undefined (succès vide)
 *
 * Retour : tableau d'agences (potentiellement vide en cas d'erreur).
 * L'appelant n'a JAMAIS à gérer les exceptions.
 */
export async function loadAgences(): Promise<DCPO_LISTE_AGENCESRead[]> {
  try {
    const r = await DCPO_LISTE_AGENCESService.getAll()
    return r.data ?? []
  } catch (err) {
    console.error('loadAgences error', err)
    return []
  }
}

/**
 * Charge tous les réseaux depuis SharePoint.
 * Mêmes garanties de robustesse que loadAgences.
 */
export async function loadReseaux(): Promise<DCPO_LISTE_RESEAUXRead[]> {
  try {
    const r = await DCPO_LISTE_RESEAUXService.getAll()
    return r.data ?? []
  } catch (err) {
    console.error('loadReseaux error', err)
    return []
  }
}

/**
 * Charge la liste des utilisateurs métier (DCPO_LISTE_USER).
 * Cette liste contient les emails et leurs rôles (fonction.Value), utilisée
 * pour le contrôle d'accès et la résolution rôle utilisateur.
 */
export async function loadUsers(): Promise<DCPO_LISTE_USERRead[]> {
  try {
    const r = await DCPO_LISTE_USERService.getAll()
    return r.data ?? []
  } catch (err) {
    console.error('loadUsers error', err)
    return []
  }
}

/**
 * Charge les 3 référentiels EN PARALLÈLE (Promise.all).
 *
 * Performance :
 *   Les 3 appels SharePoint partent simultanément → durée totale ≈ max des 3
 *   au lieu de la somme. Sur connexion lente, gain de 2-3 secondes au démarrage.
 *
 * Échecs partiels :
 *   Comme chaque load* gère ses propres erreurs et retourne [], un échec
 *   sur l'une des 3 listes n'empêche pas les 2 autres de s'afficher.
 */
export async function loadAllReferences(): Promise<ReferenceRows> {
  const [agences, reseaux, users] = await Promise.all([loadAgences(), loadReseaux(), loadUsers()])
  return { agences, reseaux, users }
}

/**
 * Trouve le LIBELLÉ d'une agence à partir d'une référence (ID ou Title).
 *
 * Utilisé pour l'affichage : la base stocke souvent l'ID en string,
 * mais l'utilisateur veut voir "First Bank Bessengue" pas "12".
 *
 * Stratégie de fallback :
 *   1. Si rawId est null/vide → chaîne vide
 *   2. Cherche par ID exact (cas standard)
 *   3. Sinon cherche par Title (au cas où la donnée stocke déjà le label)
 *   4. Si aucun match → retourne le rawId tel quel (mieux que rien)
 */
export function findAgenceLabel(agences: DCPO_LISTE_AGENCESRead[], rawId?: string | number | null): string {
  if (rawId === undefined || rawId === null || rawId === '') return ''
  const id = String(rawId)
  const match = agences.find(a => String(a.ID) === id || a.Title === id)
  return match?.Title ?? id
}

/**
 * Trouve le libellé d'un réseau (ex : "OUEST", "DOUALA-SUD").
 *
 * Le réseau est stocké dans `field_1` sur DCPO_LISTE_RESEAUX (et non Title).
 * Stratégie identique à findAgenceLabel : ID puis field_1 puis Title puis rawId.
 */
export function findReseauLabel(reseaux: DCPO_LISTE_RESEAUXRead[], rawId?: string | number | null): string {
  if (rawId === undefined || rawId === null || rawId === '') return ''
  const id = String(rawId)
  const match = reseaux.find(r => String(r.ID) === id || r.field_1 === id || r.Title === id)
  return match?.field_1 ?? match?.Title ?? id
}

/**
 * Trouve l'enregistrement utilisateur correspondant à un email donné.
 *
 * Comparaison case-insensitive (toLowerCase) — important car SharePoint
 * peut stocker les emails avec casse incohérente.
 *
 * Retour : l'objet user complet OU undefined si aucun match.
 */
export function findUserByEmail(users: DCPO_LISTE_USERRead[], email?: string | null): DCPO_LISTE_USERRead | undefined {
  if (!email) return undefined
  const target = email.toLowerCase()
  return users.find(u => u.Email?.toLowerCase() === target)
}

/**
 * Raccourci pour récupérer le rôle (fonction.Value) à partir d'un email.
 * Retourne undefined si l'utilisateur n'est pas trouvé OU n'a pas de fonction.
 *
 * Utilisé par la logique d'autorisation (ex. déterminer si l'user est manager).
 */
export function getUserRole(users: DCPO_LISTE_USERRead[], email?: string | null): string | undefined {
  return findUserByEmail(users, email)?.fonction?.Value
}

/**
 * Pour une agence donnée, retrouve son réseau parent via la relation
 * agence.field_1 → reseau.ID.
 *
 * Use case principal : au formulaire de création anomalie, quand l'user
 * sélectionne une agence, on auto-renseigne le réseau associé.
 *
 * Retourne undefined si :
 *   - agenceId vide
 *   - agence introuvable
 *   - agence sans réseau parent (field_1 vide)
 *   - réseau parent introuvable dans le référentiel chargé
 */
export function ticketReseauForAgence(
  agences: DCPO_LISTE_AGENCESRead[],
  reseaux: DCPO_LISTE_RESEAUXRead[],
  agenceId?: string | number | null,
): DCPO_LISTE_RESEAUXRead | undefined {
  if (!agenceId) return undefined
  const agence = agences.find(a => String(a.ID) === String(agenceId))
  if (!agence?.field_1) return undefined
  return reseaux.find(r => String(r.ID) === String(agence.field_1))
}
