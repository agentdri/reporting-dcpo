import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import { DCPO_LISTE_USERService } from '../generated/services/DCPO_LISTE_USERService'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import type { DCPO_LISTE_USERRead } from '../generated/models/DCPO_LISTE_USERModel'

export interface ReferenceRows {
  agences: DCPO_LISTE_AGENCESRead[]
  reseaux: DCPO_LISTE_RESEAUXRead[]
  users: DCPO_LISTE_USERRead[]
}

export async function loadAgences(): Promise<DCPO_LISTE_AGENCESRead[]> {
  try {
    const r = await DCPO_LISTE_AGENCESService.getAll()
    return r.data ?? []
  } catch (err) {
    console.error('loadAgences error', err)
    return []
  }
}

export async function loadReseaux(): Promise<DCPO_LISTE_RESEAUXRead[]> {
  try {
    const r = await DCPO_LISTE_RESEAUXService.getAll()
    return r.data ?? []
  } catch (err) {
    console.error('loadReseaux error', err)
    return []
  }
}

export async function loadUsers(): Promise<DCPO_LISTE_USERRead[]> {
  try {
    const r = await DCPO_LISTE_USERService.getAll()
    return r.data ?? []
  } catch (err) {
    console.error('loadUsers error', err)
    return []
  }
}

export async function loadAllReferences(): Promise<ReferenceRows> {
  const [agences, reseaux, users] = await Promise.all([loadAgences(), loadReseaux(), loadUsers()])
  return { agences, reseaux, users }
}

export function findAgenceLabel(agences: DCPO_LISTE_AGENCESRead[], rawId?: string | number | null): string {
  if (rawId === undefined || rawId === null || rawId === '') return ''
  const id = String(rawId)
  const match = agences.find(a => String(a.ID) === id || a.Title === id)
  return match?.Title ?? id
}

export function findReseauLabel(reseaux: DCPO_LISTE_RESEAUXRead[], rawId?: string | number | null): string {
  if (rawId === undefined || rawId === null || rawId === '') return ''
  const id = String(rawId)
  const match = reseaux.find(r => String(r.ID) === id || r.field_1 === id || r.Title === id)
  return match?.field_1 ?? match?.Title ?? id
}

export function findUserByEmail(users: DCPO_LISTE_USERRead[], email?: string | null): DCPO_LISTE_USERRead | undefined {
  if (!email) return undefined
  const target = email.toLowerCase()
  return users.find(u => u.Email?.toLowerCase() === target)
}

export function getUserRole(users: DCPO_LISTE_USERRead[], email?: string | null): string | undefined {
  return findUserByEmail(users, email)?.fonction?.Value
}

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
