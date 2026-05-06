import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import { findAgenceLabel, findReseauLabel } from './spReferenceRows'

export type BulletinStatus = 'Resolu' | 'Clos' | 'Tous'

export interface BulletinFilters {
  search: string
  status: BulletinStatus
  classification: string
  criticite: string
  agence: string
  closureFrom: string
  closureTo: string
}

export const EMPTY_BULLETIN_FILTERS: BulletinFilters = {
  search: '',
  status: 'Tous',
  classification: '',
  criticite: '',
  agence: '',
  closureFrom: '',
  closureTo: '',
}

export const BULLETIN_STATUS_VALUES = ['Resolu', 'Clos']

export interface LifecycleStep {
  date?: string
  label: string
  description?: string
  type: 'declaration' | 'opening' | 'regularization' | 'closure' | 'log' | 'event'
}

export interface ConsolidatedBulletin {
  ticket: DCPO_LISTE_ANORMALIERead
  numero: string
  titre: string
  agenceLabel: string
  reseauLabel: string
  declarantName: string
  auteurName: string
  affecteName: string
  declarationDate?: Date
  openingDate?: Date
  regularizationDate?: Date
  closureDate?: Date
  delayDays?: number
  classification: string
  criticite: string
  statut: string
  montant?: number
  description: string
  causeImmediate: string
  causeRacine: string
  observations: string
  actions: string[]
  domaine: string
  natureRisque: string
  occurrences?: number
  lifecycle: LifecycleStep[]
  sharePointUrl?: string
}

function stripHtml(html?: string | null): string {
  if (!html) return ''
  const doc = new DOMParser().parseFromString(html, 'text/html')
  return doc.body.textContent?.trim() ?? ''
}

function parseDateSafe(value?: string | null): Date | undefined {
  if (!value) return undefined
  const d = new Date(value)
  return Number.isNaN(d.getTime()) ? undefined : d
}

function diffInDays(start?: Date, end?: Date): number | undefined {
  if (!start || !end) return undefined
  const ms = end.getTime() - start.getTime()
  if (Number.isNaN(ms)) return undefined
  return Math.max(0, Math.round(ms / (1000 * 60 * 60 * 24)))
}

interface ExtendedTicket extends DCPO_LISTE_ANORMALIERead {
  field_11?: string
  field_12?: string
  field_13?: string
  field_14?: string
  field_15?: string
  field_16?: string
  field_17?: string
  field_18?: string
  field_19?: string
  field_20?: number
  field_21?: string
  field_22?: string
  field_23?: string
  natureRisque?: string
  domaineAnomalie?: string
  causeImmediate?: string
  causeRacine?: string
  actionsAMener?: string
  observationsBulletin?: string
}

function pickFirstString(...candidates: Array<string | undefined | null>): string {
  for (const c of candidates) {
    if (c && String(c).trim()) return String(c).trim()
  }
  return ''
}

function pickFirstHtml(...candidates: Array<string | undefined | null>): string {
  const raw = pickFirstString(...candidates)
  return raw ? stripHtml(raw) : ''
}

function splitActions(raw: string): string[] {
  if (!raw) return []
  return raw
    .split(/\r?\n|;|•|–|–/)
    .map(s => s.trim())
    .filter(Boolean)
}

export function buildLifecycleSteps(ticket: DCPO_LISTE_ANORMALIERead): LifecycleStep[] {
  const ext = ticket as ExtendedTicket
  const steps: LifecycleStep[] = []

  const declared = parseDateSafe(ticket.field_0) ?? parseDateSafe(ticket.Created)
  const opened = parseDateSafe(ticket.dateOuvertureTicket) ?? declared
  const regularized = parseDateSafe(ticket.field_9)
  const closed = parseDateSafe(ticket.date_cloture_ticket)

  if (declared) {
    steps.push({
      type: 'declaration',
      label: 'Déclaration',
      date: declared.toISOString(),
      description: ticket.declarant_anormalie?.DisplayName,
    })
  }
  if (opened && opened.getTime() !== declared?.getTime()) {
    steps.push({
      type: 'opening',
      label: 'Ouverture du ticket',
      date: opened.toISOString(),
    })
  }
  if (regularized) {
    steps.push({
      type: 'regularization',
      label: 'Régularisation',
      date: regularized.toISOString(),
      description: ticket.personneAffecter?.DisplayName,
    })
  }
  if (closed) {
    steps.push({
      type: 'closure',
      label: 'Clôture',
      date: closed.toISOString(),
    })
  }

  const journalRaw = ext.field_23
  if (journalRaw && typeof journalRaw === 'string') {
    const lines = stripHtml(journalRaw).split(/\r?\n+/).map(l => l.trim()).filter(Boolean)
    lines.forEach(line => {
      const dateMatch = line.match(/^\[?(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?)\]?\s*[:\-–]?\s*(.*)$/)
      if (dateMatch) {
        steps.push({
          type: 'log',
          label: dateMatch[2] || 'Note',
          date: dateMatch[1],
        })
      } else {
        steps.push({
          type: 'log',
          label: line,
        })
      }
    })
  }

  steps.sort((a, b) => {
    const da = a.date ? new Date(a.date).getTime() : 0
    const db = b.date ? new Date(b.date).getTime() : 0
    return da - db
  })

  return steps
}

export function buildConsolidatedBulletin(
  ticket: DCPO_LISTE_ANORMALIERead,
  agences: DCPO_LISTE_AGENCESRead[],
  reseaux: DCPO_LISTE_RESEAUXRead[],
): ConsolidatedBulletin {
  const ext = ticket as ExtendedTicket

  const declarationDate = parseDateSafe(ticket.field_0) ?? parseDateSafe(ticket.Created)
  const openingDate = parseDateSafe(ticket.dateOuvertureTicket) ?? declarationDate
  const regularizationDate = parseDateSafe(ticket.field_9)
  const closureDate = parseDateSafe(ticket.date_cloture_ticket)

  const delayDays = diffInDays(declarationDate, closureDate ?? regularizationDate)

  const description = pickFirstHtml(ticket.field_4, ticket.field_3, ticket.Title)
  const causeImmediate = pickFirstHtml(ext.causeImmediate, ext.field_11, ticket.field_3)
  const causeRacine = pickFirstHtml(ext.causeRacine, ext.field_12)
  const observations = pickFirstHtml(ext.observationsBulletin, ext.field_22)
  const actionsRaw = pickFirstHtml(ext.actionsAMener, ext.field_21)
  const actions = splitActions(actionsRaw)
  const domaine = pickFirstString(ext.domaineAnomalie, ext.field_13)
  const natureRisque = pickFirstString(ext.natureRisque, ext.field_14)
  const occurrences = ext.field_20 ?? undefined

  const sharePointUrl = ticket['{Link}'] ?? undefined

  const numero = ticket.ID ? `T-${ticket.ID}` : '—'
  const titre = pickFirstString(ticket.Title, ticket.field_3, description.substring(0, 80))

  return {
    ticket,
    numero,
    titre: titre || `Anomalie ${numero}`,
    agenceLabel: findAgenceLabel(agences, ticket.field_6),
    reseauLabel: findReseauLabel(reseaux, ticket.field_7),
    declarantName: ticket.declarant_anormalie?.DisplayName ?? '—',
    auteurName: ticket.auteur_anormalie?.DisplayName ?? '—',
    affecteName: ticket.personneAffecter?.DisplayName ?? '—',
    declarationDate,
    openingDate,
    regularizationDate,
    closureDate,
    delayDays,
    classification: ticket.field_5 ?? '',
    criticite: ticket.criticiteAnomalie ?? '',
    statut: ticket.field_10 ?? '',
    montant: ticket.field_8 ?? undefined,
    description,
    causeImmediate,
    causeRacine,
    observations,
    actions,
    domaine,
    natureRisque,
    occurrences,
    lifecycle: buildLifecycleSteps(ticket),
    sharePointUrl,
  }
}

export function isBulletinTicket(ticket: DCPO_LISTE_ANORMALIERead): boolean {
  if (!ticket.field_10) return false
  return BULLETIN_STATUS_VALUES.some(s => ticket.field_10 === s)
}

export function applyBulletinFilters(
  bulletins: ConsolidatedBulletin[],
  filters: BulletinFilters,
): ConsolidatedBulletin[] {
  const search = filters.search.trim().toLowerCase()
  const fromDate = filters.closureFrom ? new Date(`${filters.closureFrom}T00:00:00`) : undefined
  const toDate = filters.closureTo ? new Date(`${filters.closureTo}T23:59:59`) : undefined

  return bulletins.filter(b => {
    if (filters.status !== 'Tous' && b.statut !== filters.status) return false
    if (filters.classification && b.classification !== filters.classification) return false
    if (filters.criticite && b.criticite !== filters.criticite) return false
    if (filters.agence && String(b.ticket.field_6 ?? '') !== filters.agence) return false

    if (fromDate || toDate) {
      const ref = b.closureDate ?? b.regularizationDate
      if (!ref) return false
      if (fromDate && ref < fromDate) return false
      if (toDate && ref > toDate) return false
    }

    if (search) {
      const haystack = [
        b.numero,
        b.titre,
        b.description,
        b.declarantName,
        b.auteurName,
        b.affecteName,
        b.agenceLabel,
        b.reseauLabel,
        b.classification,
        b.criticite,
        b.causeImmediate,
        b.causeRacine,
        b.observations,
        ...b.actions,
      ]
        .filter(Boolean)
        .join(' ')
        .toLowerCase()
      if (!haystack.includes(search)) return false
    }
    return true
  })
}

export function formatDate(date?: Date | null): string {
  if (!date) return '—'
  return date.toLocaleDateString('fr-FR')
}

export function formatDateTime(date?: Date | string | null): string {
  if (!date) return ''
  const d = typeof date === 'string' ? new Date(date) : date
  if (Number.isNaN(d.getTime())) return ''
  return d.toLocaleString('fr-FR')
}

export function formatAmount(value?: number | null): string {
  if (value === null || value === undefined) return '—'
  return value.toLocaleString('fr-FR')
}

export function getCriticiteClass(criticite?: string | null): string {
  switch ((criticite ?? '').toLowerCase()) {
    case 'critique': return 'crit-critique'
    case 'haute': return 'crit-haute'
    case 'moyenne': return 'crit-moyenne'
    case 'faible': return 'crit-faible'
    default: return 'crit-default'
  }
}

export function getStatusClass(status?: string | null): string {
  switch ((status ?? '').toLowerCase()) {
    case 'clos': return 'status-clos'
    case 'resolu':
    case 'résolu': return 'status-resolu'
    case 'en cours': return 'status-en-cours'
    case 'ouvert': return 'status-ouvert'
    default: return 'status-default'
  }
}
