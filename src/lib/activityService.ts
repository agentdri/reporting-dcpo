/**
 * Service local de gestion des reportings contrôleur (Module 3).
 *
 * Implémentation actuelle : persistance localStorage.
 * La liste SharePoint ACTIVITE_Controleurs n'est pas encore générée comme
 * datasource Power Apps (cf. .power/schemas). Quand elle le sera, le contrat
 * de ce module reste compatible : les fonctions get/list/save/update peuvent
 * être réécrites pour appeler ActivityService généré sans modifier les
 * pages qui consomment ce module.
 */

export type ActivityStatus = 'Soumis' | 'Réalisé' | 'Reporté'

export interface ActivityLine {
  id: string
  domaine: string
  objet: string
  action: string
  duree: number
  observation?: string
}

export interface ActivityValidationNote {
  date: string
  manager: string
  managerEmail?: string
  decision: 'Réalisé' | 'Reporté'
  motif?: string
}

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

const STORAGE_KEY_REPORTS = 'reportingDCPO.activityReports.v1'
const STORAGE_KEY_DRAFT_PREFIX = 'reportingDCPO.activityDraft.v1.'

const READ_RAW = (key: string): string | null => {
  try { return window.localStorage.getItem(key) } catch { return null }
}

const WRITE_RAW = (key: string, value: string): void => {
  try { window.localStorage.setItem(key, value) } catch { /* ignore */ }
}

const REMOVE_RAW = (key: string): void => {
  try { window.localStorage.removeItem(key) } catch { /* ignore */ }
}

function uid(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}

function readAllReports(): ActivityReport[] {
  const raw = READ_RAW(STORAGE_KEY_REPORTS)
  if (!raw) return []
  try {
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) return parsed as ActivityReport[]
    return []
  } catch {
    return []
  }
}

function writeAllReports(reports: ActivityReport[]): void {
  WRITE_RAW(STORAGE_KEY_REPORTS, JSON.stringify(reports))
}

export function listReports(): ActivityReport[] {
  return readAllReports().sort((a, b) => (a.date < b.date ? 1 : -1))
}

export function getReport(id: string): ActivityReport | undefined {
  return readAllReports().find(r => r.id === id)
}

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

export function createReport(input: CreateReportInput): ActivityReport {
  const totals = computeTotals(input.lines)
  const now = new Date().toISOString()
  const report: ActivityReport = {
    id: uid(),
    controleurEmail: input.controleurEmail,
    controleurName: input.controleurName,
    date: input.date,
    lines: input.lines,
    totalHeures: totals.totalHeures,
    tempsOccupe: totals.tempsOccupe,
    statutJournee: totals.statutJournee,
    statut: 'Soumis',
    anomaliesDetectees: input.anomaliesDetectees,
    lienRapport: input.lienRapport,
    observationsGlobales: input.observationsGlobales,
    attachments: input.attachments,
    validationHistory: [],
    submittedAt: now,
    updatedAt: now,
  }
  const all = readAllReports()
  all.push(report)
  writeAllReports(all)
  return report
}

export function updateReport(id: string, patch: Partial<Omit<ActivityReport, 'id'>>): ActivityReport | undefined {
  const all = readAllReports()
  const idx = all.findIndex(r => r.id === id)
  if (idx === -1) return undefined
  const updated: ActivityReport = {
    ...all[idx],
    ...patch,
    updatedAt: new Date().toISOString(),
  }
  if (patch.lines) {
    const totals = computeTotals(patch.lines)
    updated.totalHeures = totals.totalHeures
    updated.tempsOccupe = totals.tempsOccupe
    updated.statutJournee = totals.statutJournee
  }
  all[idx] = updated
  writeAllReports(all)
  return updated
}

export interface ValidationInput {
  manager: string
  managerEmail?: string
  decision: 'Réalisé' | 'Reporté'
  motif?: string
}

export function validateReport(id: string, input: ValidationInput): ActivityReport | undefined {
  const all = readAllReports()
  const idx = all.findIndex(r => r.id === id)
  if (idx === -1) return undefined
  const note: ActivityValidationNote = {
    date: new Date().toISOString(),
    manager: input.manager,
    managerEmail: input.managerEmail,
    decision: input.decision,
    motif: input.motif,
  }
  const target = all[idx]
  const updated: ActivityReport = {
    ...target,
    statut: input.decision,
    validationHistory: [...(target.validationHistory ?? []), note],
    updatedAt: new Date().toISOString(),
  }
  all[idx] = updated
  writeAllReports(all)
  return updated
}

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
