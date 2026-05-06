import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'

export interface TicketAttachment {
  name: string
  url: string
  isImage: boolean
  iconType: 'image' | 'pdf' | 'word' | 'excel' | 'archive' | 'text' | 'file'
}

export const ATTACHMENT_API_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a8ced63bd1314a7897003b97eebd85e9/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EdDW19nkBl6pWigNfp0OQPaiSIjBwyhAuWeTO3s8b8E'

export interface UploadResponse {
  success: boolean
  sharePointId?: number
  attachmentUrl?: string
  message?: string
}

export function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      const result = reader.result as string
      const base64 = result.includes(',') ? result.split(',')[1] : result
      resolve(base64)
    }
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

export function getFileNameFromUrl(url: string): string {
  try {
    const pathname = new URL(url).pathname
    return decodeURIComponent(pathname.split('/').pop() ?? url)
  } catch {
    return url
  }
}

export function isImageUrl(url: string): boolean {
  return /\.(png|jpe?g|gif|bmp|webp|svg)(\?|$)/i.test(url)
}

export function getAttachmentIconType(name: string): TicketAttachment['iconType'] {
  const lower = name.toLowerCase()
  if (/\.(png|jpe?g|gif|bmp|webp|svg)(\?|$)/i.test(lower)) return 'image'
  if (lower.endsWith('.pdf')) return 'pdf'
  if (/\.(docx?|odt)$/i.test(lower)) return 'word'
  if (/\.(xlsx?|csv|ods)$/i.test(lower)) return 'excel'
  if (/\.(zip|rar|7z|tar|gz)$/i.test(lower)) return 'archive'
  if (/\.(txt|md|log)$/i.test(lower)) return 'text'
  return 'file'
}

export function getAttachmentIcon(type: TicketAttachment['iconType']): string {
  switch (type) {
    case 'image': return '🖼️'
    case 'pdf': return '📄'
    case 'word': return '📝'
    case 'excel': return '📊'
    case 'archive': return '🗜️'
    case 'text': return '📃'
    default: return '📎'
  }
}

export function getTicketAttachments(ticket: DCPO_LISTE_ANORMALIERead | null | undefined): TicketAttachment[] {
  if (!ticket) return []
  const out: TicketAttachment[] = []

  const sharepointAttachments = ticket['{Attachments}']
  if (Array.isArray(sharepointAttachments)) {
    sharepointAttachments.forEach(att => {
      if (!att?.AbsoluteUri) return
      const name = att.DisplayName ?? getFileNameFromUrl(att.AbsoluteUri)
      const iconType = getAttachmentIconType(name)
      out.push({
        name,
        url: att.AbsoluteUri,
        isImage: iconType === 'image',
        iconType,
      })
    })
  }

  if (ticket.urlPieceJointe) {
    const exists = out.some(a => a.url === ticket.urlPieceJointe)
    if (!exists) {
      const name = getFileNameFromUrl(ticket.urlPieceJointe)
      const iconType = getAttachmentIconType(name)
      out.push({
        name,
        url: ticket.urlPieceJointe,
        isImage: iconType === 'image',
        iconType,
      })
    }
  }

  return out
}

export async function uploadTicketAttachment(
  itemId: string,
  file: File,
  choise: string = 'Visite',
): Promise<string | undefined> {
  const fileContent = await fileToBase64(file)
  const response = await fetch(ATTACHMENT_API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      fileName: file.name,
      fileContent,
      Choise: choise,
      idItem: itemId,
    }),
  })
  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Echec upload piece jointe (${response.status}): ${errorText}`)
  }
  const data: UploadResponse = await response.json()
  return data.attachmentUrl
}
