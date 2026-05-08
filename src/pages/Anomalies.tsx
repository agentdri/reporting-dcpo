import { useState, useEffect, useMemo } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { DCPO_LISTE_ANORMALIERead, DCPO_LISTE_ANORMALIEWrite } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import type { User } from '../generated/models/Office365UsersModel'

interface AnomaliesProps {
  userName?: string
  userEmail?: string
}

interface FormState {
  field_0: string
  field_4: string
  field_5: string
  field_6: string
  field_7: string
  field_8: number
  field_9: string
  field_10: string
  criticiteAnomalie: string
}

const EMPTY_FORM: FormState = {
  field_0: '',
  field_4: '',
  field_5: '',
  field_6: '',
  field_7: '',
  field_8: 0,
  field_9: '',
  field_10: '',
  criticiteAnomalie: '',
}

const CRITICITE_OPTIONS = ['Faible', 'Moyenne', 'Haute', 'Critique']

//const ATTACHMENT_API_URL = 'https://e78a17afcaf0e888989bbeca000173.f8.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/19d148d0144041f49ec16f59e318d7db/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=0Yeg1G81xgXL1XUkBAyoHpqdtpFq1PbZrJJPLmixw1M'
const ATTACHMENT_API_URL = 'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a8ced63bd1314a7897003b97eebd85e9/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EdDW19nkBl6pWigNfp0OQPaiSIjBwyhAuWeTO3s8b8E'
interface UploadResponse {
  success: boolean
  sharePointId?: number
  attachmentUrl?: string
  message?: string
}

function stripHtml(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  return doc.body.textContent?.trim() ?? ''
}

function toClaims(email: string) {
  return `i:0#.f|membership|${email}`
}

function fileToBase64(file: File): Promise<string> {
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

function escapeHtml(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function statusBadgeClass(status?: string | null): string {
  switch ((status ?? '').toLowerCase()) {
    case 'ouvert': return 'ticket-badge-ouvert'
    case 'en cours': return 'ticket-badge-en-cours'
    case 'resolu': return 'ticket-badge-resolu'
    case 'clos': return 'ticket-badge-clos'
    default: return 'ticket-badge-default'
  }
}

function statusIcon(status?: string | null): string {
  switch ((status ?? '').toLowerCase()) {
    case 'ouvert': return '⚠'
    case 'en cours': return '⏳'
    case 'resolu': return '✓'
    case 'clos': return '✔'
    default: return '•'
  }
}

export default function Anomalies({ userName, userEmail }: AnomaliesProps) {
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [loading, setLoading] = useState(true)
  const [showForm, setShowForm] = useState(false)
  const [form, setForm] = useState(EMPTY_FORM)
  const [submitting, setSubmitting] = useState(false)
  const [attachment, setAttachment] = useState<File | null>(null)

  // Filtres
  const [filterDateFrom, setFilterDateFrom] = useState('')
  const [filterDateTo, setFilterDateTo] = useState('')
  const [filterAgence, setFilterAgence] = useState('')
  const [filterReseau, setFilterReseau] = useState('')
  const [filterClassification, setFilterClassification] = useState('')
  const [filterCriticite, setFilterCriticite] = useState('')
  const [filterAgent, setFilterAgent] = useState('')
  const [filterAffecte, setFilterAffecte] = useState('')
  // Filtres appliqués (snapshot lors du clic Rechercher)
  const [appliedAgent, setAppliedAgent] = useState('')
  const [appliedAffecte, setAppliedAffecte] = useState('')

  // Detail modale
  const [detailItem, setDetailItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)

  // Affectation modale
  const [affectItemId, setAffectItemId] = useState<number | null>(null)
  const [affectSearch, setAffectSearch] = useState('')
  const [affectResults, setAffectResults] = useState<User[]>([])
  const [affectLoading, setAffectLoading] = useState(false)

  // Changement de statut
  const [statusItem, setStatusItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)
  const [statusValue, setStatusValue] = useState('')
  const [statusDateRegul, setStatusDateRegul] = useState('')
  const [statusDateCloture, setStatusDateCloture] = useState('')
  const [statusSaving, setStatusSaving] = useState(false)
  const [statusError, setStatusError] = useState<string | null>(null)

  // Ticket detail (clic sur le badge ticket)
  const [ticketItem, setTicketItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)

  // Modale "Clôture de la résolution" (formulaire contrôleur)
  const [resolutionItem, setResolutionItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)
  const [resolutionForm, setResolutionForm] = useState({
    statut: 'Resolu',
    dateRegul: new Date().toISOString().split('T')[0],
    dateCloture: new Date().toISOString().split('T')[0],
    causeImmediate: '',
    causeRacine: '',
    actionsMenees: '',
    observations: '',
  })
  const [resolutionSaving, setResolutionSaving] = useState(false)
  const [resolutionError, setResolutionError] = useState<string | null>(null)

  // Tableau : toggle colonnes intermédiaires
  const [expandedColumns, setExpandedColumns] = useState(false)

  // Agences & Reseaux
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])

  // Auteur search
  const [auteurSearch, setAuteurSearch] = useState('')
  const [auteurEmail, setAuteurEmail] = useState('')
  const [auteurResults, setAuteurResults] = useState<User[]>([])
  const [showAuteurDropdown, setShowAuteurDropdown] = useState(false)

  // Personne affectee search
  const [affecteFormSearch, setAffecteFormSearch] = useState('')
  const [affecteFormEmail, setAffecteFormEmail] = useState('')
  const [affecteFormResults, setAffecteFormResults] = useState<User[]>([])
  const [showAffecteFormDropdown, setShowAffecteFormDropdown] = useState(false)

  const searchAuteur = async (term: string) => {
    setAuteurSearch(term)
    if (term.length < 2) { setAuteurResults([]); setShowAuteurDropdown(false); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) { setAuteurResults(result.data); setShowAuteurDropdown(true) }
    } catch (err) { console.error('Erreur recherche auteur', err) }
  }

  const selectAuteur = (u: User) => {
    setAuteurEmail(u.Mail ?? '')
    setAuteurSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAuteurDropdown(false)
    setAuteurResults([])
  }

  const searchAffecteForm = async (term: string) => {
    setAffecteFormSearch(term)
    if (term.length < 2) { setAffecteFormResults([]); setShowAffecteFormDropdown(false); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) { setAffecteFormResults(result.data); setShowAffecteFormDropdown(true) }
    } catch (err) { console.error('Erreur recherche personne affectee', err) }
  }

  const selectAffecteForm = (u: User) => {
    setAffecteFormEmail(u.Mail ?? '')
    setAffecteFormSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAffecteFormDropdown(false)
    setAffecteFormResults([])
  }

  const searchAffectUser = async (term: string) => {
    setAffectSearch(term)
    if (term.length < 2) { setAffectResults([]); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) setAffectResults(result.data)
    } catch (err) { console.error('Erreur recherche affectation', err) }
  }

  const confirmAffect = async (u: User) => {
    if (!affectItemId || !u.Mail) return
    setAffectLoading(true)
    try {
      await DCPO_LISTE_ANORMALIEService.update(String(affectItemId), {
        personneAffecter: {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(u.Mail),
        } as never,
      })
      closeAffectModal()
      await fetchItems()
    } catch (err) { console.error('Erreur affectation', err) }
    finally { setAffectLoading(false) }
  }

  const closeAffectModal = () => {
    setAffectItemId(null)
    setAffectSearch('')
    setAffectResults([])
  }

  const openStatusModal = (item: DCPO_LISTE_ANORMALIERead) => {
    setStatusItem(item)
    setStatusValue(item.field_10 ?? '')
    const today = new Date().toISOString().split('T')[0]
    setStatusDateRegul(item.field_9 ? item.field_9.split('T')[0] : today)
    setStatusDateCloture(item.date_cloture_ticket ? item.date_cloture_ticket.split('T')[0] : today)
    setStatusError(null)
  }

  const closeStatusModal = () => {
    setStatusItem(null)
    setStatusValue('')
    setStatusDateRegul('')
    setStatusDateCloture('')
    setStatusError(null)
    setStatusSaving(false)
  }

  const saveStatus = async () => {
    if (!statusItem || !statusValue) {
      setStatusError('Veuillez choisir un statut.')
      return
    }
    setStatusSaving(true)
    setStatusError(null)
    try {
      const payload: Record<string, unknown> = { field_10: statusValue }
      const today = new Date().toISOString().split('T')[0]
      const closesTicket = statusValue === 'Resolu' || statusValue === 'Clos'
      if (closesTicket) {
        const regulDate = statusDateRegul || today
        payload.field_9 = `${regulDate}T00:00:00Z`
        // Le ticket est clos dès qu'il passe à Resolu ou Clos
        const clotureDate = statusDateCloture || today
        payload.date_cloture_ticket = `${clotureDate}T00:00:00Z`
      }
      const result = await DCPO_LISTE_ANORMALIEService.update(
        String(statusItem.ID),
        payload as Partial<Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>>,
      )
      if (!result.success) {
        setStatusError('Erreur lors de la mise à jour du statut.')
        return
      }
      closeStatusModal()
      await fetchItems()
    } catch (err) {
      console.error('Erreur changement statut', err)
      setStatusError('Erreur lors de la mise à jour du statut.')
    } finally {
      setStatusSaving(false)
    }
  }

  const openResolution = (item: DCPO_LISTE_ANORMALIERead) => {
    const today = new Date().toISOString().split('T')[0]
    const existingDescription = item.field_4 ? stripHtml(item.field_4) : ''
    setResolutionForm({
      statut: item.field_10 === 'Clos' ? 'Clos' : 'Resolu',
      dateRegul: item.field_9 ? item.field_9.split('T')[0] : today,
      dateCloture: item.date_cloture_ticket ? item.date_cloture_ticket.split('T')[0] : today,
      causeImmediate: existingDescription,
      causeRacine: '',
      actionsMenees: '',
      observations: '',
    })
    setResolutionError(null)
    setResolutionItem(item)
    setTicketItem(null)
  }

  const closeResolution = () => {
    setResolutionItem(null)
    setResolutionError(null)
    setResolutionSaving(false)
  }

  const updateResolutionForm = <K extends keyof typeof resolutionForm>(field: K, value: string) => {
    setResolutionForm(prev => ({ ...prev, [field]: value }))
  }

  const buildResolutionHtml = (): string => {
    const blocks: string[] = []
    if (resolutionForm.causeImmediate.trim()) {
      blocks.push(`<p><strong>Cause immédiate :</strong> ${escapeHtml(resolutionForm.causeImmediate)}</p>`)
    }
    if (resolutionForm.causeRacine.trim()) {
      blocks.push(`<p><strong>Cause racine :</strong> ${escapeHtml(resolutionForm.causeRacine)}</p>`)
    }
    if (resolutionForm.actionsMenees.trim()) {
      const items = resolutionForm.actionsMenees
        .split(/\r?\n/)
        .map(s => s.trim())
        .filter(Boolean)
      if (items.length > 1) {
        blocks.push(`<p><strong>Actions menées :</strong></p><ul>${items.map(i => `<li>${escapeHtml(i)}</li>`).join('')}</ul>`)
      } else if (items.length === 1) {
        blocks.push(`<p><strong>Actions menées :</strong> ${escapeHtml(items[0])}</p>`)
      }
    }
    if (resolutionForm.observations.trim()) {
      blocks.push(`<p><strong>Observations :</strong> ${escapeHtml(resolutionForm.observations)}</p>`)
    }
    return blocks.join('')
  }

  const saveResolution = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!resolutionItem) return
    if (!resolutionForm.actionsMenees.trim()) {
      setResolutionError('Les actions menées sont obligatoires pour clôturer une résolution.')
      return
    }
    setResolutionSaving(true)
    setResolutionError(null)
    try {
      const today = new Date().toISOString().split('T')[0]
      const html = buildResolutionHtml()
      // La modale "Clore la résolution" ferme toujours le ticket,
      // que le statut final soit Resolu ou Clos.
      const payload: Record<string, unknown> = {
        field_10: resolutionForm.statut,
        field_9: `${resolutionForm.dateRegul || today}T00:00:00Z`,
        date_cloture_ticket: `${resolutionForm.dateCloture || today}T00:00:00Z`,
      }
      if (html) payload.field_4 = html
      const result = await DCPO_LISTE_ANORMALIEService.update(
        String(resolutionItem.ID),
        payload as Partial<Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>>,
      )
      if (!result.success) {
        setResolutionError('Erreur lors de l\'enregistrement de la résolution.')
        return
      }
      closeResolution()
      await fetchItems()
    } catch (err) {
      console.error('Erreur résolution', err)
      setResolutionError('Erreur lors de l\'enregistrement de la résolution.')
    } finally {
      setResolutionSaving(false)
    }
  }

  const buildFilter = () => {
    const clauses: string[] = []
    if (filterDateFrom) {
      clauses.push(`Created ge '${filterDateFrom}T00:00:00Z'`)
    }
    if (filterDateTo) {
      clauses.push(`Created le '${filterDateTo}T23:59:59Z'`)
    }
    if (filterAgence) clauses.push(`field_6 eq '${filterAgence}'`)
    if (filterReseau) clauses.push(`field_7 eq '${filterReseau}'`)
    if (filterClassification) clauses.push(`field_5 eq '${filterClassification}'`)
    if (filterCriticite) clauses.push(`criticiteAnomalie eq '${filterCriticite}'`)
    return clauses.join(' and ')
  }

  const fetchItems = async () => {
    setLoading(true)
    try {
      const filter = buildFilter()
      const options: { orderBy: string[]; filter?: string } = {
        orderBy: ['Created desc'],
      }
      if (filter) options.filter = filter

      const result = await DCPO_LISTE_ANORMALIEService.getAll(options)
      if (result.data) setItems(result.data)
    } catch (err) { console.error('Erreur chargement anomalies', err) }
    finally { setLoading(false) }
  }

  useEffect(() => {
    const loadLists = async () => {
      try {
        const [agencesRes, reseauxRes] = await Promise.all([
          DCPO_LISTE_AGENCESService.getAll(),
          DCPO_LISTE_RESEAUXService.getAll(),
        ])
        if (agencesRes.data) setAgences(agencesRes.data)
        if (reseauxRes.data) setReseaux(reseauxRes.data)
      } catch (err) { console.error('Erreur chargement agences/reseaux', err) }
    }
    loadLists()
    fetchItems()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const handleSearch = () => {
    setAppliedAgent(filterAgent)
    setAppliedAffecte(filterAffecte)
    fetchItems()
  }

  const handleResetFilters = () => {
    setFilterDateFrom('')
    setFilterDateTo('')
    setFilterAgence('')
    setFilterReseau('')
    setFilterClassification('')
    setFilterCriticite('')
    setFilterAgent('')
    setFilterAffecte('')
    setAppliedAgent('')
    setAppliedAffecte('')
    setTimeout(() => fetchItems(), 0)
  }

  // Filtrage client pour Agent (auteur) et Personne affectée — appliqué uniquement au clic
  const filteredItems = useMemo(() => {
    const agentTerm = appliedAgent.trim().toLowerCase()
    const affecteTerm = appliedAffecte.trim().toLowerCase()
    if (!agentTerm && !affecteTerm) return items
    return items.filter(it => {
      if (agentTerm) {
        const haystack = `${it.auteur_anormalie?.DisplayName ?? ''} ${it.auteur_anormalie?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(agentTerm)) return false
      }
      if (affecteTerm) {
        const haystack = `${it.personneAffecter?.DisplayName ?? ''} ${it.personneAffecter?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(affecteTerm)) return false
      }
      return true
    })
  }, [items, appliedAgent, appliedAffecte])

  const handleChange = (field: string, value: string | number) => {
    setForm(prev => ({ ...prev, [field]: value }))
  }

  const resetForm = () => {
    setForm(EMPTY_FORM)
    setAuteurSearch(''); setAuteurEmail('')
    setAffecteFormSearch(''); setAffecteFormEmail('')
    setAttachment(null)
  }

  const getFileNameFromUrl = (url: string) => {
    try {
      const pathname = new URL(url).pathname
      return decodeURIComponent(pathname.split('/').pop() ?? url)
    } catch {
      return url
    }
  }

  const isImageUrl = (url: string) => /\.(png|jpe?g|gif|bmp|webp|svg)(\?|$)/i.test(url)

  const uploadAttachment = async (itemId: string, file: File): Promise<string | undefined> => {
    const fileContent = await fileToBase64(file)
    const response = await fetch(ATTACHMENT_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        fileName: file.name,
        fileContent,
        Choise: 'Visite',
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

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setSubmitting(true)
    try {
      const payload: Record<string, unknown> = {}
      if (form.field_0) payload.field_0 = form.field_0 + 'T00:00:00Z'
      if (form.field_4) payload.field_4 = form.field_4
      if (form.field_5) payload.field_5 = form.field_5
      if (form.field_6) payload.field_6 = form.field_6
      if (form.field_7) payload.field_7 = form.field_7
      if (form.field_8) payload.field_8 = form.field_8
      if (form.field_9) payload.field_9 = form.field_9 + 'T00:00:00Z'
      if (form.field_10) payload.field_10 = form.field_10
      else payload.field_10 = 'Ouvert'
      if (form.criticiteAnomalie) payload.criticiteAnomalie = form.criticiteAnomalie

      // Ouverture automatique du ticket
      payload.dateOuvertureTicket = new Date().toISOString()

      if (auteurEmail) {
        payload['auteur_anormalie'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(auteurEmail),
        }
      }
      if (userEmail) {
        payload['declarant_anormalie'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(userEmail),
        }
      }
      if (affecteFormEmail) {
        payload['personneAffecter'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(affecteFormEmail),
        }
      }

      const result = await DCPO_LISTE_ANORMALIEService.create(payload as Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>)
      if (!result.success) { console.error('Erreur SharePoint:', result.error); return }

      if (attachment && result.data?.ID) {
        try {
          const attachmentUrl = await uploadAttachment(String(result.data.ID), attachment)
          if (attachmentUrl) {
            await DCPO_LISTE_ANORMALIEService.update(String(result.data.ID), {
              urlPieceJointe: attachmentUrl,
            })
          }
        } catch (err) {
          console.error('Erreur upload piece jointe:', err)
        }
      }

      resetForm()
      setShowForm(false)
      await fetchItems()
    } catch (err) { console.error('Erreur creation:', err) }
    finally { setSubmitting(false) }
  }

  return (
    <>
      <div className="content-header">
        <h2>Liste des anomalies</h2>
        <button
          type="button"
          className={`btn-cta btn-cta-primary btn-cta-add ${showForm ? 'is-active' : ''}`}
          onClick={() => setShowForm(!showForm)}
          aria-expanded={showForm}
        >
          {showForm ? '× Annuler' : '+ Nouvelle anomalie'}
        </button>
      </div>

      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{filteredItems.length}</span>
          <span className="stat-label">Total</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'Ouvert').length}</span>
          <span className="stat-label">Ouvert</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'En cours').length}</span>
          <span className="stat-label">En cours</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'Resolu').length}</span>
          <span className="stat-label">Resolu</span>
        </div>
        <div className="stat-card clos">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'Clos').length}</span>
          <span className="stat-label">Clos</span>
        </div>
        <div className="stat-card montant">
          <span className="stat-value">{filteredItems.reduce((s, i) => s + (i.field_8 ?? 0), 0).toLocaleString()}</span>
          <span className="stat-label">Montant total</span>
        </div>
      </div>

      <div className="filters-bar">
        <div className="filter-field">
          <label>Date du</label>
          <input type="date" value={filterDateFrom} onChange={e => setFilterDateFrom(e.target.value)} />
        </div>
        <div className="filter-field">
          <label>Date au</label>
          <input type="date" value={filterDateTo} onChange={e => setFilterDateTo(e.target.value)} />
        </div>
        <div className="filter-field">
          <label>Agence</label>
          <select value={filterAgence} onChange={e => setFilterAgence(e.target.value)}>
            <option value="">Toutes</option>
            {agences.map(a => (
              <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Reseau</label>
          <select value={filterReseau} onChange={e => setFilterReseau(e.target.value)}>
            <option value="">Tous</option>
            {reseaux.map(r => (
              <option key={r.ID} value={String(r.ID)}>{r.field_1 ?? r.Title}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Classification</label>
          <select value={filterClassification} onChange={e => setFilterClassification(e.target.value)}>
            <option value="">Toutes</option>
            <option value="Operationnel">Operationnel</option>
            <option value="Fraude">Fraude</option>
            <option value="Commercial">Commercial</option>
          </select>
        </div>
        <div className="filter-field">
          <label>Criticite</label>
          <select value={filterCriticite} onChange={e => setFilterCriticite(e.target.value)}>
            <option value="">Toutes</option>
            {CRITICITE_OPTIONS.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Agent (auteur)</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filterAgent}
            onChange={e => setFilterAgent(e.target.value)}
          />
        </div>
        <div className="filter-field">
          <label>Personne affectée</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filterAffecte}
            onChange={e => setFilterAffecte(e.target.value)}
          />
        </div>
        <button className="btn-search-filters" onClick={handleSearch} disabled={loading}>
          {loading ? 'Recherche...' : 'Rechercher'}
        </button>
        <button className="btn-reset-filters" onClick={handleResetFilters}>Reinitialiser</button>
      </div>

      {showForm && (
        <form className="create-form modern-form" onSubmit={handleSubmit}>
          <header className="modern-form-header">
            <div>
              <h2>Nouvelle anomalie</h2>
              <p className="modern-form-subtitle">
                Renseignez les informations ci-dessous. Un ticket de suivi sera automatiquement ouvert.
              </p>
            </div>
            <button type="button" className="btn-cta btn-cta-ghost" onClick={() => setShowForm(false)}>
              Annuler
            </button>
          </header>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">👤</span> Identité</legend>
            <div className="form-grid">
              <div className="form-field">
                <label>Déclarant</label>
                <input type="text" readOnly value={userName ?? ''} />
                {userEmail && <span className="selected-email">{userEmail}</span>}
              </div>

              <div className="form-field">
                <label htmlFor="anom-auteur">Auteur</label>
                <div className="autocomplete-wrapper">
                  <input
                    id="anom-auteur"
                    type="text"
                    value={auteurSearch}
                    placeholder="Rechercher un auteur..."
                    onChange={e => searchAuteur(e.target.value)}
                    onBlur={() => setTimeout(() => setShowAuteurDropdown(false), 200)}
                  />
                  {auteurEmail && <span className="selected-email">{auteurEmail}</span>}
                  {showAuteurDropdown && auteurResults.length > 0 && (
                    <ul className="autocomplete-dropdown">
                      {auteurResults.map(u => (
                        <li key={u.Id} onClick={() => selectAuteur(u)}>
                          <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>

              <div className="form-field">
                <label htmlFor="anom-affecte">Personne affectée</label>
                <div className="autocomplete-wrapper">
                  <input
                    id="anom-affecte"
                    type="text"
                    value={affecteFormSearch}
                    placeholder="Rechercher une personne..."
                    onChange={e => searchAffecteForm(e.target.value)}
                    onBlur={() => setTimeout(() => setShowAffecteFormDropdown(false), 200)}
                  />
                  {affecteFormEmail && <span className="selected-email">{affecteFormEmail}</span>}
                  {showAffecteFormDropdown && affecteFormResults.length > 0 && (
                    <ul className="autocomplete-dropdown">
                      {affecteFormResults.map(u => (
                        <li key={u.Id} onClick={() => selectAffecteForm(u)}>
                          <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">📋</span> Caractérisation</legend>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="anom-date">Date de l'anomalie</label>
                <input
                  id="anom-date"
                  type="date"
                  value={form.field_0}
                  max={new Date().toISOString().split('T')[0]}
                  onChange={e => handleChange('field_0', e.target.value)}
                />
              </div>
              <div className="form-field">
                <label htmlFor="anom-classification">Classification</label>
                <select
                  id="anom-classification"
                  value={form.field_5}
                  onChange={e => handleChange('field_5', e.target.value)}
                >
                  <option value="">— Choisir —</option>
                  <option value="Operationnel">Opérationnel</option>
                  <option value="Fraude">Fraude</option>
                  <option value="Commercial">Commercial</option>
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="anom-criticite">Criticité</label>
                <select
                  id="anom-criticite"
                  value={form.criticiteAnomalie}
                  onChange={e => handleChange('criticiteAnomalie', e.target.value)}
                >
                  <option value="">— Choisir —</option>
                  {CRITICITE_OPTIONS.map(opt => (
                    <option key={opt} value={opt}>{opt}</option>
                  ))}
                </select>
              </div>
              <div className="form-field form-field-full">
                <label htmlFor="anom-cause">Cause / description</label>
                <textarea
                  id="anom-cause"
                  rows={3}
                  value={form.field_4}
                  onChange={e => handleChange('field_4', e.target.value)}
                  placeholder="Décrivez la nature de l'anomalie observée..."
                />
              </div>
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">📍</span> Localisation</legend>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="anom-agence">Agence</label>
                <select
                  id="anom-agence"
                  value={form.field_6}
                  onChange={e => {
                    const agenceId = e.target.value
                    handleChange('field_6', agenceId)
                    const agence = agences.find(a => String(a.ID) === agenceId)
                    handleChange('field_7', agence?.field_1 ? String(agence.field_1) : '')
                  }}
                >
                  <option value="">— Choisir une agence —</option>
                  {agences.map(a => (
                    <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="anom-reseau">Réseau</label>
                <input
                  id="anom-reseau"
                  type="text"
                  readOnly
                  value={reseaux.find(r => String(r.ID) === form.field_7)?.field_1 ?? ''}
                  placeholder="Sélectionner une agence d'abord"
                />
              </div>
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">💰</span> Impact &amp; suivi</legend>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="anom-montant">Montant</label>
                <input
                  id="anom-montant"
                  type="number"
                  min={0}
                  value={form.field_8}
                  onChange={e => handleChange('field_8', Number(e.target.value))}
                />
              </div>
              <div className="form-field">
                <label htmlFor="anom-statut">Statut initial</label>
                <select
                  id="anom-statut"
                  value={form.field_10}
                  onChange={e => handleChange('field_10', e.target.value)}
                >
                  <option value="">— Choisir (Ouvert par défaut) —</option>
                  <option value="Ouvert">Ouvert</option>
                  <option value="En cours">En cours</option>
                  <option value="Resolu">Résolu</option>
                  <option value="Clos">Clos</option>
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="anom-dateRegul">Date régularisation</label>
                <input
                  id="anom-dateRegul"
                  type="date"
                  value={form.field_9}
                  onChange={e => handleChange('field_9', e.target.value)}
                />
              </div>
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">📎</span> Pièce jointe</legend>
            <div className="form-grid">
              <div className="form-field form-field-full">
                <label htmlFor="anom-attachment">Fichier (optionnel)</label>
                <input
                  id="anom-attachment"
                  type="file"
                  onChange={e => setAttachment(e.target.files?.[0] ?? null)}
                />
                {attachment && <span className="selected-email">📎 {attachment.name}</span>}
              </div>
            </div>
          </fieldset>

          <div className="modern-form-actions">
            <button type="button" className="btn-cta btn-cta-ghost" onClick={() => setShowForm(false)} disabled={submitting}>
              Annuler
            </button>
            <button className="btn-cta btn-cta-primary" type="submit" disabled={submitting}>
              {submitting ? 'Enregistrement...' : 'Créer l\'anomalie'}
            </button>
          </div>
        </form>
      )}

      {loading ? (
        <p className="loading-text">Chargement des anomalies...</p>
      ) : filteredItems.length === 0 ? (
        <p className="loading-text">Aucune anomalie trouvée.</p>
      ) : (
        <div className="table-wrapper">
          <table className={`data-table anomalies-table ${expandedColumns ? 'is-expanded' : 'is-compact'}`}>
            <thead>
              <tr>
                <th className="col-ticket">Numéro</th>
                <th className="col-status">Statut</th>
                <th className="col-date">Date</th>
                <th className="col-criticite">Criticité</th>
                <th className="col-toggle" aria-label="Déplier / replier les colonnes">
                  <button
                    type="button"
                    className="col-toggle-btn"
                    onClick={() => setExpandedColumns(v => !v)}
                    aria-pressed={expandedColumns}
                    aria-label={expandedColumns ? 'Replier les colonnes' : 'Déplier les colonnes'}
                    title={expandedColumns ? 'Replier les colonnes' : 'Déplier les colonnes'}
                  >
                    {expandedColumns ? '−' : '+'}
                  </button>
                </th>
                {expandedColumns && (
                  <>
                    <th className="col-collapsible">Déclarant</th>
                    <th className="col-collapsible">Auteur</th>
                    <th className="col-collapsible">Personne affectée</th>
                    <th className="col-collapsible">Cause</th>
                    <th className="col-collapsible">Classification</th>
                    <th className="col-collapsible">Agence</th>
                    <th className="col-collapsible">Réseau</th>
                    <th className="col-collapsible">Montant</th>
                    <th className="col-collapsible">Date régularisation</th>
                  </>
                )}
                <th className="col-actions">Actions</th>
              </tr>
            </thead>
            <tbody>
              {filteredItems.map((item, index) => (
                <tr key={item.ID}>
                  <td className="col-ticket">
                    <span
                      className={`ticket-badge ${statusBadgeClass(item.field_10)}`}
                      aria-label={`Ticket #${index + 1}, statut ${item.field_10 ?? 'inconnu'}`}
                    >
                      <span className="ticket-badge-icon" aria-hidden="true">{statusIcon(item.field_10)}</span>
                      <span className="ticket-badge-num">T-{index + 1}</span>
                    </span>
                  </td>
                  <td className="col-status">
                    <span className={`status-chip ${statusBadgeClass(item.field_10)}`}>
                      {item.field_10 ?? '—'}
                    </span>
                  </td>
                  <td className="col-date">{item.field_0 ? new Date(item.field_0).toLocaleDateString('fr-FR') : '—'}</td>
                  <td className="col-criticite">
                    {item.criticiteAnomalie
                      ? <span className={`criticite-chip crit-${item.criticiteAnomalie.toLowerCase()}`}>{item.criticiteAnomalie}</span>
                      : '—'}
                  </td>
                  <td className="col-toggle" aria-hidden="true"></td>
                  {expandedColumns && (
                    <>
                      <td className="col-collapsible">{item.declarant_anormalie?.DisplayName ?? '—'}</td>
                      <td className="col-collapsible">{item.auteur_anormalie?.DisplayName ?? '—'}</td>
                      <td className="col-collapsible">{item.personneAffecter?.DisplayName ?? '—'}</td>
                      <td className="col-collapsible">{item.field_4 ? stripHtml(item.field_4) : '—'}</td>
                      <td className="col-collapsible">{item.field_5 ?? '—'}</td>
                      <td className="col-collapsible">{agences.find(a => String(a.ID) === item.field_6)?.Title ?? item.field_6 ?? '—'}</td>
                      <td className="col-collapsible">{reseaux.find(r => String(r.ID) === item.field_7)?.field_1 ?? item.field_7 ?? '—'}</td>
                      <td className="col-collapsible">{item.field_8?.toLocaleString('fr-FR') ?? '—'}</td>
                      <td className="col-collapsible">{item.field_9 ? new Date(item.field_9).toLocaleDateString('fr-FR') : '—'}</td>
                    </>
                  )}
                  <td className="col-actions">
                    <div className="actions-col">
                      <button type="button" className="btn-cta btn-cta-detail" onClick={() => setDetailItem(item)}>
                        Détail
                      </button>
                      <button type="button" className="btn-cta btn-cta-affect" onClick={() => setAffectItemId(item.ID ?? null)}>
                        Affecter
                      </button>
                      <button type="button" className="btn-cta btn-cta-status" onClick={() => setTicketItem(item)}>
                        Ticket
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {detailItem && (
        <div className="modal-overlay" onClick={() => setDetailItem(null)}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 560 }}>
            <div className="modal-header">
              <h2>Detail de l'anomalie</h2>
              <button className="modal-close" onClick={() => setDetailItem(null)}>&times;</button>
            </div>
            <div className="modal-body">
              <dl className="detail-grid">
                <dt>N° Ticket</dt><dd><strong>{detailItem.ID ? `T-${detailItem.ID}` : '-'}</strong></dd>
                <dt>Ouverture ticket</dt><dd>{detailItem.dateOuvertureTicket ? new Date(detailItem.dateOuvertureTicket).toLocaleString() : '-'}</dd>
                <dt>Cloture ticket</dt><dd>{detailItem.date_cloture_ticket ? new Date(detailItem.date_cloture_ticket).toLocaleString() : '-'}</dd>
                <dt>Declarant</dt><dd>{detailItem.declarant_anormalie?.DisplayName ?? '-'}</dd>
                <dt>Auteur</dt><dd>{detailItem.auteur_anormalie?.DisplayName ?? '-'}</dd>
                <dt>Personne affectee</dt><dd>{detailItem.personneAffecter?.DisplayName ?? '-'}</dd>
                <dt>Date</dt><dd>{detailItem.field_0 ? new Date(detailItem.field_0).toLocaleDateString() : '-'}</dd>
                <dt>Cause</dt><dd>{detailItem.field_4 ? stripHtml(detailItem.field_4) : '-'}</dd>
                <dt>Classification</dt><dd>{detailItem.field_5 ?? '-'}</dd>
                <dt>Agence</dt><dd>{agences.find(a => String(a.ID) === detailItem.field_6)?.Title ?? detailItem.field_6 ?? '-'}</dd>
                <dt>Reseau</dt><dd>{reseaux.find(r => String(r.ID) === detailItem.field_7)?.field_1 ?? detailItem.field_7 ?? '-'}</dd>
                <dt>Montant</dt><dd>{detailItem.field_8?.toLocaleString() ?? '-'}</dd>
                <dt>Date regularisation</dt><dd>{detailItem.field_9 ? new Date(detailItem.field_9).toLocaleDateString() : '-'}</dd>
                <dt>Statut</dt><dd>{detailItem.field_10 ?? '-'}</dd>
                <dt>Criticite</dt><dd>{detailItem.criticiteAnomalie ?? '-'}</dd>
              </dl>

              <div className="detail-attachments" style={{ marginTop: 16 }}>
                <h3 style={{ marginBottom: 8 }}>Piece jointe</h3>
                {detailItem.urlPieceJointe ? (
                  <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
                    {isImageUrl(detailItem.urlPieceJointe) && (
                      <img src={detailItem.urlPieceJointe} alt={getFileNameFromUrl(detailItem.urlPieceJointe)}
                        style={{ width: 80, height: 80, objectFit: 'cover', borderRadius: 4, border: '1px solid #ddd' }} />
                    )}
                    <a href={detailItem.urlPieceJointe} target="_blank" rel="noreferrer">
                      {getFileNameFromUrl(detailItem.urlPieceJointe)}
                    </a>
                  </div>
                ) : (
                  <p className="loading-text">Aucune piece jointe.</p>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {affectItemId !== null && (
        <div className="modal-overlay" onClick={closeAffectModal}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Affecter l'anomalie</h2>
              <button className="modal-close" onClick={closeAffectModal}>&times;</button>
            </div>
            <div className="modal-body">
              <input type="text" value={affectSearch} placeholder="Rechercher un utilisateur..."
                onChange={e => searchAffectUser(e.target.value)} autoFocus />
              {affectResults.length > 0 && (
                <ul className="modal-results">
                  {affectResults.map(u => (
                    <li key={u.Id} onClick={() => confirmAffect(u)}>
                      <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                    </li>
                  ))}
                </ul>
              )}
              {affectLoading && <p className="loading-text">Affectation en cours...</p>}
            </div>
          </div>
        </div>
      )}

      {ticketItem !== null && (
        <div className="modal-overlay" onClick={() => setTicketItem(null)}>
          <div className="modal ticket-modal" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="ticket-modal-title">
                <span className={`ticket-badge ticket-badge-lg ${statusBadgeClass(ticketItem.field_10)}`}>
                  <span className="ticket-badge-icon" aria-hidden="true">{statusIcon(ticketItem.field_10)}</span>
                  <span className="ticket-badge-num">T-{ticketItem.ID}</span>
                </span>
                <h2>Suivi du ticket</h2>
              </div>
              <button className="modal-close" onClick={() => setTicketItem(null)} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body ticket-modal-body">
              <div className="ticket-stat-row">
                <span className={`status-chip ${statusBadgeClass(ticketItem.field_10)}`}>
                  {ticketItem.field_10 ?? '—'}
                </span>
                {ticketItem.criticiteAnomalie && (
                  <span className={`criticite-chip crit-${ticketItem.criticiteAnomalie.toLowerCase()}`}>
                    {ticketItem.criticiteAnomalie}
                  </span>
                )}
                {ticketItem.field_5 && <span className="ticket-tag">{ticketItem.field_5}</span>}
              </div>

              <dl className="detail-grid">
                <dt>Ouverture</dt>
                <dd>{ticketItem.dateOuvertureTicket ? new Date(ticketItem.dateOuvertureTicket).toLocaleString('fr-FR') : '—'}</dd>
                <dt>Date déclaration</dt>
                <dd>{ticketItem.field_0 ? new Date(ticketItem.field_0).toLocaleDateString('fr-FR') : '—'}</dd>
                <dt>Régularisation</dt>
                <dd>{ticketItem.field_9 ? new Date(ticketItem.field_9).toLocaleDateString('fr-FR') : '—'}</dd>
                <dt>Clôture</dt>
                <dd>{ticketItem.date_cloture_ticket ? new Date(ticketItem.date_cloture_ticket).toLocaleString('fr-FR') : '—'}</dd>
                <dt>Déclarant</dt>
                <dd>{ticketItem.declarant_anormalie?.DisplayName ?? '—'}</dd>
                <dt>Auteur</dt>
                <dd>{ticketItem.auteur_anormalie?.DisplayName ?? '—'}</dd>
                <dt>Affecté à</dt>
                <dd>{ticketItem.personneAffecter?.DisplayName ?? '—'}</dd>
                <dt>Agence</dt>
                <dd>{agences.find(a => String(a.ID) === ticketItem.field_6)?.Title ?? '—'}</dd>
                <dt>Réseau</dt>
                <dd>{reseaux.find(r => String(r.ID) === ticketItem.field_7)?.field_1 ?? '—'}</dd>
                <dt>Montant</dt>
                <dd>{ticketItem.field_8?.toLocaleString('fr-FR') ?? '—'}</dd>
              </dl>

              {ticketItem.field_4 && (
                <div className="ticket-description">
                  <h3>Description / cause</h3>
                  <p>{stripHtml(ticketItem.field_4)}</p>
                </div>
              )}

              <div className="ticket-modal-actions">
                <button
                  type="button"
                  className="btn-cta btn-cta-secondary"
                  onClick={() => { setTicketItem(null); setDetailItem(ticketItem) }}
                >
                  Voir le détail complet
                </button>
                <button
                  type="button"
                  className="btn-cta btn-cta-status"
                  onClick={() => { const t = ticketItem; setTicketItem(null); openStatusModal(t) }}
                >
                  Changer le statut
                </button>
                {ticketItem.field_10 !== 'Clos' && (
                  <button
                    type="button"
                    className="btn-cta btn-cta-primary btn-cta-resolve"
                    onClick={() => openResolution(ticketItem)}
                  >
                    ✓ Clore la résolution
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {resolutionItem !== null && (
        <div className="modal-overlay" onClick={closeResolution}>
          <form className="modal resolution-modal" onClick={e => e.stopPropagation()} onSubmit={saveResolution}>
            <div className="modal-header">
              <div className="ticket-modal-title">
                <span className={`ticket-badge ticket-badge-lg ${statusBadgeClass(resolutionForm.statut)}`}>
                  <span className="ticket-badge-icon" aria-hidden="true">{statusIcon(resolutionForm.statut)}</span>
                  <span className="ticket-badge-num">T-{resolutionItem.ID}</span>
                </span>
                <h2>Clôture de la résolution</h2>
              </div>
              <button type="button" className="modal-close" onClick={closeResolution} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              <p className="resolution-intro">
                Renseignez les actions menées par le contrôleur pour résoudre cette anomalie.
                Le bulletin sera automatiquement consultable dans le sous-menu « Bulletins d'anomalies » après enregistrement.
              </p>

              <div className="resolution-grid">
                <div className="form-field">
                  <label htmlFor="resolution-statut">Statut final</label>
                  <select
                    id="resolution-statut"
                    value={resolutionForm.statut}
                    onChange={e => updateResolutionForm('statut', e.target.value)}
                  >
                    <option value="Resolu">Résolu</option>
                    <option value="Clos">Clos</option>
                  </select>
                </div>
                <div className="form-field">
                  <label htmlFor="resolution-dateRegul">Date de régularisation</label>
                  <input
                    id="resolution-dateRegul"
                    type="date"
                    value={resolutionForm.dateRegul}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => updateResolutionForm('dateRegul', e.target.value)}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="resolution-dateCloture">Date de clôture du ticket</label>
                  <input
                    id="resolution-dateCloture"
                    type="date"
                    value={resolutionForm.dateCloture}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => updateResolutionForm('dateCloture', e.target.value)}
                  />
                </div>
              </div>

              <div className="form-field">
                <label htmlFor="resolution-causeImmediate">Cause immédiate</label>
                <textarea
                  id="resolution-causeImmediate"
                  rows={2}
                  value={resolutionForm.causeImmediate}
                  onChange={e => updateResolutionForm('causeImmediate', e.target.value)}
                  placeholder="Quelle est la cause directe de l'anomalie ?"
                />
              </div>

              <div className="form-field">
                <label htmlFor="resolution-causeRacine">Cause racine</label>
                <textarea
                  id="resolution-causeRacine"
                  rows={2}
                  value={resolutionForm.causeRacine}
                  onChange={e => updateResolutionForm('causeRacine', e.target.value)}
                  placeholder="Pourquoi cette cause s'est-elle produite (analyse de fond) ?"
                />
              </div>

              <div className="form-field">
                <label htmlFor="resolution-actions">
                  Actions menées <span className="required-mark">*</span>
                  <small className="field-hint" style={{ marginLeft: 8 }}>une action par ligne</small>
                </label>
                <textarea
                  id="resolution-actions"
                  rows={5}
                  value={resolutionForm.actionsMenees}
                  onChange={e => updateResolutionForm('actionsMenees', e.target.value)}
                  placeholder={'Ex :\n- Vérification des écritures\n- Régularisation comptable\n- Notification du déclarant'}
                  required
                />
              </div>

              <div className="form-field">
                <label htmlFor="resolution-observations">Observations</label>
                <textarea
                  id="resolution-observations"
                  rows={2}
                  value={resolutionForm.observations}
                  onChange={e => updateResolutionForm('observations', e.target.value)}
                  placeholder="Remarques additionnelles, recommandations..."
                />
              </div>

              {resolutionError && <p className="status-error" role="alert">{resolutionError}</p>}

              <div className="resolution-actions">
                <button type="button" className="btn-cta btn-cta-secondary" onClick={closeResolution} disabled={resolutionSaving}>
                  Annuler
                </button>
                <button type="submit" className="btn-cta btn-cta-primary btn-cta-resolve" disabled={resolutionSaving}>
                  {resolutionSaving ? 'Enregistrement...' : `Enregistrer (${resolutionForm.statut})`}
                </button>
              </div>
            </div>
          </form>
        </div>
      )}

      {statusItem !== null && (
        <div className="modal-overlay" onClick={closeStatusModal}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 460 }}>
            <div className="modal-header">
              <h2>Changer le statut — Ticket T-{statusItem.ID}</h2>
              <button className="modal-close" onClick={closeStatusModal}>&times;</button>
            </div>
            <div className="modal-body">
              <div className="form-field">
                <label>Statut actuel</label>
                <input type="text" readOnly value={statusItem.field_10 ?? '-'} />
              </div>
              <div className="form-field">
                <label>Nouveau statut</label>
                <select
                  value={statusValue}
                  onChange={e => setStatusValue(e.target.value)}
                  autoFocus
                >
                  <option value="">-- Choisir --</option>
                  <option value="Ouvert">Ouvert</option>
                  <option value="En cours">En cours</option>
                  <option value="Resolu">Résolu</option>
                  <option value="Clos">Clos</option>
                </select>
              </div>
              {(statusValue === 'Resolu' || statusValue === 'Clos') && (
                <div className="form-field">
                  <label>Date de régularisation</label>
                  <input
                    type="date"
                    value={statusDateRegul}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => setStatusDateRegul(e.target.value)}
                  />
                </div>
              )}
              {(statusValue === 'Resolu' || statusValue === 'Clos') && (
                <div className="form-field">
                  <label>Date de clôture du ticket</label>
                  <input
                    type="date"
                    value={statusDateCloture}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => setStatusDateCloture(e.target.value)}
                  />
                </div>
              )}
              {(statusValue === 'Resolu' || statusValue === 'Clos') && (
                <p className="status-hint">
                  Le bulletin d'anomalie sera automatiquement disponible dans le sous-menu « Bulletins d'anomalies » après enregistrement.
                </p>
              )}
              {statusError && <p className="status-error" role="alert">{statusError}</p>}
              <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end' }}>
                <button type="button" className="btn-reset-filters" onClick={closeStatusModal} disabled={statusSaving}>
                  Annuler
                </button>
                <button type="button" className="btn-submit" style={{ width: 'auto', marginTop: 0 }} onClick={saveStatus} disabled={statusSaving || !statusValue}>
                  {statusSaving ? 'Enregistrement...' : 'Enregistrer'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  )
}
