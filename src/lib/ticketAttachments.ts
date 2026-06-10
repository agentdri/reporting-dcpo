/**
 * ============================================================================
 * GESTION DES PIÈCES JOINTES DES TICKETS D'ANOMALIES
 * ============================================================================
 *
 * Encapsule toute la logique liée aux pièces jointes :
 *   - Détection du type de fichier (image / pdf / word / excel / etc.)
 *   - Choix d'une icône emoji adaptée
 *   - Upload via Power Automate (workflow externe)
 *   - Lecture des pièces jointes depuis un ticket SharePoint (deux sources :
 *     {Attachments} natif + urlPieceJointe custom multi-URL)
 *   - Concaténation / parsing des URLs multiples dans le champ urlPieceJointe
 *
 * Architecture des pièces jointes côté SharePoint :
 *
 *   Source 1 — {Attachments} (collection native SharePoint)
 *     Tableau d'objets avec AbsoluteUri + DisplayName.
 *     Alimenté automatiquement quand on attache un fichier à l'item via
 *     le workflow Power Automate.
 *
 *   Source 2 — urlPieceJointe (champ texte custom)
 *     Stocke des URLs concaténées par " | " (séparateur custom). Permet
 *     de garder un lien direct cliquable même si l'attachment SharePoint
 *     change d'URL (ex : déplacement de bibliothèque).
 *
 *   getTicketAttachments() merge ces deux sources en dédoublonnant par URL.
 * ============================================================================
 */

import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Représentation enrichie d'une pièce jointe pour l'affichage.
 *
 *   - name     : nom de fichier lisible (extrait du DisplayName ou de l'URL)
 *   - url      : URL absolue vers le fichier
 *   - isImage  : true si extension image → permet l'affichage en miniature
 *   - iconType : type pour le mapping vers emoji (cf. getAttachmentIcon)
 */
export interface TicketAttachment {
  name: string
  url: string
  isImage: boolean
  iconType: 'image' | 'pdf' | 'word' | 'excel' | 'archive' | 'text' | 'file'
}

/**
 * Réponse retournée par le workflow Power Automate après upload.
 *   - success         : booléen indiquant le succès
 *   - sharePointId    : ID de l'item SharePoint touché (rappel)
 *   - attachmentUrl   : URL absolue du fichier uploadé (à stocker en BDD)
 *   - message         : message d'erreur éventuel
 */
export interface UploadResponse {
  success: boolean
  sharePointId?: number
  attachmentUrl?: string
  message?: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — URL DU WORKFLOW POWER AUTOMATE
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Endpoint HTTP du flux Power Automate qui :
 *   1. Reçoit { fileName, fileContent (base64), Choise, idItem }
 *   2. Décode le base64 et attache le fichier à l'item SharePoint indiqué
 *   3. Retourne l'URL absolue de la pièce jointe créée
 *
 * Note de sécurité :
 *   L'URL contient une signature `sig=...` propre à ce flux. Elle vaut
 *   "clé d'API" — quiconque a l'URL peut déclencher des uploads. C'est
 *   acceptable parce que :
 *     - Le flux exige idItem (un ID existant) → pas de création anarchique
 *     - Les ACL SharePoint s'appliquent quand l'item est lu
 *     - Le flux ne peut être appelé que via HTTPS depuis l'app Power Apps
 *
 * Si on devait régénérer cette URL : il faut éditer le déclencheur HTTP
 * du flow dans Power Automate, puis remettre l'URL ICI uniquement.
 */
export const ATTACHMENT_API_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a8ced63bd1314a7897003b97eebd85e9/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EdDW19nkBl6pWigNfp0OQPaiSIjBwyhAuWeTO3s8b8E'

/**
 * Endpoint Power Automate dédié à l'upload des pièces jointes des rapports
 * d'activité contrôleur (liste DCPO_ACTIVICTE_CONTROLLER).
 *
 * Distinct de ATTACHMENT_API_URL ci-dessus : chaque flow est lié à une
 * liste cible spécifique côté Power Automate.
 */
export const ACTIVITY_ATTACHMENT_API_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/092b4922268f439eb5cdbc77960a2875/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=txQA9pHc3e3Qb5ziRPWlG7uTiWIy1LuGFm0obmZBbS0'

/**
 * Endpoint Power Automate dédié aux pièces jointes du Plan de Contrôle
 * (liste DCPO_LISTE_PLAN_CONTROLE). Workflow indépendant de ceux ci-dessus.
 */
export const PLAN_CONTROLE_ATTACHMENT_API_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/cc6b39a566b64d0aadd786e92ccb77d0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TmPz_HfScrWpBJym-7vh5trjQf3DP7jYXMED8l3PDMc'

/**
 * Endpoint Power Automate dédié aux pièces jointes du Plan d'Action Correctif
 * (liste DCPO_LISTE_PLAN_ACTION_CORRECTIF). Workflow indépendant.
 */
export const PAC_ATTACHMENT_API_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7b31444423c94e6083979d95929f1992/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=sLe6L6-bAw6dp1V56TrljC8gVebCwVVnWcqJXys1GCY'

/**
 * Endpoint Power Automate dédié aux pièces jointes des ÉVALUATIONS du
 * Plan de Contrôle (liste DCPO_EVALUATION_PLAN_CONTROLE). Workflow indépendant.
 */
export const EVALUATION_ATTACHMENT_API_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/34768a8a92324fc3a3e5181d95b604d0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fFhe-8z3azTw1jkdekzHF8_inr9zsH7ydsSrUr7Z2ks'

/**
 * Endpoint Power Automate dédié aux pièces jointes des ÉVALUATIONS du
 * Plan d'Action Correctif (liste DCPO_EVALUATION_PLAN_ACTION_CORRECTIF).
 * Workflow indépendant — distinct du flux d'upload PAC lui-même
 * (PAC_ATTACHMENT_API_URL) et du flux des évaluations Plan de Contrôle.
 */
export const PAC_EVALUATION_ATTACHMENT_API_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/3bfe06475a9c4e0c99306ea3ba62cad1/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=1u00IQtrIXJ8CKgxFHYbwf6dLjrlrA1xswB9a7w0tI8'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — UTILITAIRES BAS-NIVEAU
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Convertit un File (issu d'un input file) en string base64.
 *
 * Pourquoi base64 ?
 *   Le workflow Power Automate accepte du JSON. On ne peut pas envoyer
 *   un binaire brut, donc on encode le contenu en base64 ASCII.
 *
 * Mécanisme :
 *   FileReader.readAsDataURL → produit une string format
 *   "data:image/png;base64,iVBORw0KGgo..."
 *   On split sur la virgule pour ne garder QUE le base64 (pas le préfixe).
 *
 * Note : pour de gros fichiers (> 5 MB), cette opération peut être lente
 * et bloquer un peu le thread principal. À considérer si on rencontre
 * des problèmes UX.
 */
export function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      const result = reader.result as string
      // result inclut le préfixe data: → on isole le base64 après la virgule
      const base64 = result.includes(',') ? result.split(',')[1] : result
      resolve(base64)
    }
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

/**
 * Extrait le nom de fichier depuis une URL.
 *
 * Stratégie :
 *   1. Parser l'URL avec new URL() (plus robuste qu'une regex)
 *   2. Récupérer le pathname (partie après le domaine)
 *   3. Découper sur '/' et prendre le dernier segment
 *   4. decodeURIComponent pour transformer "rapport%20final.pdf" en
 *      "rapport final.pdf"
 *
 * Fallback : si l'URL est invalide (pas un format URL standard),
 * on retourne l'URL telle quelle plutôt que de planter.
 */
export function getFileNameFromUrl(url: string): string {
  try {
    const pathname = new URL(url).pathname
    return decodeURIComponent(pathname.split('/').pop() ?? url)
  } catch {
    return url
  }
}

/**
 * Détecte si une URL pointe vers un fichier image (par extension).
 * Utilisé pour décider si on affiche une miniature ou juste un lien.
 */
export function isImageUrl(url: string): boolean {
  return /\.(png|jpe?g|gif|bmp|webp|svg)(\?|$)/i.test(url)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — MAPPING VERS ICÔNES
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Détermine le type d'icône à afficher selon l'extension du fichier.
 *
 * Patterns reconnus :
 *   - image    : png, jpg, jpeg, gif, bmp, webp, svg
 *   - pdf      : pdf
 *   - word     : doc, docx, odt
 *   - excel    : xls, xlsx, csv, ods
 *   - archive  : zip, rar, 7z, tar, gz
 *   - text     : txt, md, log
 *   - file     : tout le reste (fallback générique)
 *
 * `(\?|$)` dans les regex : permet aux URLs avec query string
 * (ex : "?download=1") d'être détectées correctement.
 */
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

/**
 * Mappe un type d'icône vers son emoji représentatif.
 *
 * Pourquoi emoji et pas SVG ?
 *   - Zéro asset à charger
 *   - Bonne couverture visuelle multi-plateforme
 *   - Lisible immédiatement par l'utilisateur sans formation
 */
export function getAttachmentIcon(type: TicketAttachment['iconType']): string {
  switch (type) {
    case 'image': return '🖼️'
    case 'pdf': return '📄'
    case 'word': return '📝'
    case 'excel': return '📊'
    case 'archive': return '🗜️'
    case 'text': return '📃'
    default: return '📎'  // emoji "trombone" générique
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — LECTURE DES PIÈCES JOINTES D'UN TICKET
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Construit la liste consolidée des pièces jointes d'un ticket en
 * fusionnant deux sources :
 *
 *   1. {Attachments} : collection SharePoint native (alimentée automatiquement
 *      quand on attache un fichier via le workflow Power Automate)
 *
 *   2. urlPieceJointe : champ texte custom contenant une ou PLUSIEURS URLs
 *      séparées par " | " (cf. parseUrlList)
 *
 * Dédoublonnage :
 *   Si la même URL apparaît dans les deux sources, on ne la pousse qu'une
 *   fois (vérification via `.some(a => a.url === url)`).
 *
 * Ordre :
 *   {Attachments} en premier (pièces jointes natives historiques),
 *   urlPieceJointe ensuite (URLs ajoutées au fil des résolutions).
 */
export function getTicketAttachments(ticket: DCPO_LISTE_ANORMALIERead | null | undefined): TicketAttachment[] {
  if (!ticket) return []
  const out: TicketAttachment[] = []

  // ─── Source 1 : collection native SharePoint ──────────────────────
  const sharepointAttachments = ticket['{Attachments}']
  if (Array.isArray(sharepointAttachments)) {
    sharepointAttachments.forEach(att => {
      if (!att?.AbsoluteUri) return  // skip si entrée corrompue
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

  // ─── Source 2 : champ urlPieceJointe (multi-URLs séparées par |) ──
  if (ticket.urlPieceJointe) {
    const urls = parseUrlList(ticket.urlPieceJointe)
    for (const url of urls) {
      // Dédoublonnage : ne pas réajouter une URL déjà présente
      const exists = out.some(a => a.url === url)
      if (exists) continue
      const name = getFileNameFromUrl(url)
      const iconType = getAttachmentIconType(name)
      out.push({
        name,
        url,
        isImage: iconType === 'image',
        iconType,
      })
    }
  }

  return out
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 — GESTION MULTI-URL DANS urlPieceJointe
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Caractère utilisé pour séparer plusieurs URLs concaténées dans le
 * champ urlPieceJointe (qui est un text simple SharePoint).
 *
 * Choix de "|" (pipe) :
 *   - URL-safe (n'apparaît pas dans les URLs SharePoint réelles)
 *   - Visuellement distinct
 *   - Compatible single-line ET multi-line text fields
 */
export const URL_LIST_SEPARATOR = '|'

/**
 * Découpe la valeur brute de urlPieceJointe en URLs individuelles.
 *
 * Étapes :
 *   1. Si raw est null/undefined → []
 *   2. Split sur "|"
 *   3. Trim chaque morceau (supprime les espaces autour de " | ")
 *   4. Filter (Boolean) : élimine les chaînes vides (cas " | | ")
 *
 * Compatible avec l'ancien format "url unique sans séparateur" :
 *   "https://...pdf" → ["https://...pdf"] ✓
 */
export function parseUrlList(raw: string | null | undefined): string[] {
  if (!raw) return []
  return raw
    .split(URL_LIST_SEPARATOR)
    .map(s => s.trim())
    .filter(s => s.length > 0)
}

/**
 * Ajoute une nouvelle URL à une liste existante, sans doublon.
 *
 *   1. Parse l'existant
 *   2. Trim newUrl et vérifie qu'elle n'est pas déjà dans la liste
 *   3. Push si nouvelle
 *   4. Re-join avec " | " (avec espaces pour la lisibilité dans SharePoint)
 *
 * Use case typique : à la clôture d'une résolution, on ajoute la nouvelle
 * pièce jointe sans écraser celles existantes.
 */
export function appendUrl(existing: string | null | undefined, newUrl: string): string {
  const list = parseUrlList(existing)
  if (newUrl && !list.includes(newUrl.trim())) {
    list.push(newUrl.trim())
  }
  return list.join(` ${URL_LIST_SEPARATOR} `)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 7 — UPLOAD VIA POWER AUTOMATE
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Upload une pièce jointe à un item SharePoint via le workflow Power Automate.
 *
 * Étapes :
 *   1. Encoder le file en base64 (via FileReader)
 *   2. POST JSON vers ATTACHMENT_API_URL avec :
 *        - fileName     : nom du fichier original
 *        - fileContent  : contenu base64
 *        - Choise       : valeur métier (par défaut 'Visite' — le workflow
 *                         peut router vers différentes listes selon ce flag)
 *        - idItem       : ID de l'item SharePoint cible
 *   3. Récupérer l'URL absolue retournée par le workflow
 *
 * Gestion d'erreurs :
 *   Si la réponse HTTP n'est pas OK (4xx, 5xx), on lit le texte d'erreur
 *   et on lève une Error explicite. L'appelant doit catch et afficher.
 *
 * Retour :
 *   - URL absolue (string) si tout va bien
 *   - undefined si le workflow réussit mais ne retourne pas d'attachmentUrl
 *     (cas anormal, ne devrait pas arriver en production)
 *
 * @param itemId - ID de l'anomalie SharePoint à laquelle attacher le fichier
 * @param file - File object issu d'un <input type="file">
 * @param choise - Catégorie métier transmise au workflow (défaut 'Visite')
 */
/**
 * POST JSON vers un workflow Power Automate avec retry sur erreurs transitoires (5xx, network).
 *
 * Stratégie :
 *   - 1 tentative initiale + jusqu'à `maxRetries` retries
 *   - Backoff linéaire 2 s entre tentatives
 *   - Retry uniquement sur statuts 5xx OU exceptions réseau
 *   - Pas de retry sur 4xx (erreur client = la requête est mal formée)
 *
 * Le 502 "NoResponse" de Power Automate (upstream timeout) est typiquement
 * transitoire et bénéficie d'un retry.
 */
async function postWorkflowWithRetry(
  url: string,
  body: unknown,
  maxRetries: number = 1,
): Promise<Response> {
  const payload = JSON.stringify(body)
  let lastError: Error | undefined
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: payload,
      })
      if (response.ok) return response
      // 5xx = retry possible
      if (response.status >= 500 && response.status < 600 && attempt < maxRetries) {
        const errorText = await response.text().catch(() => '')
        console.warn(`Power Automate ${response.status} (tentative ${attempt + 1}/${maxRetries + 1}) — retry dans 2s : ${errorText}`)
        await new Promise(r => setTimeout(r, 2000))
        continue
      }
      // 4xx ou retries épuisés → on retourne la réponse pour que l'appelant lise l'erreur
      return response
    } catch (err) {
      lastError = err instanceof Error ? err : new Error(String(err))
      if (attempt < maxRetries) {
        console.warn(`Erreur réseau Power Automate (tentative ${attempt + 1}/${maxRetries + 1}) — retry dans 2s`, err)
        await new Promise(r => setTimeout(r, 2000))
        continue
      }
    }
  }
  throw lastError ?? new Error('Échec de l\'upload après retries.')
}

export async function uploadTicketAttachment(
  itemId: string,
  file: File,
  choise: string = 'Visite',
): Promise<string | undefined> {
  const fileContent = await fileToBase64(file)
  const response = await postWorkflowWithRetry(ATTACHMENT_API_URL, {
    fileName: file.name,
    fileContent,
    Choise: choise,
    idItem: itemId,
  })
  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Echec upload piece jointe (${response.status}): ${errorText}`)
  }
  const data: UploadResponse = await response.json()
  return data.attachmentUrl
}

/**
 * Upload une pièce jointe à un item de la liste DCPO_ACTIVICTE_CONTROLLER
 * (rapport quotidien contrôleur) via le workflow Power Automate dédié.
 *
 * Mêmes étapes que uploadTicketAttachment :
 *   1. Encoder le file en base64
 *   2. POST JSON vers ACTIVITY_ATTACHMENT_API_URL avec :
 *      - fileName / fileContent / idItem
 *      - Choise = 'Activite' (catégorie métier transmise au workflow)
 *   3. Récupérer l'URL absolue de la pièce jointe créée
 *
 * @param itemId - ID SharePoint du rapport d'activité cible
 * @param file - File issu d'un <input type="file">
 * @param choise - Catégorie métier (défaut 'Activite')
 * @returns URL absolue de la pièce jointe attachée (ou undefined si réponse vide)
 */
export async function uploadActivityAttachment(
  itemId: string,
  file: File,
  choise: string = 'Visite',
): Promise<string | undefined> {
  const fileContent = await fileToBase64(file)
  const response = await postWorkflowWithRetry(ACTIVITY_ATTACHMENT_API_URL, {
    fileName: file.name,
    fileContent,
    Choise: choise,
    idItem: itemId,
  })
  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Echec upload piece jointe activite (${response.status}): ${errorText}`)
  }
  const data: UploadResponse = await response.json()
  return data.attachmentUrl
}

/**
 * Upload une pièce jointe à un item de la liste DCPO_LISTE_PLAN_CONTROLE.
 * Workflow Power Automate dédié (PLAN_CONTROLE_ATTACHMENT_API_URL).
 *
 * Le fichier est attaché à l'item SharePoint via la collection {Attachments}
 * native (alimentée par le workflow). Le workflow retourne aussi l'URL
 * absolue de la pièce jointe — la liste n'a pas de champ texte custom pour
 * la stocker, mais l'URL reste accessible via {Attachments} à la lecture.
 *
 * @param itemId - ID SharePoint du contrôle cible
 * @param file - File issu d'un <input type="file">
 * @param choise - Catégorie métier (défaut 'Visite' — selon attente du workflow)
 * @returns URL absolue de la pièce jointe (ou undefined si réponse vide)
 */
export async function uploadPlanControleAttachment(
  itemId: string,
  file: File,
  choise: string = 'Visite',
): Promise<string | undefined> {
  const fileContent = await fileToBase64(file)
  const response = await postWorkflowWithRetry(PLAN_CONTROLE_ATTACHMENT_API_URL, {
    fileName: file.name,
    fileContent,
    Choise: choise,
    idItem: itemId,
  })
  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Echec upload piece jointe plan de controle (${response.status}): ${errorText}`)
  }
  const data: UploadResponse = await response.json()
  return data.attachmentUrl
}

/**
 * Upload une pièce jointe à un item de la liste DCPO_LISTE_PLAN_ACTION_CORRECTIF.
 * Workflow Power Automate dédié (PAC_ATTACHMENT_API_URL). Même mécanique que
 * uploadPlanControleAttachment ci-dessus, juste un endpoint différent.
 *
 * @param itemId - ID SharePoint du PAC cible
 * @param file - File issu d'un <input type="file">
 * @param choise - Catégorie métier (défaut 'Visite' — selon attente du workflow)
 * @returns URL absolue de la pièce jointe (ou undefined si réponse vide)
 */
export async function uploadPacAttachment(
  itemId: string,
  file: File,
  choise: string = 'Visite',
): Promise<string | undefined> {
  const fileContent = await fileToBase64(file)
  const response = await postWorkflowWithRetry(PAC_ATTACHMENT_API_URL, {
    fileName: file.name,
    fileContent,
    Choise: choise,
    idItem: itemId,
  })
  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Echec upload piece jointe PAC (${response.status}): ${errorText}`)
  }
  const data: UploadResponse = await response.json()
  return data.attachmentUrl
}

/**
 * Upload une pièce jointe à un item de la liste DCPO_EVALUATION_PLAN_CONTROLE
 * (évaluation périodique d'un contrôle). Workflow Power Automate dédié.
 *
 * Identique aux autres uploadXxxAttachment : encodage base64, POST JSON,
 * retour de l'URL absolue du fichier attaché.
 *
 * @param itemId - ID SharePoint de l'évaluation cible
 * @param file - File issu d'un <input type="file">
 * @param choise - Catégorie métier (défaut 'Visite')
 * @returns URL absolue de la pièce jointe (ou undefined si réponse vide)
 */
export async function uploadEvaluationAttachment(
  itemId: string,
  file: File,
  choise: string = 'Visite',
): Promise<string | undefined> {
  const fileContent = await fileToBase64(file)
  const response = await postWorkflowWithRetry(EVALUATION_ATTACHMENT_API_URL, {
    fileName: file.name,
    fileContent,
    Choise: choise,
    idItem: itemId,
  })
  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Echec upload piece jointe évaluation (${response.status}): ${errorText}`)
  }
  const data: UploadResponse = await response.json()
  return data.attachmentUrl
}

/**
 * Upload une pièce jointe à un item de la liste
 * DCPO_EVALUATION_PLAN_ACTION_CORRECTIF (évaluation d'un PAC).
 * Workflow Power Automate dédié.
 *
 * @param itemId - ID SharePoint de l'évaluation PAC cible
 * @param file - File issu d'un <input type="file">
 * @param choise - Catégorie métier (défaut 'Visite')
 * @returns URL absolue de la pièce jointe (ou undefined si réponse vide)
 */
export async function uploadPacEvaluationAttachment(
  itemId: string,
  file: File,
  choise: string = 'Visite',
): Promise<string | undefined> {
  const fileContent = await fileToBase64(file)
  const response = await postWorkflowWithRetry(PAC_EVALUATION_ATTACHMENT_API_URL, {
    fileName: file.name,
    fileContent,
    Choise: choise,
    idItem: itemId,
  })
  if (!response.ok) {
    const errorText = await response.text()
    throw new Error(`Echec upload piece jointe évaluation PAC (${response.status}): ${errorText}`)
  }
  const data: UploadResponse = await response.json()
  return data.attachmentUrl
}
