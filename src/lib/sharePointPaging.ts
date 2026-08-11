/**
 * ============================================================================
 * SHAREPOINT PAGING — Helper universel pour récupérer TOUTES les pages
 * ============================================================================
 *
 * Pourquoi ce helper ?
 * --------------------
 * Le connecteur SharePoint du SDK Power Apps Code Apps ne renvoie qu'UNE
 * page par appel à `getAll()` (typiquement 100 items par défaut). Un
 * `getAll()` naïf renvoie donc **seulement la première page** — sur une
 * liste de 300 rapports, l'appelant n'en verra que 100, les 200 plus
 * anciens seront invisibles. C'est ce qui a causé le bug de production
 * sur les pages "Validation reporting" et "Mes rapports" (rapports
 * anciens absents alors qu'ils existent bien côté SP).
 *
 * Comment fonctionne la pagination du SDK ?
 * -----------------------------------------
 * ⚠ Le SDK Power Apps 0.0.4 **ne remonte PAS** de `skipToken` pour le
 * connecteur SharePoint (uniquement pour Dataverse). Le champ existe
 * dans le type `IOperationResult` mais reste toujours `undefined` sur
 * SharePoint (cf. `runtimeDataClient.js` : le connecteur retourne
 * `parsedResult.value` sans propager `@odata.nextLink`).
 *
 * → Impossible de paginer via `skipToken`.
 *
 * En revanche, le SDK honore bien `$top` et `$skip` en URL OData
 * (cf. `runtimeDataOperations.js`). On pagine donc **par offset** :
 * on incrémente `skip` de `top` à chaque tour, jusqu'à recevoir une
 * page plus courte que `top` (signal de fin).
 *
 * Utilisation
 * -----------
 *   const all = await getAllPages<MyModelRead>(MyModelService, {
 *     orderBy: ['Created desc'],
 *     filter: "field_10 eq 'Ouvert'",
 *   })
 *
 * ⚠ ATTENTION VOLUMÉTRIE
 * ----------------------
 * Sur une liste de plusieurs milliers d'items sans filtre, ce helper peut
 * faire des dizaines d'appels HTTP en cascade. TOUJOURS accompagner d'un
 * `filter` OData restrictif (par période, par statut, etc.) pour ne
 * rapatrier que ce qui est nécessaire à la vue courante.
 * ============================================================================
 */

import type { IGetAllOptions } from '../generated/models/CommonModels'
import type { IOperationResult } from '@pa-client/power-code-sdk'

/**
 * Interface minimale d'un service SharePoint généré : on n'a besoin que
 * de sa méthode statique `getAll`. Cela permet au helper d'accepter
 * n'importe quel service généré (Anomalies, Contrôleurs, PACs, etc.)
 * sans coupler à un type précis.
 */
export interface SharePointService<T> {
  getAll(options?: IGetAllOptions): Promise<IOperationResult<T[]>>
}

/**
 * Taille d'une page OData. Le connecteur SharePoint plafonne `$top` à
 * 5000 côté SP (List View Threshold), mais 500 est un compromis raisonné :
 *   - Assez grand pour minimiser le nombre d'appels réseau
 *   - Assez petit pour rester sous les limites de payload et éviter les
 *     timeouts sur les listes lourdes (Person, Choice, etc. gonflent la
 *     réponse).
 */
const PAGE_SIZE = 500

/**
 * Garde-fou : nombre maximum de pages parcourues en cascade.
 * À 100 000 items (200 pages × 500 items/page), on est bien au-delà des
 * besoins de l'app. Si jamais on l'atteint, c'est symptomatique d'une
 * requête sans filtre → il faut ajouter une contrainte OData en amont
 * plutôt que d'augmenter cette borne.
 */
const MAX_PAGES = 200

/**
 * Récupère TOUS les items d'une liste SharePoint en paginant par offset
 * (`$top` + `$skip`) — car le SDK Power Apps ne remonte pas le
 * `skipToken` pour SharePoint (cf. bloc doc en tête de fichier).
 *
 * Comportement :
 *   - Boucle en incrémentant `skip` de `PAGE_SIZE` à chaque tour
 *   - Concatène les `data` de chaque page dans un unique tableau
 *   - Condition d'arrêt : la page renvoyée contient MOINS de `PAGE_SIZE`
 *     items → dernière page atteinte
 *   - Si un appel échoue (success=false) en cours de route, on stoppe
 *     la boucle et on retourne ce qu'on a collecté (fail-soft)
 *   - Cap de sécurité : `MAX_PAGES` pour éviter les boucles infinies
 *
 * @param service Service SharePoint généré (classe avec méthode statique
 *                `getAll`), ex: `DCPO_ACTIVICTE_CONTROLLERService`.
 * @param options Options OData standard (filter, orderBy, select). Les
 *                clés `top` et `skip` sont gérées en interne — ne pas
 *                les passer (elles seraient écrasées).
 * @returns Tous les items concaténés (peut être vide si erreur immédiate)
 */
export async function getAllPages<T>(
  service: SharePointService<T>,
  options?: Omit<IGetAllOptions, 'top' | 'skip' | 'skipToken'>,
): Promise<T[]> {
  const all: T[] = []
  // Dédoublonnage par ID SharePoint pour se prémunir contre :
  //   1. Un connecteur qui IGNORE `$skip` (comportement observé sur certaines
  //      listes SP → chaque page renvoie les mêmes items, ce qui fait boucler
  //      jusqu'au cap MAX_PAGES et remonter 100 000 items dupliqués).
  //   2. Une race condition où un item modifié entre 2 pages apparaît sur
  //      les deux (rare mais possible).
  const seenIds = new Set<string>()
  let skip = 0
  let pageCount = 0

  while (true) {
    const result: IOperationResult<T[]> = await service.getAll({
      ...(options ?? {}),
      top: PAGE_SIZE,
      skip,
    })
    if (!result.success) {
      // Erreur réseau ou OData rejeté → on log et on retourne ce qu'on a.
      console.warn('getAllPages: appel échoué, arrêt de la pagination', result.error)
      break
    }
    const batch = result.data ?? []

    // Ajout dédoublonné : on ne push que les items dont l'ID SP n'a pas
    // encore été vu. Compte les VRAIS nouveaux items pour détecter le cas
    // "connecteur ignore $skip".
    let addedThisPage = 0
    for (const item of batch) {
      const rawId = (item as { ID?: string | number }).ID
      // Fallback défensif : si un item n'a pas d'ID (cas très rare, données
      // corrompues), on utilise sa position pour éviter de le perdre.
      const key = rawId !== undefined && rawId !== null
        ? String(rawId)
        : `__noid_${all.length + addedThisPage}`
      if (!seenIds.has(key)) {
        seenIds.add(key)
        all.push(item)
        addedThisPage++
      }
    }
    pageCount++

    // Fin naturelle : page pas remplie = plus rien à récupérer.
    if (batch.length < PAGE_SIZE) break

    // Fin défensive : le connecteur a renvoyé UNIQUEMENT des items déjà vus
    // → il ignore probablement `$skip`. On stoppe pour éviter la boucle
    // pathologique qui accumule 200 × PAGE_SIZE doublons.
    if (addedThisPage === 0) {
      console.warn(
        `getAllPages: page ${pageCount} sans nouveaux items (connecteur ` +
        `qui ignore $skip ?). Arrêt de la pagination — ${all.length} item(s) ` +
        `uniques retournés.`,
      )
      break
    }

    if (pageCount >= MAX_PAGES) {
      console.warn(
        `getAllPages: cap de ${MAX_PAGES} pages atteint (` +
        `${all.length} items chargés). Ajoutez un filtre OData en amont.`,
      )
      break
    }

    skip += PAGE_SIZE
  }

  return all
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — CHUNKING PAR MOIS CALENDAIRE
 *
 * Contexte : le connecteur SharePoint plafonne pratiquement à ~500 items
 * par appel `getAll`, `$skip` étant souvent ignoré. Pour dépasser ce
 * plafond sur une longue période, on émet PLUSIEURS requêtes OData,
 * chacune filtrée sur un mois calendaire. Les résultats sont fusionnés
 * et dédoublonnés par ID.
 *
 * Utilisation typique :
 *   const items = await getAllChunkedByMonth(MyService, {
 *     from: '2026-01-01',
 *     to: '2026-12-31',
 *     dateColumn: 'Created',
 *     extraFilter: "statut eq 'Actif'",
 *     orderBy: ['Created desc'],
 *   })
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Découpe une plage `[from, to]` (dates YYYY-MM-DD) en sous-plages
 * mensuelles calendaires.
 *
 * Exemple : `splitRangeByMonth('2026-01-15', '2026-03-10')` →
 *   [ {from: '2026-01-15', to: '2026-01-31'},
 *     {from: '2026-02-01', to: '2026-02-28'},
 *     {from: '2026-03-01', to: '2026-03-10'} ]
 */
export function splitRangeByMonth(
  from: string,
  to: string,
): Array<{ from: string; to: string }> {
  const chunks: Array<{ from: string; to: string }> = []
  const [fy, fm, fd] = from.split('-').map(Number)
  const [ty, tm, td] = to.split('-').map(Number)

  let cursorY = fy
  let cursorM = fm
  let isFirst = true

  while (cursorY < ty || (cursorY === ty && cursorM <= tm)) {
    const startDay = isFirst ? fd : 1
    const startStr = `${cursorY}-${String(cursorM).padStart(2, '0')}-${String(startDay).padStart(2, '0')}`

    const isLast = cursorY === ty && cursorM === tm
    // `new Date(y, m, 0).getDate()` = dernier jour du mois `m` (mois 1-indexé)
    const endDay = isLast ? td : new Date(cursorY, cursorM, 0).getDate()
    const endStr = `${cursorY}-${String(cursorM).padStart(2, '0')}-${String(endDay).padStart(2, '0')}`

    chunks.push({ from: startStr, to: endStr })
    isFirst = false

    cursorM++
    if (cursorM > 12) {
      cursorM = 1
      cursorY++
    }
  }
  return chunks
}

/**
 * Exécute des tâches asynchrones par vagues de `concurrency` items en
 * parallèle. Évite d'ouvrir trop de connexions HTTP en même temps sur le
 * connecteur SP (rate limit + risque de saturation).
 */
export async function runWithConcurrency<T>(
  tasks: Array<() => Promise<T>>,
  concurrency: number,
): Promise<T[]> {
  const results: T[] = new Array(tasks.length)
  let index = 0
  const workers = Array.from(
    { length: Math.min(concurrency, tasks.length) },
    async () => {
      while (index < tasks.length) {
        const i = index++
        results[i] = await tasks[i]()
      }
    },
  )
  await Promise.all(workers)
  return results
}

/** Options pour `getAllChunkedByMonth`. */
export interface ChunkedByMonthOptions {
  /** Borne basse YYYY-MM-DD (incluse). */
  from: string
  /** Borne haute YYYY-MM-DD (incluse). */
  to: string
  /** Nom de la colonne SP sur laquelle appliquer le filtre par plage. */
  dateColumn: string
  /**
   * Clause OData additionnelle combinée en AND avec le filtre par plage.
   * Ex : `"field_10 eq 'Resolu' or field_10 eq 'Clos'"`.
   * Sera automatiquement entourée de parenthèses.
   */
  extraFilter?: string
  /** orderBy OData standard. Par défaut : `[<dateColumn> desc]`. */
  orderBy?: string[]
  /** Nombre max de fetch parallèles. Par défaut : 4. */
  concurrency?: number
}

/**
 * Récupère TOUS les items d'une liste sur une plage de dates, en chunkant
 * la plage par MOIS calendaire pour contourner le plafond de 500 items du
 * connecteur SP. Chaque chunk est fetché en parallèle (limité à
 * `concurrency`), puis les résultats sont fusionnés + dédoublonnés + triés.
 *
 * @returns Tous les items uniques de la plage, triés selon `orderBy`.
 */
export async function getAllChunkedByMonth<T>(
  service: SharePointService<T>,
  options: ChunkedByMonthOptions,
): Promise<T[]> {
  const chunks = splitRangeByMonth(options.from, options.to)
  const orderBy = options.orderBy ?? [`${options.dateColumn} desc`]
  const concurrency = options.concurrency ?? 4

  const perChunk = await runWithConcurrency(
    chunks.map(chunk => () => {
      const rangeClause =
        `${options.dateColumn} ge '${chunk.from}T00:00:00Z' and ` +
        `${options.dateColumn} le '${chunk.to}T23:59:59Z'`
      const filter = options.extraFilter
        ? `(${options.extraFilter}) and ${rangeClause}`
        : rangeClause
      return getAllPages<T>(service, { orderBy, filter })
    }),
    concurrency,
  )

  // Fusion + dédup par ID (belt & braces — les chunks ne se chevauchent pas)
  const seen = new Set<string>()
  const merged: T[] = []
  for (const batch of perChunk) {
    for (const item of batch) {
      const rawId = (item as { ID?: string | number }).ID
      const key = rawId !== undefined && rawId !== null ? String(rawId) : ''
      if (key && seen.has(key)) continue
      if (key) seen.add(key)
      merged.push(item)
    }
  }
  return merged
}
