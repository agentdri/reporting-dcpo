/**
 * ============================================================================
 * NOTIFICATIONS TEAMS — AFFECTATIONS
 * ============================================================================
 *
 * Envoie une notification Microsoft Teams (chat 1-to-1 via Flow bot) à la
 * personne qui vient d'être affectée à un item de la plateforme :
 *   - Anomalie : bouton "Affecter" ou champ "Personne affectée" du formulaire
 *     de création
 *   - Plan de Contrôle : modale "Affectation" ou champ responsable à la création
 *   - PAC : idem
 *
 * IMPLÉMENTATION
 * --------------
 * Un unique workflow Power Automate (URL ci-dessous) accepte un JSON décrivant
 * l'affectation et publie un message ciblé dans Teams. Côté front, on appelle
 * ce workflow en "fire-and-forget" — la notification ne doit JAMAIS bloquer
 * la sauvegarde principale (= mise à jour SP).
 *
 * Si Teams / le workflow échoue, on log mais on n'affiche aucune erreur à
 * l'utilisateur : son action métier (affectation) a réussi de toute façon.
 *
 * POUR MODIFIER L'URL DU FLOW
 * ---------------------------
 * Si le flow Power Automate est régénéré (nouvelle signature `sig=...`),
 * mettre à jour `TEAMS_NOTIFY_URL` ici uniquement. Tous les points d'appel
 * de l'app utiliseront automatiquement la nouvelle URL.
 *
 * SCHEMA ATTENDU CÔTÉ FLOW
 * ------------------------
 *   {
 *     "type":    "anomalie" | "plan-controle" | "pac",
 *     "email":   "<email du destinataire>",
 *     "subject": "<titre court>",
 *     "details": "<phrases supplémentaires (échéance, commentaire...)>",
 *     "appUrl":  "<lien direct vers l'item ou la page>"  // optionnel
 *   }
 * ============================================================================
 */

/**
 * URL du flow Power Automate qui poste dans Teams.
 *
 * Workflow : `df9a997557cc4d4f8d3693b40db28e63`
 * Trigger  : HTTP manuel (signature `sig=...` dans l'URL)
 * Action   : Post message in a chat (Flow bot) au destinataire `email`.
 */
const TEAMS_NOTIFY_URL =
  'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/df9a997557cc4d4f8d3693b40db28e63/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=cTkMeE3-Lbet9tNPcAsqEaMsVFCi1Ox18bCAT6re5zQ'

/** Type métier de l'affectation — permet au flow d'adapter le wording. */
export type AffectationType = 'anomalie' | 'plan-controle' | 'pac'

/** Payload envoyé au workflow Power Automate. */
export interface NotifyAffectationInput {
  /** Type métier (anomalie / plan-controle / pac). */
  type: AffectationType
  /** Email du destinataire (= personne affectée). REQUIS. */
  email: string
  /** Titre court qui figurera en gras dans le message Teams. */
  subject: string
  /** Phrases additionnelles (délai, commentaire, échéance...). Optionnel. */
  details?: string
  /** Lien profond vers l'item (URL SharePoint ou page de l'app). Optionnel. */
  appUrl?: string
}

/**
 * Envoie une notification Teams à la personne affectée.
 *
 * Pattern "fire-and-forget" : on lance la requête et on ne l'attend pas
 * (en pratique on `await` quand même pour catcher proprement les erreurs,
 * mais l'appelant peut ne pas attendre).
 *
 * NE jette JAMAIS — toute erreur réseau / 5xx est silencieusement avalée
 * et loggée en `console.warn`. Une notification ratée NE DOIT PAS faire
 * échouer l'affectation côté UI.
 *
 * Court-circuit défensif :
 *   - email vide → no-op (sans message d'erreur, c'est une situation
 *     normale ex: pas de personne sélectionnée)
 *
 * @example
 * notifyAffectation({
 *   type: 'anomalie',
 *   email: 'jdoe@afrilandfirstbank.com',
 *   subject: 'T-46 — Eclatement de pneu du véhicule d\'agence',
 *   details: 'Délai : 3 jour(s). Commentaire : Merci de vérifier au plus vite.',
 * })
 */
export async function notifyAffectation(input: NotifyAffectationInput): Promise<void> {
  // Court-circuit : pas d'email destinataire → rien à notifier
  if (!input.email || !input.email.trim()) return

  try {
    await fetch(TEAMS_NOTIFY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: input.type,
        email: input.email.trim(),
        subject: input.subject ?? '',
        details: input.details ?? '',
        appUrl: input.appUrl ?? '',
      }),
    })
  } catch (err) {
    // Erreur réseau ou autre — on log mais on ne perturbe pas l'utilisateur.
    // Sa sauvegarde SharePoint a déjà réussi à ce stade ; la notification
    // Teams est juste un nice-to-have.
    console.warn('notifyAffectation: échec envoi notification Teams', err)
  }
}
