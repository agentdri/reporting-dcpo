/**
 * ============================================================================
 * RÉFÉRENTIELS MÉTIER PARTAGÉS
 * ============================================================================
 *
 * Constantes de valeurs autorisées utilisées dans PLUSIEURS pages/services.
 * Centralisées ici pour éviter la duplication et garantir que la même
 * valeur s'écrit toujours à l'identique côté SharePoint.
 *
 * Pour ajouter une nouvelle valeur :
 *   1. L'ajouter dans le tableau ci-dessous
 *   2. La déclarer aussi dans la colonne Choice côté SharePoint
 *      (sinon SP rejettera l'écriture)
 *
 * Les référentiels propres à UN seul module restent dans le fichier de ce
 * module (ex: CLASSIFICATION_OPTIONS / TYPE_SANCTION_OPTIONS dans
 * Anomalies.tsx, FREQUENCE_OPTIONS dans planControleService.ts).
 * ============================================================================
 */

/**
 * Domaines d'activité — référentiel DCPO 2026.
 *
 * Utilisé par :
 *   - `Anomalies.tsx` (champ `domaineActivite` d'une anomalie)
 *   - `PlanControle.tsx` (champ `natureActivicte` d'un plan de contrôle —
 *     même référentiel à la demande métier pour aligner les axes de reporting)
 */
export const DOMAINE_ACTIVITE_OPTIONS = [
  'Engagements',
  'Exploitation et Reseau',
  'Opérations internationales',
  'Opérations digitales',
  'Administratif et Financier',
  'Surveillance IT',
] as const

/** Type union des valeurs autorisées (utile pour typer un select). */
export type DomaineActivite = typeof DOMAINE_ACTIVITE_OPTIONS[number]

/**
 * Mode de traitement appliqué à une anomalie à sa clôture
 * (champ SP `typeSanction`).
 *
 * Liste mixte action (Relance, Demande, Formation…) et sanction
 * (Avertissement, Blâme, Suspension…). Ordonnée du plus léger au plus
 * lourd pour aider à la sélection.
 *
 * Utilisé par :
 *   - `Anomalies.tsx` (formulaire de clôture)
 *   - `AnomalyBulletins.tsx` (modale d'édition Directeur)
 */
export const TYPE_SANCTION_OPTIONS = [
  'Relance outlook',
  'Demande d\'informations',
  'Demande d\'explications',
  'Ultime relance',
  'Mise en garde',
  'Avertissement',
  'Blâme',
  'Suspension',
  'Teams',
  'Produit de controles',
  'Demande de régularisation',
  'Lettre d\'observation',
  'Mise a pied 3 Jours',
  'Mise a pied 8 Jours',
  'Licenciement',
  'Formations',
] as const
