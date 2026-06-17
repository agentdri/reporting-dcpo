# Documentation des listes SharePoint — ReportingDCPO

> Référence technique des **10 listes SharePoint** utilisées par l'application
> ReportingDCPO (DCPO = Département Contrôle Permanent et Opérationnel — Afriland First Bank).
>
> À tenir à jour quand on ajoute/renomme une colonne en SharePoint et qu'on
> régénère les services via `pac code add-data-source`.

---

## Vue d'ensemble

| Liste SharePoint | Rôle métier | Service généré |
|---|---|---|
| **DCPO_LISTE_USER** | Référentieswl des utilisateurs autorisés + rôle métier | `DCPO_LISTE_USERService` |
| **DCPO_LISTE_AGENCES** | Référentiel des agences bancaires | `DCPO_LISTE_AGENCESService` |
| **DCPO_LISTE_RESEAUX** | Référentiel des réseaux (regroupements d'agences) | `DCPO_LISTE_RESEAUXService` |
| **DCPO_LISTE_DIRECTION** | Référentiel des directions concernées par les PAC | `DCPO_LISTE_DIRECTIONService` |
| **DCPO_LISTE_ANORMALIE** | Anomalies déclarées + ticketing + clôture | `DCPO_LISTE_ANORMALIEService` |
| **DCPO_ACTIVICTE_CONTROLLER** | Rapports d'activité quotidiens des contrôleurs | `DCPO_ACTIVICTE_CONTROLLERService` |
| **DCPO_LISTE_PLAN_CONTROLE** | Plans de contrôle planifiés sur l'année | `DCPO_LISTE_PLAN_CONTROLEService` |
| **DCPO_EVALUATION_PLAN_CONTROLE** | Évaluations périodiques d'un plan de contrôle | `DCPO_EVALUATION_PLAN_CONTROLEService` |
| **DCPO_LISTE_PLAN_ACTION_CORRECTIF** | Plans d'action correctifs (PAC) | `DCPO_LISTE_PLAN_ACTION_CORRECTIFService` |
| **DCPO_EVALUATION_PLAN_ACTION_CORRECTIF** | Suivis périodiques d'un PAC | `DCPO_EVALUATION_PLAN_ACTION_CORRECTIFService` |

### Légende des types

| Type | Sens |
|---|---|
| `string (max N)` | Texte court (single line of text), longueur max N |
| `string` | Texte long (multiple lines of text) |
| `html` | Texte long avec mise en forme HTML conservée |
| `number` / `integer` | Numérique |
| `date` / `date-time` | Date avec ou sans heure |
| `uri` | URL (texte stocké avec validation hyperlien) |
| `Personne` | Champ Person/Group SP (renvoie un objet avec Claims/DisplayName/Email) |
| `Choix` | Champ Choice SP (renvoie un objet avec Value + Id) |

### Champs système communs (présents partout)

Toutes les listes ont en plus :

| Colonne | Type | Notes |
|---|---|---|
| `ID` | integer | Identifiant unique auto-incrémenté (read-only) |
| `Author` | Personne | Créateur de l'item (read-only) |
| `Editor` | Personne | Dernier modificateur (read-only) |
| `Created` | date-time | Horodatage création (read-only) |
| `Modified` | date-time | Horodatage dernière modification (read-only) |

Ne sont PAS documentés ci-dessous pour la lisibilité.

---

## DCPO_LISTE_USER

**Rôle** : référentiel des utilisateurs autorisés à se connecter à l'application,
avec leur fonction métier. Lu au démarrage (`App.tsx`) pour matcher l'email
Office 365 et autoriser/refuser l'accès.

**Utilisation dans l'app** : `App.tsx` (autorisation), `UserManagement.tsx` (CRUD).

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Titre | string (max 255) | Identifiant texte de l'utilisateur (en pratique : nom complet) |
| `nom` | nom | string (max 255) | Nom complet (DisplayName Office 365) |
| `Email` | Email | string (max 255) | Email Office 365 — sert de **clé de matching** |
| `fonction` | fonction | Choix | Rôle métier — valeurs autorisées : **`Chef_Departement`**, **`Directeur`**, **`Controleur`** |

---

## DCPO_LISTE_AGENCES

**Rôle** : référentiel des agences bancaires. Sert de source pour le select
« Agence » sur tous les formulaires anomalie.

**Utilisation dans l'app** : `Anomalies.tsx`, `AnomalyBulletins.tsx`,
`ReportingAgent.tsx` — chargée via `loadAgences()` dans `spReferenceRows.ts`.

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Title | string (max 255) | Nom affiché de l'agence (ex: « First Bank Bessengue ») |
| `field_1` | reseau_id | number | FK vers `DCPO_LISTE_RESEAUX.ID` — auto-déduit le réseau quand l'utilisateur choisit l'agence |

---

## DCPO_LISTE_RESEAUX

**Rôle** : référentiel des regroupements d'agences (zones commerciales).

**Utilisation dans l'app** : sélecteur en lecture seule (auto-rempli depuis
l'agence choisie) sur tous les formulaires anomalie. Chargé via
`loadReseaux()`.

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Title | string (max 255) | Code court / identifiant interne |
| `field_1` | nom_reseau | string (max 255) | Nom affiché du réseau (ex: « YAOUNDE-SUD », « DOUALA-EST ») |

---

## DCPO_LISTE_DIRECTION

**Rôle** : référentiel des directions de la banque, utilisé pour les directions
concernées par un PAC (un PAC peut concerner plusieurs directions).

**Utilisation dans l'app** : `DirectionManagement.tsx` (CRUD), formulaire de
création PAC.

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Titre | string (max 255) | Copie du `sigle` (pour affichage natif SP) |
| `libelle` | libelle | string (max 255) | Libellé complet (ex: « Direction des Systèmes d'Information ») |
| `sigle` | sigle | string (max 255) | Code court (ex: `DSI`, `DJC`, `DCE`) — **clé fonctionnelle** utilisée dans les PAC |

---

## DCPO_LISTE_ANORMALIE

**Rôle** : entité centrale de l'application. Stocke chaque anomalie déclarée
avec son cycle de vie complet (création → affectation → traitement → clôture).

**Utilisation dans l'app** : `Anomalies.tsx` (CRUD principal),
`AnomalyBulletins.tsx` (fiche récapitulative des résolues/closes),
`ReportingAgent.tsx` (agrégation par agent).

### Identification et classification

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Title | string (max 255) | Titre court de l'anomalie |
| `field_0` | date | date-time | Date métier de l'anomalie (saisie utilisateur, ≠ Created) |
| `field_3` | unite | string (max 255) | Unité organisationnelle concernée |
| `field_4` | cause | html | Description longue + concaténation des sections (cause immédiate, racine, observations) — voir parser `parseDescriptionSections` |
| `field_5` | classification | string (max 255) | Classification (« Fraude interne », « Exécution livraison... », etc. — référentiel **`CLASSIFICATION_OPTIONS`** dans `Anomalies.tsx`) |
| `field_6` | agence | string (max 255) | FK texte vers `DCPO_LISTE_AGENCES.ID` |
| `field_7` | reseau | string (max 255) | FK texte vers `DCPO_LISTE_RESEAUX.ID` (auto-déduit de l'agence) |
| `field_8` | montant | number | Montant financier en FCFA |
| `criticiteAnomalie` | criticite | string (max 255) | Faible / Moyenne / Haute / Critique |
| `domaineActivite` | domaineActivite | string (max 255) | Domaine fonctionnel (Engagements, Opérations digitales, Surveillance IT, etc. — référentiel **`DOMAINE_ACTIVITE_OPTIONS`**) |

### Cycle de vie / dates

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `dateOuvertureTicket` | date_ouverture_ticket | date | Horodatage automatique à la création (= ouverture du ticket) |
| `field_9` | date_de_regularisation | date-time | Date de régularisation (saisie à la résolution) |
| `date_cloture_ticket` | date_cloture_ticket | date | Horodatage automatique à la clôture |
| `field_10` | statut | string (max 255) | Statut workflow : **`Ouvert`**, **`En cours`**, **`Resolu`**, **`Clos`** |

### Affectation et acteurs

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `declarant_anormalie` | declarant_anormalie | Personne | Qui a déclaré l'anomalie (souvent celui qui crée le ticket) |
| `auteur_anormalie` | auteur_anormalie | Personne | Auteur "métier" de l'anomalie (peut être modifié à la clôture) |
| `personneAffecter` | personneAffecter | Personne | **Contrôleur assigné** au traitement — pilote la visibilité côté contrôleur |
| `dateAffection` | dateAffection | string (max 255) | Horodatage de l'affectation (ISO) |
| `delai` | delai | string (max 255) | Délai de traitement en **jours** (nombre stocké en string) — défaut 3 j |
| `commentaireAffectation` | commentaireAffectation | string (max 255) | Instructions du manager à la personne affectée |

### Résolution / clôture

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `actionsMenees` | actionsMenees | string | Actions menées pour résoudre — saisi à la clôture (textarea, retours à la ligne préservés) |
| `typeSanction` | typeSanction | string (max 255) | Mode de traitement appliqué — référentiel **`TYPE_SANCTION_OPTIONS`** dans `Anomalies.tsx` (Avertissement, Blâme, Suspension, etc.) |

### Pièces jointes

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `urlPieceJointe` | urlPieceJointe | uri | URLs **multi-PJ** concaténées par ` \| ` (cf. `parseUrlList` / `appendUrl` dans `ticketAttachments.ts`) |

---

## DCPO_ACTIVICTE_CONTROLLER

**Rôle** : rapports d'activité quotidiens soumis par les contrôleurs.
Chaque contrôleur déclare ses actions de la journée ; un manager valide/refuse.

**Utilisation dans l'app** : `ControllerReporting.tsx` (saisie),
`ControllerMyReports.tsx` (mes rapports), `ControllerReportingList.tsx`
(validation manager).

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Titre | string (max 255) | Titre auto-généré : « Rapport {date} - {nom contrôleur} » |
| `controlleur` | controlleur | Personne | Auteur du rapport |
| `Date` | Date | date | Date du jour rapporté |
| `actionDeLaJournee` | actionDeLaJournee | string | Format **hybride** : liste humaine + bloc JSON encadré par `<!--LINES_JSON ... -->` pour le round-trip machine (cf. `serializeLines` / `parseLines` dans `activityService.ts`) |
| `nombreAnomalieDetectee` | nombreAnomalieDetectee | number | Nombre d'anomalies détectées dans la journée |
| `observationsGlobales` | observationsGlobales | string | Observations du contrôleur sur sa journée |
| `urlPieceJointes` | urlPieceJointes | uri | Pièce jointe optionnelle (rapport PDF, photos, etc.) |
| `statutValidation` | statutValidation | string (max 255) | Workflow validation manager : **`Soumis`** (en attente) / **`Valider`** / **`Refuser`** |
| `motifRejet` | motifRejet | string (max 255) | Motif obligatoire en cas de refus (vidé automatiquement si re-validation) |

---

## DCPO_LISTE_PLAN_CONTROLE

**Rôle** : plans de contrôle planifiés sur l'année par les managers DCPO.
Chaque plan a une fréquence d'exécution (quotidienne / hebdomadaire / etc.).

**Utilisation dans l'app** : `PlanControle.tsx` (CRUD + affectation +
évaluation), service `planControleService.ts`.

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Title | string (max 255) | Libellé du contrôle |
| `field_1` | categorie | string (max 255) | Catégorie (regroupement thématique — cf. `PLAN_CONTROLE_CATEGORIES`) |
| `field_2` | frequence | string (max 255) | Fréquence d'exécution : **`Quotidienne`**, **`Hebdomadaire`**, **`Mensuelle`**, **`Annuelle`** |
| `field_3` | annee | number | Année du plan (ex: 2026) |
| `field_5` | statutInitial | string (max 255) | Statut : **`À planifier`**, **`Planifié`**, **`En cours`**, **`Réalisé`** |
| `field_6` | objectif | string | Objectif qualitatif (texte libre) |
| `field_7` | objectifChiffre | string | Objectif chiffré (KPI cible) |
| `responsable` | responsable | Personne | Contrôleur responsable de l'exécution |
| `urlPieceJointe` | urlPieceJointe | uri | Pièce jointe (note de cadrage, procédure, etc.) |

---

## DCPO_EVALUATION_PLAN_CONTROLE

**Rôle** : trace l'**exécution périodique** d'un plan de contrôle. Une
évaluation = une période effectuée (1 jour si Quotidien, 1 semaine si
Hebdomadaire, etc.). Le ratio évaluations réalisées / attendues sert au calcul
du taux d'évolution affiché en progress bar.

**Utilisation dans l'app** : modale Évaluation dans `PlanControle.tsx`, helper
`createEvaluation` / `listEvaluationsForControle` dans `planControleService.ts`.

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Titre | string (max 255) | Auto-généré : « Évaluation {période} » |
| `plan_controle_id` | plan_controle_id | number | **FK vers `DCPO_LISTE_PLAN_CONTROLE.ID`** — lien sémantique (pas un Lookup SP) |
| `periode` | periode | string (max 255) | Format dépendant de la fréquence du contrôle parent : `YYYY-MM-DD` (Quotidienne), `YYYY-Www` (Hebdomadaire ISO), `YYYY-MM` (Mensuelle), `YYYY` (Annuelle) |
| `observations` | observations | string | Constat, anomalies relevées, points d'attention |
| `urlPieceJointe` | urlPieceJointe | uri | Preuve de l'évaluation (PDF, screenshot, etc.) |

---

## DCPO_LISTE_PLAN_ACTION_CORRECTIF

**Rôle** : plans d'action correctifs (PAC) émis par la DCPO suite à un constat.
Structure inspirée du tableau de bord Excel « PAC DCPO ».

**Utilisation dans l'app** : `PlanActionCorrectif.tsx` (CRUD + affectation +
évaluation), service `pacService.ts`.

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Title | string (max 255) | Intitulé du PAC |
| `field_1` | sourcePAC | string (max 255) | Origine du PAC (ex: « Contrôle administratifs ») |
| `field_2` | dateOuverture | date-time | Date d'ouverture du PAC |
| `field_3` | descriptionProbleme | string | Description du problème observé |
| `field_4` | causeImmediate | string | Cause immédiate identifiée |
| `field_5` | causeRacine | string | Cause profonde / racine |
| `field_6` | actionsCorrectives | string | Actions correctives à mener |
| `field_7` | directionsConcernees | string | **Sigles** des directions concernées, concaténés par `;` (ex: `DSI;DJC`) — pointe vers `DCPO_LISTE_DIRECTION.sigle` |
| `field_8` | echeances | date-time | Échéance initiale |
| `field_9` | annee | number | Année du PAC |
| `field_11` | statutInitial | string (max 255) | **`En cours`**, **`Exécutée`**, **`Non Exécutée`** — un PAC en `Exécutée` ou `Non Exécutée` est figé (plus d'évaluation possible, cf. `canEvaluatePac`) |
| `field_12` | KPI | string (max 255) | Indicateur de succès attendu |
| `field_13` | ObservationsInitiale | string | Notes / contexte à la création |
| `responsableMiseEnOeuvre` | responsableMiseEnOeuvre | Personne | Responsable de l'exécution du PAC |
| `urlPiecesJointes` | urlPieceJointe | uri | Pièces jointes (multi-URL) |

---

## DCPO_EVALUATION_PLAN_ACTION_CORRECTIF

**Rôle** : point de suivi périodique d'un PAC en cours. Trace l'évolution avant
la clôture finale (passage en Exécutée / Non Exécutée).

**Utilisation dans l'app** : modale Évaluation dans `PlanActionCorrectif.tsx`,
helpers `createPacEvaluation` / `listEvaluationsForPac` dans `pacService.ts`.

| Colonne | Libellé | Type | Description |
|---|---|---|---|
| `Title` | Titre | string (max 255) | Auto-généré |
| `plan_action_correctif_id` | plan_action_correctif_id | string (max 255) | **FK vers `DCPO_LISTE_PLAN_ACTION_CORRECTIF.ID`** (stocké en string SP) |
| `observations` | observations | string (max 255) | Constat du suivi — pas de notion de période (un PAC est ponctuel, pas récurrent) |
| `urlPieceJointe` | urlPieceJointe | string (max 255) | Pièce jointe optionnelle |

---

## Relations entre les listes

```
┌──────────────────────────┐
│ DCPO_LISTE_RESEAUX       │
└──────────────────────────┘
           ▲
           │ field_1 (id réseau)
           │
┌──────────────────────────┐
│ DCPO_LISTE_AGENCES       │
└──────────────────────────┘
           ▲
           │ field_6 (id agence) + field_7 (id réseau)
           │
┌──────────────────────────┐         ┌────────────────────────────┐
│ DCPO_LISTE_ANORMALIE     │◀────────│ DCPO_LISTE_USER            │
│                          │  Person │ (auteur, déclarant,        │
│                          │  fields │  personne affectée)        │
└──────────────────────────┘         └────────────────────────────┘

┌────────────────────────────────┐
│ DCPO_LISTE_PLAN_CONTROLE       │
│ (responsable: Personne)        │◀────┐
└────────────────────────────────┘     │
           ▲                           │
           │ plan_controle_id          │
           │                           │
┌────────────────────────────────┐     │
│ DCPO_EVALUATION_PLAN_CONTROLE  │     │
└────────────────────────────────┘     │
                                       │
┌────────────────────────────────┐     │
│ DCPO_LISTE_DIRECTION (sigles)  │     │
└────────────────────────────────┘     │
           ▲                           │
           │ field_7 (sigles concat ;) │
           │                           │
┌────────────────────────────────┐     │
│ DCPO_LISTE_PLAN_ACTION_CORRECTIF│◀───┤ Person fields → DCPO_LISTE_USER
└────────────────────────────────┘     │
           ▲                           │
           │ plan_action_correctif_id  │
           │                           │
┌────────────────────────────────┐     │
│ DCPO_EVALUATION_PLAN_ACTION_   │     │
│ CORRECTIF                      │     │
└────────────────────────────────┘     │
                                       │
┌────────────────────────────────┐     │
│ DCPO_ACTIVICTE_CONTROLLER      │◀────┘
│ (controlleur: Personne)        │
└────────────────────────────────┘
```

---

## Conventions et points d'attention

### Liens entre listes

Les listes ont des **FK sémantiques** (numéros stockés en string ou nombre),
**pas des Lookup SP natifs**. Cela évite de complexifier la régénération du
schéma quand on change une liste, et permet le multi-valeurs (PAC →
plusieurs directions via `;`).

### Champ Person

Côté **lecture** SP : objet `{ DisplayName, Email, Claims, ... }`.
Côté **écriture** : objet `{ '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser', Claims: 'i:0#.f|membership|email' }` (cf. helper `toClaims` dans chaque service).

### Multi-URL dans une colonne URI

Plusieurs colonnes (`urlPieceJointe`, `urlPieceJointes`) stockent **plusieurs URLs**
concaténées par ` | ` (cf. `appendUrl` / `parseUrlList` dans
`ticketAttachments.ts`). C'est un workaround pour ne pas avoir à créer des
listes filles juste pour les attachments.

### Dates "calendar day" vs timestamps

| Cas | Stockage | Lecture |
|---|---|---|
| Date saisie via `<input type="date">` | `YYYY-MM-DDT00:00:00Z` (minuit UTC) | Utiliser `formatDateOnlyFR` pour éviter le décalage TZ |
| Vrai instant (`Created`, `Modified`, `dateOuvertureTicket`) | ISO complet avec heure | `new Date(value).toLocaleString('fr-FR')` direct |

### Mapping field_N opaque

Les colonnes `field_N` ont un nom interne **opaque** (numéro auto-attribué par
SP au moment de la création). Le mapping `field_N → sens métier` est documenté
ci-dessus et reproduit dans le header docstring des services
(`planControleService.ts`, `pacService.ts`).

> ⚠️ Si tu **renommes** une colonne `field_N` côté SP (ce qui ne change pas son
> nom interne) sans la recréer, le code continue de fonctionner. Si tu la
> **supprimes/recrées**, son nouveau nom interne sera différent (`field_M`)
> et il faut adapter le code en conséquence.

---

## Mise à jour de cette doc

Cette documentation est statique. Quand une colonne est ajoutée/renommée :

1. Ajouter/supprimer la colonne côté SP (SharePoint Online)
2. Régénérer les services : `pac code add-data-source` (ou redémarrer `npm run dev`)
3. Mettre à jour ce fichier dans la section concernée
4. Mettre à jour les commentaires dans le service / page React qui utilise la colonne

Pour vérifier la source de vérité d'un champ : ouvrir le fichier de schéma JSON
correspondant dans `.power/schemas/sharepointonline/` — c'est ce que le SDK
Power Apps utilise au runtime.
