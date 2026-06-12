# ReportingDCPO

> Application interne **Afriland First Bank Cameroun** — Département Contrôle
> Permanent et Opérationnel (**DCPO**).
>
> Plateforme React/TypeScript hébergée par **Power Apps Code Apps**, persistée
> sur **SharePoint Online**, intégrée à **Office 365** et **Power BI**.

---

## 1. À quoi sert cette app ?

ReportingDCPO outille les contrôleurs et managers de la DCPO sur 4 grands
domaines :

| Domaine | Ce que l'app permet |
|---|---|
| **Anomalies** | Déclarer / suivre / clôturer les anomalies remontées du terrain, avec ticketing, affectation, échéance, suivi disciplinaire |
| **Reporting quotidien** | Le contrôleur soumet chaque jour son rapport d'activité ; le manager le valide ou le refuse avec motif |
| **Plan de Contrôle** | Le manager planifie des activités de contrôle annuelles à fréquence variable, et les contrôleurs les **évaluent** périodiquement |
| **Plan d'Action Correctif (PAC)** | Suivi des plans d'action correctifs émis suite à un constat, avec direction concernée, responsable, KPI et clôture |

Plus :

- **Reporting transverse** par agent (charge des contrôleurs / production de signalements)
- **Fiche récapitulative** d'une anomalie clôturée (vue d'audit)
- **Tableaux de bord Power BI** embarqués (numérisation, créance hors bilan, apurement Comex)
- **Référentiels** gérés depuis l'app : utilisateurs, directions
- **Lien direct** vers l'app Power Apps "Clôture des journées comptables"

---

## 2. Stack technique

| Couche | Tech | Version cible |
|---|---|---|
| Build / dev server | **Vite** | 5.x |
| UI | **React 19** + **TypeScript** | 19.2 |
| Hébergement de l'app | **Power Apps Code Apps** (`pac code run/push`) | — |
| Persistance des données | **SharePoint Online** (10 listes — voir [docs/SHAREPOINT_LISTS.md](docs/SHAREPOINT_LISTS.md)) | — |
| Authentification | **Office 365** (transparent — fournie par Power Apps) | — |
| Recherche personne | **Microsoft Graph** via `Office365UsersService` | — |
| Upload pièces jointes | **Power Automate** (workflows HTTP) | — |
| Tableaux de bord externes | **Power BI** (mode "Publish to web" en iframe) | — |
| Lint | ESLint 9 + plugin react-hooks (règles React 19) | — |

> 📌 Pas de framework UI tiers (pas de Material UI, Chakra, Tailwind). Tout le
> CSS est custom dans `Dashboard.css` + petits fichiers par page. Volontaire :
> simplicité, taille de bundle, contrôle total.

---

## 3. Architecture en un coup d'œil

```
┌──────────────────────────────────────────────────────────────────────┐
│                       Navigateur de l'utilisateur                    │
│                                                                      │
│   ┌────────────────────────────────────────────────────────────┐    │
│   │  apps.powerapps.com/play/... (shell Power Apps)            │    │
│   │   ┌──────────────────────────────────────────────────┐     │    │
│   │   │ App React (Vite build)                           │     │    │
│   │   │   ┌────────────┐  ┌────────────┐  ┌──────────┐   │     │    │
│   │   │   │ Pages      │  │ Components │  │ Lib      │   │     │    │
│   │   │   │ /src/pages │  │ /src/      │  │ /src/lib │   │     │    │
│   │   │   └─────┬──────┘  │ components │  └────┬─────┘   │     │    │
│   │   │         │         └────────────┘       │         │     │    │
│   │   │         └────────► Services générés ◄──┘         │     │    │
│   │   │            /src/generated/services               │     │    │
│   │   └───────────────────┬──────────────────────────────┘     │    │
│   └───────────────────────┼────────────────────────────────────┘    │
│                           │                                          │
└───────────────────────────┼──────────────────────────────────────────┘
                            │
                            ▼ (SDK Power Code SDK)
        ┌───────────────────────────────────────┐
        │  SharePoint Online (10 listes DCPO_*) │
        │  Office 365 (recherche personnes)     │
        │  Power Automate (workflows upload PJ) │
        └───────────────────────────────────────┘
```

### Flux de données

1. **Authentification** : le shell Power Apps remet automatiquement le token
   de l'utilisateur connecté ; `App.tsx` interroge `MyProfile_V2` (Graph) puis
   matche l'email dans `DCPO_LISTE_USER` pour récupérer le **rôle métier**.
2. **CRUD données** : chaque page React utilise un **service métier** (`src/lib/`)
   qui encapsule un ou plusieurs **services générés** (`src/generated/services/`)
   eux-mêmes pilotés par le **SDK `@pa-client/power-code-sdk`**.
3. **Pièces jointes** : on POST un payload `{ fileName, fileContent (base64), idItem }`
   à un workflow Power Automate qui crée la PJ côté SP et renvoie l'URL.
   L'URL est ensuite ajoutée à un champ texte `urlPieceJointe` (multi-URL séparées par `|`).
4. **Filtrage / recherche** : pattern "saisie / applied" partout — l'utilisateur
   tape dans les inputs (`filter*`), rien ne se passe ; au clic sur **Rechercher**,
   on snapshot l'état dans `applied*` et c'est ce qui déclenche le filtrage
   (côté serveur via OData quand c'est possible, sinon côté client).

---

## 4. Structure du projet

```
reportingDCPO/
├── .power/                          # Configuration Power Apps (généré)
│   └── schemas/sharepointonline/    # Schémas JSON des listes SP
├── docs/
│   └── SHAREPOINT_LISTS.md          # Doc référence des colonnes SP
├── dist/                            # Build de production (généré par `npm run build`)
├── files/                           # Sources Excel / docs métier (non commit)
├── public/                          # Assets statiques
├── src/
│   ├── App.tsx                      # Auth + autorisation + bascule home/dashboard
│   ├── PowerProvider.tsx            # Wrapper SDK Power Code (init au montage)
│   ├── main.tsx                     # Entrée Vite
│   ├── App.css, index.css           # Styles racine
│   │
│   ├── components/                  # Composants réutilisables
│   │   ├── Pagination.tsx           # Barre de pagination
│   │   ├── usePagination.ts         # Hook + config globale (page size, options)
│   │   ├── ProgressBar.tsx          # Barre de progression (Plan de Contrôle)
│   │   └── UserPicker.tsx           # Autocomplete Office 365
│   │
│   ├── lib/                         # Couche métier — encapsule les services SP
│   │   ├── activityService.ts       # Rapports d'activité contrôleur (sérialisation hybride)
│   │   ├── anomalyBulletin.ts       # Bulletin consolidé d'anomalie + timeline
│   │   ├── directionService.ts      # CRUD référentiel directions
│   │   ├── formatters.ts            # Helpers nombres + dates (anti-bug TZ)
│   │   ├── pacService.ts            # PAC + évaluations PAC
│   │   ├── planControleService.ts   # Plans de contrôle + évaluations
│   │   ├── spReferenceRows.ts       # Chargement référentiels (agences, réseaux)
│   │   └── ticketAttachments.ts     # Upload Power Automate + multi-URL
│   │
│   ├── pages/                       # Une page React par menu
│   │   ├── Dashboard.tsx            # Layout principal + navigation
│   │   ├── Dashboard.css            # Styles globaux + sidebar
│   │   ├── Anomalies.tsx            # Module 1 — gestion anomalies
│   │   ├── AnomalyBulletins.tsx     # Fiche récapitulative
│   │   ├── ReportingAgent.tsx       # Reporting par agent + bulletin notation
│   │   ├── ControllerReporting.tsx  # Saisie du rapport quotidien
│   │   ├── ControllerMyReports.tsx  # Mes rapports (vue contrôleur)
│   │   ├── ControllerReportingList.tsx  # Validation reportings (vue manager)
│   │   ├── PlanControle.tsx         # Plans de contrôle + évaluations
│   │   ├── PlanActionCorrectif.tsx  # PAC + évaluations
│   │   ├── DirectionManagement.tsx  # CRUD directions
│   │   ├── UserManagement.tsx       # CRUD utilisateurs (managers)
│   │   ├── DashboardPowerBI.tsx     # Wrapper iframe Power BI générique
│   │   └── ActiviteControleur.tsx   # (placeholder réservé futur)
│   │
│   └── generated/                   # ⚠️ AUTO-GÉNÉRÉ — ne pas éditer à la main
│       ├── models/                  # Types TypeScript des listes SP
│       └── services/                # Classes wrapper SDK pour CRUD SP
│
├── scripts/                         # Scripts utilitaires (fix-services, etc.)
├── package.json
├── tsconfig.json / tsconfig.app.json
├── vite.config.ts
├── power.config.json                # Config Power Apps (appId, env, connections)
└── eslint.config.js
```

---

## 5. Rôles et permissions

3 rôles métier (champ `fonction` de `DCPO_LISTE_USER`, valeurs Choice) :

| Rôle | Périmètre |
|---|---|
| **`Directeur`** | Voit tout, fait tout. Peut simuler un rôle inférieur via le sélecteur de la topbar |
| **`Chef_Departement`** | Idem Directeur (managers). Peut simuler Contrôleur |
| **`Controleur`** | Visibilité restreinte à ses anomalies/contrôles/PAC + actions limitées |

### Tableau récap des restrictions

| Action | Manager | Contrôleur (affecté) | Contrôleur (non affecté) |
|---|---|---|---|
| Voir liste anomalies | ✅ toutes | ✅ siennes uniquement | (filtre serveur) |
| Créer anomalie | ✅ | ✅ | ✅ |
| Affecter une anomalie | ✅ | ❌ | ❌ |
| Éditer une anomalie (mode inline) | ✅ | ✅ (sienne) | ❌ |
| Changer le statut (workflow rapide) | ✅ | ❌ | ❌ |
| Clore la résolution | ✅ | ✅ (sienne) | ❌ |
| Soumettre un rapport quotidien | ❌ (ne soumet pas) | ✅ | ✅ |
| Valider / refuser un rapport | ✅ | ❌ | ❌ |
| Créer / éditer / affecter un Plan de Contrôle | ✅ | ❌ | ❌ |
| Évaluer un Plan de Contrôle | ✅ | ✅ (sien) | ❌ |
| Créer / éditer / affecter un PAC | ✅ | ❌ | ❌ |
| Évaluer un PAC (si pas figé) | ✅ | ✅ (sien) | ❌ |
| Référentiels (Directions, Users) | ✅ | ❌ | ❌ |

> ⚠️ **Sécurité** : ces restrictions sont **côté client**. Pour une vraie
> défense en profondeur, configurer aussi les permissions SharePoint
> au niveau des listes (Item-level Permissions sur `DCPO_LISTE_ANORMALIE`,
> `DCPO_LISTE_PLAN_CONTROLE`, `DCPO_LISTE_PLAN_ACTION_CORRECTIF`).

---

## 6. Démarrer en local

```bash
# Pré-requis : Power Platform CLI (pac) installé et authentifié
# https://learn.microsoft.com/en-us/power-platform/developer/cli/

# 1. Installer les dépendances
npm install

# 2. Démarrer (lance `pac code run` + Vite en parallèle)
npm run dev
```

Une fenêtre de terminal `pac code run` s'ouvre — elle **doit rester ouverte**
pendant tout le développement, c'est elle qui établit les connexions SP /
Office 365 / Power Automate.

L'app est ensuite accessible :

- **Localement** : `http://localhost:5174` (Vite dev server)
- **Via Power Apps publié** : ouvre l'app `Reporting-Version-DCPO` dans le
  portail Power Apps (le shell se connecte au `localAppUrl` configuré dans
  `power.config.json`)

### Si une liste SharePoint est ajoutée / modifiée

```bash
# Ajouter une nouvelle data source au projet
pac code add-data-source

# Si "Data source not found" au runtime → redémarrer :
# 1. Fermer la fenêtre `pac code run`
# 2. Ctrl+C dans le terminal Vite
# 3. npm run dev
# 4. Re-publier ensuite côté Power Apps Maker
```

---

## 7. Publication / déploiement

```bash
# Build prod
npm run build

# Push vers Power Apps (utilise power.config.json)
pac code push
```

Ou via le portail Power Apps Maker → **Modifier** → **Publier**.

L'app publiée est accessible à toute personne dans le tenant `2bd82a68-2c7d-4c43-b080-9b064410f6cf`,
**à condition d'être présent dans `DCPO_LISTE_USER` avec un rôle autorisé**.

---

## 8. Conventions de code clés

### Pattern de filtres

- Inputs liés à `filters` (saisie en cours)
- Au clic sur **Rechercher** → snapshot dans `appliedFilters`
- Le filtrage / refetch dépend uniquement de `appliedFilters`
- → permet à l'utilisateur de modifier 5 champs sans déclencher 5 requêtes

### Pattern "compute during render"

React 19 interdit `setState` synchrone dans `useEffect`. On utilise donc :

```ts
const [prev, setPrev] = useState(somethingFromProps)
if (prev !== somethingFromProps) {
  setPrev(somethingFromProps)
  setSomethingDerived(reset)
}
```

(cf. `DashboardPowerBI.tsx`, `usePagination.ts`)

### Pattern d'écriture des dates

```ts
// CALENDAR DAY (saisi via <input type="date">)
field_0: form.field_0 + 'T00:00:00Z'   // minuit UTC

// VRAI TIMESTAMP (action utilisateur)
dateAffection: new Date().toISOString()
```

À la lecture, utiliser `formatDateOnlyFR(value)` pour les calendar days
(évite le décalage de fuseau horaire — voir `lib/formatters.ts`).

### Multi-URL dans un champ texte

```ts
// Read
const urls = parseUrlList(item.urlPieceJointe)  // ['url1', 'url2']

// Append sans écraser
const next = appendUrl(item.urlPieceJointe, newUrl)
await service.update(id, { urlPieceJointe: next })
```

### Format Person SP

```ts
// Read
const name = item.responsable?.DisplayName
const email = item.responsable?.Email

// Write
payload.responsable = {
  '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
  Claims: toClaims(email),  // 'i:0#.f|membership|email@domain.com'
}
```

---

## 9. Documentation complémentaire

| Fichier | Contenu |
|---|---|
| **[docs/SHAREPOINT_LISTS.md](docs/SHAREPOINT_LISTS.md)** | Référence détaillée des 10 listes SP avec toutes les colonnes, types, relations |
| Header docstrings dans `src/pages/*.tsx` | Rôle métier, permissions, structure UI de chaque page |
| Header docstrings dans `src/lib/*.ts` | Mapping SP → domaine, conventions de sérialisation |

---

## 10. Workflows Power Automate utilisés

L'app appelle 4 workflows pour l'upload de pièces jointes (un par liste cible) —
URLs centralisées dans [src/lib/ticketAttachments.ts](src/lib/ticketAttachments.ts).
Si une URL est régénérée côté Power Automate, la mettre à jour à cet endroit
uniquement.

| Workflow | Liste cible |
|---|---|
| `ATTACHMENT_API_URL` | `DCPO_LISTE_ANORMALIE` |
| `ACTIVITY_ATTACHMENT_API_URL` | `DCPO_ACTIVICTE_CONTROLLER` |
| `EVALUATION_ATTACHMENT_API_URL` | `DCPO_EVALUATION_PLAN_CONTROLE` |
| (URL PAC eval) | `DCPO_EVALUATION_PLAN_ACTION_CORRECTIF` |

---

## 11. Référentiels (constantes côté code)

Certaines listes de choix sont **côté code** plutôt qu'en SP Choice column pour
permettre une évolution rapide sans toucher au schéma SP :

| Constante | Fichier | Contenu |
|---|---|---|
| `CLASSIFICATION_OPTIONS` | `Anomalies.tsx` | Classifications d'anomalie (Fraude interne, Exécution livraison, etc.) |
| `DOMAINE_ACTIVITE_OPTIONS` | `Anomalies.tsx` | Engagements / Opérations / Surveillance IT / etc. |
| `TYPE_SANCTION_OPTIONS` | `Anomalies.tsx` | Avertissement, Blâme, Suspension, etc. |
| `CRITICITE_OPTIONS` | `Anomalies.tsx` | Faible / Moyenne / Haute / Critique |
| `PLAN_CONTROLE_CATEGORIES` | `planControleService.ts` | Catégories de plan de contrôle |
| `FREQUENCE_OPTIONS` | `planControleService.ts` | Quotidienne / Hebdomadaire / Mensuelle / Annuelle |
| `STATUS_OPTIONS` | `planControleService.ts` | À planifier / Planifié / En cours / Réalisé |
| `PAC_STATUS_OPTIONS` | `pacService.ts` | En cours / Exécutée / Non Exécutée |

Pour ajouter une valeur, l'ajouter dans le tableau **et** dans la colonne Choice
SP correspondante si elle existe.

---

## 12. Pour aller plus loin

- 🚀 **Build & deploy automatisés** : pas encore de CI/CD configuré. Pour
  l'ajouter : GitHub Actions ou Azure DevOps avec `npm run build` + `pac code push`.
- 🔒 **Permissions SharePoint** : à durcir côté listes SP pour empêcher les
  contournements UI (modification directe via SharePoint native).
- 📊 **Power BI Embedded** : remplacer le mode "Publish to web" par Power BI
  Embedded + SDK `powerbi-client` pour bénéficier de l'auth utilisateur et
  contourner les CSP restrictives (cf. limitations actuelles dans
  `DashboardPowerBI.tsx`).
- ⚙️ **Code splitting** : `npm run build` warn sur la taille du bundle
  (~500 kB). Le splitting des routes via `React.lazy()` réduirait
  significativement le temps de chargement initial.

---

## Contact

Pour toute question sur le projet : se référer aux header docstrings de
chaque page / service, qui explicitent les choix de conception. Pour la
structure SP : [docs/SHAREPOINT_LISTS.md](docs/SHAREPOINT_LISTS.md).
