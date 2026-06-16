# Fiche de recette — Application ReportingDCPO

---

## 1. Description du projet

### Identité

- **Nom** : ReportingDCPO
- **Commanditaire** : Direction du Contrôle Permanent et Opérationnel (DCPO) — Afriland First Bank Cameroun
- **Stack** : React 19 + TypeScript + Vite, hébergée par **Power Apps Code Apps**
- **Persistance** : 10 listes **SharePoint Online**, authentification **Office 365**, workflows **Power Automate** pour les pièces jointes, **Power BI** pour les tableaux de bord embarqués

### Objectif

Outiller les **contrôleurs** et **managers** (Chef de département, Directeur) de la DCPO pour gérer leur activité quotidienne sur 4 grands domaines :

1. **Anomalies** : remontée, ticketing, affectation à un contrôleur, traitement et clôture des anomalies détectées sur le terrain.
2. **Reporting quotidien** : les contrôleurs déclarent leur activité journalière (actions menées, durées, observations) ; les managers valident ou refusent avec motif.
3. **Plan de Contrôle** : planification annuelle d'activités de contrôle (catégorie, fréquence, responsable, KPI), avec exécution périodique sous forme d'évaluations.
4. **Plan d'Action Correctif (PAC)** : suivi des plans d'action correctifs émis suite à un constat, avec direction concernée, responsable de mise en œuvre, KPI et suivi périodique.

### Rôles métier

| Rôle | Permissions principales |
|---|---|
| **Directeur** | Voit tout, fait tout. Peut simuler un rôle inférieur via le sélecteur de la topbar |
| **Chef de département** | Idem Directeur (managers) |
| **Contrôleur** | Voit uniquement ses items affectés, soumet ses rapports quotidiens, évalue ses contrôles et PAC, ne peut pas affecter ni changer les statuts |

### Architecture

```
Navigateur ▶ shell Power Apps (apps.powerapps.com) ▶ App React (Vite)
                                                        │
                                                        ▼
                                  SharePoint Online (10 listes DCPO_*)
                                  Office 365 (recherche personne)
                                  Power Automate (upload pièces jointes)
                                  Power BI (iframes embarqués)
```

### Listes SharePoint (10)

| Liste | Rôle |
|---|---|
| DCPO_LISTE_USER | Référentiel utilisateurs + rôles |
| DCPO_LISTE_AGENCES | Référentiel agences |
| DCPO_LISTE_RESEAUX | Référentiel réseaux |
| DCPO_LISTE_DIRECTION | Référentiel directions concernées par les PAC |
| DCPO_LISTE_ANORMALIE | Anomalies déclarées (entité centrale) |
| DCPO_ACTIVICTE_CONTROLLER | Rapports d'activité quotidiens |
| DCPO_LISTE_PLAN_CONTROLE | Plans de contrôle annuels |
| DCPO_EVALUATION_PLAN_CONTROLE | Exécutions périodiques d'un plan de contrôle |
| DCPO_LISTE_PLAN_ACTION_CORRECTIF | Plans d'action correctifs |
| DCPO_EVALUATION_PLAN_ACTION_CORRECTIF | Suivis périodiques d'un PAC |

---

## 2. Évaluation de la couverture fonctionnelle

### 2.1 Authentification & Permissions

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Connexion Office 365 | À l'ouverture de l'app, récupération automatique de l'identité Office 365 de l'utilisateur (DisplayName + Email + jobTitle) | L'utilisateur est identifié sans saisir d'identifiants ; ses informations apparaissent dans la carte de profil |
| Autorisation par liste DCPO_LISTE_USER | Vérification que l'email Office 365 figure dans la liste avec un rôle autorisé (Chef_Departement / Directeur / Controleur) | Si oui : accès au dashboard. Sinon : message « Vous n'avez pas les droits d'accès au dashboard » |
| Affichage du rôle métier | Le rôle réel est affiché dans la topbar (en majuscules) | Ex : « CHEF DE DÉPARTEMENT » visible en permanence |
| Simulation de rôle (managers uniquement) | Sélecteur dans la topbar permettant à un Directeur ou Chef de département de simuler un rôle inférieur | Le menu et les permissions s'adaptent immédiatement au rôle simulé ; badge « rôle simulé » jaune affiché |
| Retour au rôle réel | Sélectionner « (mon rôle) » dans le sélecteur | Le rôle effectif revient au rôle réel ; le badge disparaît |
| Restriction d'écriture par rôle | Les actions sensibles (affectation, validation, CRUD référentiels) sont masquées pour les contrôleurs | Les boutons concernés n'apparaissent pas dans l'UI pour un contrôleur |

### 2.2 Anomalies en cours (module 1 — liste principale)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Affichage de la liste des anomalies actives | Tableau paginé des anomalies dont le statut n'est ni Resolu ni Clos | Seules les anomalies en Ouvert / En cours s'affichent ; les Clos / Résolu sont masquées (visibles sur la page « Anomalies résolues ») |
| Numéro de ticket cohérent | Chaque anomalie est affichée avec son numéro T-{ID SharePoint} | Le même numéro est utilisé partout dans l'app (liste, modale, fiche récap, ticket) |
| Filtre par dates (du / au) | Filtres serveur OData sur le champ Created | Seules les anomalies créées dans la période s'affichent ; pattern saisie/applied (clic Rechercher pour appliquer) |
| Filtre par agence | Select alimenté depuis DCPO_LISTE_AGENCES | Filtre serveur OData |
| Filtre par réseau | Select alimenté depuis DCPO_LISTE_RESEAUX | Filtre serveur OData |
| Filtre par classification | Select avec valeurs « Fraude interne », « Exécution livraison… », « Fraude externe », « Interruptions de l'activité… », « Pratiques en matière d'emploi… », « Clients, produits… », « Dommages occasionnés… » | Filtre serveur OData |
| Filtre par criticité | Select Faible / Moyenne / Haute / Critique | Filtre serveur OData |
| Filtre par domaine d'activité | Select Engagements / Exploitation et Reseau / Opérations internationales / Opérations digitales / Administratif et Financier / Surveillance IT | Filtre serveur OData |
| Filtre par statut (Ouvert / En cours) | Select limité aux statuts non clôturés | Filtre serveur OData |
| Filtre par agent (auteur) | Recherche texte sur le nom / email de l'auteur | Filtrage client après chargement |
| Filtre par personne affectée | Recherche texte sur le nom / email du contrôleur affecté | Filtrage client |
| Bouton Réinitialiser | Remet tous les filtres à vide | Toutes les anomalies en cours réapparaissent |
| Restriction de visibilité contrôleur | Un contrôleur ne voit que les anomalies dont il est affecté (filtre OData sur personneAffecter/Email) | Les autres anomalies ne s'affichent pas dans son tableau |
| Tableau dépliable | Bouton +/- permet d'afficher/masquer les colonnes secondaires (déclarant, auteur, affecté, cause, classification, agence, réseau, montant, date régularisation) | Les colonnes apparaissent ou disparaissent en un clic |
| Colonnes sticky | Numéro à gauche et Actions à droite restent visibles lors du scroll horizontal | Repère visuel maintenu pendant le scroll |
| Pastille « Délai » | Affichage coloré du temps restant ou dépassé (date d'affectation + délai jours vs aujourd'hui) | Vert = en cours, orange = échéance jour J, rouge = dépassée |
| Création d'une anomalie | Formulaire structuré en 5 sections (Identité / Caractérisation / Localisation / Impact / Pièce jointe) | Une nouvelle anomalie apparaît en haut du tableau, statut Ouvert |
| Pré-remplissage Déclarant | Le déclarant est auto-rempli avec l'utilisateur connecté (read-only) | Pas de saisie manuelle, cohérence garantie |
| Picker Office 365 pour Auteur | Recherche autocomplete sur les utilisateurs O365 | Le DisplayName et l'email s'affichent ; sélection en un clic |
| Picker Office 365 pour Personne affectée (managers uniquement) | Idem auteur — réservé Chef_Departement / Directeur | Le champ est masqué pour les contrôleurs |
| Affectation automatique à la création | Si une personne est affectée à la création → dateAffection et délai (3 jours par défaut) sont automatiquement renseignés | Le contrôleur affecté voit immédiatement la nouvelle anomalie dans sa liste |
| Auto-déduction du réseau | Sélection d'une agence → le réseau parent est auto-rempli en lecture seule | Cohérence garantie agence / réseau |
| Upload de pièce jointe | Workflow Power Automate qui encode le fichier en base64 et l'attache à l'item SP | Le fichier est visible dans le détail ; URL stockée dans urlPieceJointe |
| Bouton Détail (clic sur ligne) | Ouvre une modale de détail avec tous les champs de l'anomalie | Affichage complet en lecture seule |
| Bouton Affecter (managers uniquement) | Modale en 2 étapes : recherche O365 → délai (défaut 3j) + commentaire d'affectation | Le responsable, dateAffection, délai et commentaire sont mis à jour ; le contrôleur est notifié implicitement |
| Bouton Ticket | Ouvre la modale de suivi avec actions disponibles (Détail, Changer statut, Clore résolution) | Vue synthétique avec actions contextuelles |
| Bouton Changer le statut (managers uniquement) | Modale rapide pour passer Ouvert ↔ En cours | Le statut est mis à jour ; les options Resolu / Clos sont masquées (passage par la modale Clôture) |
| Bouton Clore la résolution | Formulaire complet : statut final (Resolu / Clos), dates, auteur, causes immédiate/racine, actions menées, observations, type d'action/sanction, pièce jointe | L'anomalie disparaît de la liste « Anomalies en cours » et apparaît dans « Anomalies résolues » |
| Restriction de clôture | Un contrôleur ne peut clore qu'une anomalie qui lui est affectée | Le bouton Clore n'apparaît pas pour les autres contrôleurs |
| Mode édition inline (modale Détail) | Bouton Modifier permet d'éditer tous les champs métier sauf le statut (et personne affectée pour les contrôleurs) | Les modifications sont persistées en SharePoint au clic Enregistrer |
| Restriction d'édition contrôleur | Un contrôleur ne peut éditer que les anomalies qui lui sont affectées | Le bouton Modifier n'apparaît pas sur les autres |

### 2.3 Anomalies résolues (Fiche récapitulatif)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Affichage des anomalies clôturées | Cards des anomalies en statut Resolu ou Clos uniquement | Filtre serveur OData ; chargement initial via getAll filtré |
| Stats globales | Compteurs Total / Résolues / Closes + délai moyen + montant cumulé | Mise à jour automatique selon les filtres appliqués |
| Filtres avancés | Statut, classification, criticité, agence, agent, personne affectée, période de clôture, recherche libre | Pattern saisie/applied ; pagination compatible |
| Card cliquable | Clic ouvre la modale de fiche détaillée | Tous les détails affichés en mode read-only |
| Modale fiche détaillée | Sections : Identification, Dates clés, Caractérisation, Description & causes parsées, Cycle de vie (timeline), Actions à mener, Pièces jointes | Affichage structuré et complet |
| Parsing automatique des causes | Si les colonnes dédiées (causeImmediate / causeRacine / etc.) sont vides, extraction depuis la description concaténée | Les sections Cause immédiate / Cause racine / Observations s'affichent correctement même sur les anciennes anomalies |
| Affichage typeSanction | Type d'action / sanction décidée à la clôture | Visible dans la section Caractérisation |
| Délai calculé (audit) | Différence en jours entre date de déclaration et date de clôture | Affichage du nombre de jours réels de traitement |
| Bouton Imprimer | Lance window.print() ; les éléments hors fiche sont masqués via @media print | Fiche prête pour PDF ou impression papier |
| Bouton SharePoint | Ouvre l'item dans l'UI SharePoint native | Utile pour accéder aux pièces jointes natives ou aux versions |
| Pagination | Pagination configurable (taille de page commune à toute l'app, défaut 50) | Comportement homogène avec les autres pages |

### 2.4 Reporting par Agent

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Sélecteur de regroupement | Toggle entre « Par personne affectée » (charge des contrôleurs) et « Par auteur » (production de signalements) | L'agrégation et le bulletin imprimable s'adaptent au mode choisi |
| Vue principale (tableau d'agents) | Stats agrégées par agent : Total, Ouvert, En cours, Résolu, Clos, Montant total | Tri par volume décroissant |
| Colonnes Domaines d'activité et Sanctions | Chips colorées listant les valeurs distinctes rencontrées sur les anomalies de chaque agent | Lecture rapide des axes d'activité |
| Filtres globaux | Période, agence, réseau, classification, criticité, statut, agent, personne affectée | Mise à jour des stats et de la liste détaillée |
| Drill-down agent (clic ligne) | Affiche la vue détail avec toutes les anomalies de l'agent sélectionné | Stats individuelles + tableau anomalies de l'agent |
| Filtres détail | Filtres propres à la vue agent : dates, agence, réseau, classification, criticité, statut, personne affectée, recherche cause | Affinement de la sélection sur l'agent courant |
| Bouton Imprimer le bulletin | Génère une Fiche de notation de l'agent (titre adapté au mode de regroupement) | Sections Anomalies / Plans de Contrôle / PAC avec compteurs Prévu / Réalisé / Écart |
| Retour à la liste | Bouton qui referme la vue détail | Retour à la table des agents |

### 2.5 Saisie du rapport quotidien (contrôleurs)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Saisie multi-lignes d'actions | Pour chaque action : domaine, objet, action, durée (min), observation | Au moins une ligne requise ; bouton Ajouter une ligne |
| Calcul automatique des totaux | Total heures, % temps occupé (sur 8h), statut journée (Calme / Normal / Chargé / Très chargé) | Recalcul instantané à chaque modification |
| Sauvegarde automatique en brouillon | Persisté en localStorage avec debounce 600 ms | Restauré automatiquement au prochain montage avec badge « Brouillon restauré » |
| Upload de pièces jointes | Workflow Power Automate (peut envoyer plusieurs fichiers) | URLs stockées dans urlPieceJointes (multi-URL concaténées) |
| Validation des lignes | Domaine, objet, action (3-500 chars), durée (5-600 min) | Affichage des erreurs sous chaque champ invalide |
| Soumission du rapport | Création SharePoint avec statut initial 'Soumis' | Le rapport apparaît dans la file de validation du manager |
| Anti-doublon par jour | Un seul rapport par couple (contrôleur, date) — lookup serveur avant soumission | Si un rapport existe déjà : bannière contextuelle + bouton submit désactivé |
| Bannière « Soumis » | Information bleue si un rapport est déjà en attente de validation | Champs verrouillés ; impossible de re-soumettre |
| Bannière « Validé » | Information verte si le rapport a déjà été validé | Champs verrouillés ; modification définitivement impossible |
| Bannière « Rejeté » avec motif | Information orange affichant le motif du manager | Formulaire pré-rempli avec l'ancien contenu ; bouton renommé « Re-soumettre le rapport corrigé » |
| Re-soumission après rejet | Met à jour le contenu + remet statut à 'Soumis' + vide motif rejet (dans une seule requête) | Le rapport repart en file d'attente du manager |
| Date de rapport toujours modifiable | Le champ Date reste actif même quand le formulaire est verrouillé | L'utilisateur peut naviguer entre les jours sans réinitialiser manuellement |
| Pré-chargement via Mes rapports | Clic « Modifier » sur un rapport rejeté → redirection vers Saisie avec la date pré-chargée | Bannière orange + formulaire pré-rempli automatiquement |

### 2.6 Mes rapports (contrôleurs)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Liste des rapports du contrôleur | Filtrage côté client sur l'email du contrôleur connecté | Seuls les rapports de l'utilisateur s'affichent |
| Stats cards | Compteurs : Total, En attente (Soumis), Validés, Refusés, Total heures | Recalculés sur la liste filtrée |
| Filtre par période | Par défaut les 30 derniers jours | Inputs Date du / Date au |
| Filtre par statut | Soumis / Validé / Refusé / Tous | Select dédié |
| Recherche libre | Mot-clé dans observations + actions | Filtrage en temps réel |
| Pagination | Défaut 50 par page, configurable | Tableau scrollable horizontal |
| Modale détail (lecture seule) | Clic sur ligne → tous les détails affichés | Aucune action de validation possible (page contrôleur) |
| Décision manager affichée | Statut courant + motif si rejet | Bloc visible dans le détail |
| Bouton Modifier et re-soumettre | Visible uniquement sur les rapports en statut Refuser | Redirection vers la page Saisie avec la date pré-chargée |

### 2.7 Validation reportings (managers)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Liste consolidée des rapports | Regroupés par couple (contrôleur, date) | Une carte par groupe avec compteurs pending/validés/refusés |
| Stats globales | Compteurs : Total, En attente, Validés, Refusés, Total heures | Mise à jour selon les filtres |
| Filtres | Nom/email contrôleur, période, « en attente uniquement » | Pattern saisie/applied |
| Colonne Synthèse des actions | Aperçu compact : 2 premières actions tronquées + nombre total + bouton « Voir plus → » | Aide à la décision sans ouvrir chaque rapport |
| Modale détail | Résumé + lignes + observations + pièces jointes + décision actuelle si déjà tranchée | Affichage complet |
| Validation (Valider) | Bouton vert ; passage à statut 'Valider' + vide motif rejet | Modale se ferme automatiquement après succès |
| Invalidation (Refuser) | Bouton rouge ; passage à statut 'Refuser' avec motif obligatoire | Le textarea Motif doit être rempli, sinon alerte |
| Fermeture automatique de la modale | Après validation ou invalidation réussie | Le manager enchaîne sur le rapport suivant |
| Pagination | Par groupe (1 carte = 1 contrôleur+date) | Configurable |

### 2.8 Plan de Contrôle

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Liste des plans de contrôle | Tableau paginé avec colonnes : libellé, catégorie, fréquence, responsable, année, statut, taux d'évolution | Triable et filtrable |
| Stats cards | Compteurs par statut : Total, À planifier, Planifié, En cours, Réalisé | Mise à jour automatique |
| Filtres | Catégorie, nature d'activité, fréquence, responsable, statut, année, recherche | Pattern saisie/applied |
| Visibilité contrôleur | Un contrôleur ne voit que les contrôles dont il est responsable | Filtre client par responsableEmail |
| Colonne Taux d'évolution | Barre de progression affichant le ratio évaluations réalisées / attendues par fréquence | Calcul dynamique : Quotidienne=365, Hebdo=52, Mensuelle=12, Annuelle=1 |
| Création d'un contrôle (managers uniquement) | Formulaire : libellé, catégorie, nature d'activité, fréquence, responsable (UserPicker), année, statut, objectif qualitatif, objectif chiffré, pièce jointe | Le contrôle apparaît dans la liste |
| Édition d'un contrôle (managers uniquement) | Bouton « ✎ Modifier » dans la modale Détail | Mise à jour SharePoint via updateControle |
| Affectation rapide (managers uniquement) | Modale dédiée pour changer le responsable sans ouvrir le formulaire complet | UserPicker pré-rempli ; mise à jour du seul champ Person |
| Évaluation d'un contrôle | Modale en 2 parties : formulaire (période + observations + pièce jointe) + historique des évaluations précédentes | Création d'un item DCPO_EVALUATION_PLAN_CONTROLE lié au contrôle parent |
| Picker de période adapté à la fréquence | Quotidienne → input date, Hebdo → input week, Mensuelle → input month, Annuelle → input number | Le format est cohérent avec la fréquence du contrôle parent |
| Pièce jointe sur évaluation | Upload via workflow Power Automate dédié | URL stockée dans le champ urlPieceJointe de l'évaluation |
| Modale Détail | Affiche tous les champs du contrôle + cycle de vie + historique des évaluations + pièces jointes | Lecture seule + bouton Modifier si manager |
| Cycle de vie (timeline) | Création + évaluations triées chronologiquement + clôture si statut Réalisé | Affichage timeline verticale avec pastilles colorées par type |

### 2.9 Plan d'Action Correctif (PAC)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Liste des PAC | Tableau paginé avec colonnes : intitulé, source, échéance, directions concernées, responsable, statut, etc. | Filtrable et paginé |
| Stats cards | Compteurs par statut : Total, En cours, Exécutées, Non Exécutées | Mise à jour selon filtres |
| Filtres | Statut, direction, année, recherche libre | Pattern saisie/applied |
| Visibilité contrôleur | Un contrôleur ne voit que les PAC dont il est responsable de mise en œuvre | Filtre client |
| Multi-sélection des directions | Au moment de la création : dropdown chip avec ajout / retrait progressif | Plusieurs directions stockées en SP concaténées par `;` |
| Création d'un PAC (managers uniquement) | Formulaire complet : intitulé, source, date, descriptions, causes immédiate/racine, actions correctives, directions, responsable, échéance, KPI, année, statut, observations | Le PAC apparaît dans la liste |
| Édition d'un PAC (managers uniquement) | Bouton Modifier dans la modale Détail | Mise à jour SP via updatePAC |
| Affectation rapide (managers uniquement) | Modale UserPicker pour changer le responsable | Mise à jour ciblée du seul champ Person |
| Évaluation d'un PAC | Création d'un item DCPO_EVALUATION_PLAN_ACTION_CORRECTIF (observations + pièce jointe) | Trace les points de suivi avant clôture |
| Verrou métier sur PAC figé | Un PAC en statut Exécutée ou Non Exécutée ne peut plus être évalué | Le bouton Évaluation est désactivé |
| Modale Détail | Affiche tous les champs + historique des évaluations + pièces jointes | Lecture seule + bouton Modifier si manager |
| Pièces jointes multi-URL | Concaténation par `\|` dans le champ urlPieceJointe | Multiples fichiers attachables sans nouvelle colonne SP |

### 2.10 Dashboards Power BI

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Embarquement iframe | 3 dashboards Power BI (Numérisation, Créance Hors Bilan, Apurement Comex) intégrés via iframe | Le dashboard s'affiche en plein écran dans la zone de contenu |
| Spinner de chargement | Overlay pendant le chargement initial du bundle Power BI | Disparaît dès que onLoad de l'iframe se déclenche |
| Bouton Ouvrir dans un nouvel onglet | Lien externe vers l'URL Power BI publique | Ouvre dans un nouvel onglet (filet de sécurité si CSP bloque l'iframe) |
| Compatibilité Power Apps publié | Si la CSP du tenant bloque l'iframe, le bouton externe reste fonctionnel | Pas de blocage total ; l'utilisateur a toujours accès au contenu |

### 2.11 Configuration — Référentiel Directions (managers uniquement)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Liste des directions | Tableau avec Sigle, Libellé, Actions | Tri alphabétique par sigle |
| Création d'une direction | Formulaire : sigle (code court) + libellé complet | Sigle est copié dans Title pour l'UI native SP |
| Modification d'une direction | Bouton Modifier dans le tableau | Mise à jour de Title + libelle + sigle |
| Suppression d'une direction | Bouton Supprimer avec confirmation | Item supprimé en SP ; n'affecte pas les PAC qui référencent ce sigle |

### 2.12 Configuration — Gestion des utilisateurs (managers uniquement)

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Liste des utilisateurs | Tableau paginé avec nom, email, rôle | Tri alphabétique |
| Filtres | Recherche libre, filtre par rôle | Pattern saisie/applied |
| Création d'un utilisateur | UserPicker Office 365 qui auto-remplit nom + email, puis sélection du rôle (Chef_Departement / Directeur / Controleur) | Item créé en SP avec champ Choice fonction |
| Modification d'un utilisateur | Mêmes champs ; même picker | Mise à jour Title + nom + Email + fonction |
| Anti-doublon email | Vérification qu'aucun autre item n'a le même email | Refus avec message si conflit |
| Garde-fou auto | Impossible de se modifier soi-même via ce module | Comportement explicitement bloqué |

### 2.13 Lien externe — Clôture des journées comptables

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Lien direct vers l'app Power Apps | Lien `<a target="_blank">` dans le menu Dashboards | Ouverture dans un nouvel onglet sans changer la page courante de ReportingDCPO |
| Icône externe ↗ | Signale visuellement la nature externe de l'entrée | Distinction claire vs les pages internes |

### 2.14 Comportements transverses

| Fonctionnalité | Description | Résultat attendu |
|---|---|---|
| Pagination globale | Taille de page configurable (5 / 10 / 25 / 50 / 100), défaut 50 | Comportement homogène sur toutes les pages |
| Reset pagination sur changement de filtre | Retour automatique à la page 1 si les filtres changent | Pas d'état incohérent |
| Pattern saisie/applied | Les filtres ne sont appliqués qu'au clic Rechercher | Évite des requêtes serveur à chaque frappe |
| Bouton Réinitialiser sur tous les filtres | Remet tous les champs à leur valeur par défaut | Filtres effacés ET appliqués |
| Tri descendant par défaut | Les listes sont triées du plus récent au plus ancien (Created desc) | Les items récents en haut |
| Format de dates locale FR | jj/mm/aaaa pour les dates calendrier, jj/mm/aaaa HH:MM pour les timestamps | Cohérence FR sur toute l'app |
| Anti-bug fuseau horaire | Helper formatDateOnlyFR pour les dates « calendar day » stockées en UTC | Pas de décalage +1 jour selon le fuseau du navigateur |
| Format montants compact | Affichage en millions (M) au-delà d'1 M sur les cartes de stats | Pas de débordement visuel pour les gros agrégats |
| Pièces jointes multi-URL | Concaténation par `\|` dans un champ texte SP | Plusieurs fichiers sans liste fille SP |
| Sticky topbar et sidebar | La barre supérieure et la sidebar restent fixes lors du scroll de la zone de contenu | Navigation toujours accessible |
| Icônes dans la sidebar | Chaque entrée de menu a un emoji représentatif | Lecture visuelle rapide |
| Restauration de session | Brouillon (saisie rapport) restauré automatiquement à la reconnexion | L'utilisateur ne perd pas son travail en cas de fermeture inattendue |

---

## 3. Validation de recette

### Modalités

- **Plateforme de test** : application publiée sur `apps.powerapps.com/play/e/.../a/...`
- **Comptes de test recommandés** : 1 Directeur + 1 Chef de département + 2 Contrôleurs (au moins un avec des anomalies affectées)
- **Données de test** : un jeu d'au moins 10 anomalies couvrant tous les statuts, 5 plans de contrôle de fréquences variées, 5 PAC avec plusieurs directions

### Critères d'acceptation

Chaque ligne du tableau de couverture fonctionnelle constitue un critère d'acceptation. Une fonctionnalité est considérée comme **validée** si :

1. Le **résultat attendu** est obtenu sur l'environnement de production
2. Aucune **régression** n'est observée sur les fonctionnalités voisines
3. Les **permissions** par rôle sont correctement appliquées

### Signature

| Rôle | Nom | Date | Signature |
|---|---|---|---|
| Maîtrise d'ouvrage (DCPO) | | | |
| Maîtrise d'œuvre (développement) | | | |
| Validation finale | | | |
