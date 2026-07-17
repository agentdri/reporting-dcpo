# Guide utilisateur — ReportingDCPO

Application de gestion du contrôle permanent opérationnel — Direction du Contrôle Permanent Opérationnel (DCPO), Afriland First Bank Cameroun.

Version du guide : 1.0
Dernière mise à jour : 09/07/2026

---

## Table des matières

1. [Introduction](#1-introduction)
2. [Connexion et premier accès](#2-connexion-et-premier-acces)
3. [Interface générale](#3-interface-generale)
4. [Rôles et permissions](#4-roles-et-permissions)
5. [Module Anomalies — En cours](#5-module-anomalies--en-cours)
6. [Module Anomalies — Résolues](#6-module-anomalies--resolues)
7. [Module Rapport quotidien](#7-module-rapport-quotidien)
8. [Module Reporting par Agent](#8-module-reporting-par-agent)
9. [Module Plan de Contrôle](#9-module-plan-de-controle)
10. [Module Plan d'Action Correctif (PAC)](#10-module-plan-daction-correctif-pac)
11. [Dashboards Power BI](#11-dashboards-power-bi)
12. [Configuration (managers uniquement)](#12-configuration-managers-uniquement)
13. [Fonctionnalités transverses](#13-fonctionnalites-transverses)
14. [FAQ et dépannage](#14-faq-et-depannage)

---

## 1. Introduction

**ReportingDCPO** est l'application interne de la Direction du Contrôle Permanent Opérationnel d'Afriland First Bank Cameroun. Elle centralise :

- La déclaration, le suivi et la clôture des **anomalies**
- Les **rapports d'activité quotidiens** des contrôleurs
- Les **plans de contrôle** annuels et leurs évaluations
- Les **plans d'action correctifs** (PAC) issus des anomalies
- Le **reporting par agent** et les fiches de notation
- Des tableaux de bord analytiques **Power BI**

L'application est hébergée sur la plateforme **Microsoft Power Apps Code Apps**, avec SharePoint Online comme socle de données et Office 365 pour l'authentification.

Ce guide s'adresse à **tous les utilisateurs** de l'application : contrôleurs, chefs de département et directeurs.

---

## 2. Connexion et premier accès

### 2.1 Prérequis

- Un compte **Office 365 Afriland First Bank** actif
- Un navigateur récent : **Microsoft Edge** ou **Google Chrome** (recommandés)
- Une connexion internet stable

### 2.2 Se connecter

1. Cliquer sur le lien de l'application (fourni par la DSI ou disponible dans le portail Power Apps)
2. Se laisser rediriger vers la page d'authentification Microsoft
3. Saisir son adresse email professionnelle et son mot de passe
4. À la première connexion, autoriser l'application à accéder à votre profil Office 365

L'application charge alors votre profil et détermine automatiquement votre **rôle** (Directeur, Chef de département ou Contrôleur) à partir de la liste des utilisateurs configurée par l'administrateur.

### 2.3 Si votre rôle est manquant

Si le message « Rôle non défini » apparaît, contacter le manager du service pour être ajouté à la liste des utilisateurs autorisés. Sans rôle, la navigation reste possible mais les fonctions d'écriture sont bloquées.

---

## 3. Interface générale

### 3.1 Barre supérieure (topbar)

En haut de l'écran :

- **Logo & titre** de l'application à gauche
- **Sélecteur de rôle** (managers uniquement) : permet de simuler un rôle de niveau inférieur pour tester l'expérience utilisateur. Le rôle réel reste affiché à droite du sélecteur.
- **Nom et fonction** de l'utilisateur connecté à droite

### 3.2 Menu latéral (navigation)

À gauche, un menu vertical regroupe les modules par thématique :

| Section | Contenu |
|---|---|
| **Accueil** | Vue d'ensemble avec graphiques et indicateurs clés |
| **Anomalies** | Anomalies en cours + Anomalies résolues |
| **Reporting par Agent** | Statistiques agrégées par personne |
| **Rapport quotidien** | Saisie du rapport + Mes rapports (contrôleurs) |
| **Validation reportings** | Validation manager (managers uniquement) |
| **Plan de Contrôle** | Suivi des contrôles récurrents |
| **Plan d'Action Correctif** | Suivi des PAC |
| **Dashboards** | Numérisation, Créance Hors Bilan, Apurement Comex |
| **Configuration** | Directions, Gestion des utilisateurs (managers uniquement) |

Les sections avec un chevron sont **repliables** — cliquer sur le libellé pour ouvrir/fermer.

### 3.3 Contenu principal

La zone centrale affiche la page correspondant à l'onglet actif. Chaque page suit une structure homogène :

1. **En-tête** : titre + boutons d'action (Nouveau, Actualiser, Export…)
2. **Cartes KPI** : indicateurs chiffrés en tête (totaux, taux, montants)
3. **Barre de filtres** : critères de recherche avec bouton **Rechercher** et **Réinitialiser**
4. **Tableau ou grille** de données paginées
5. **Barre de pagination** : navigation entre les pages (taille par défaut : 50 lignes)

### 3.4 Cartes KPI et pagination

**Important** : les cartes KPI en haut de page reflètent **toutes les données filtrées**, pas uniquement la page affichée. Si vous avez 500 anomalies après filtre et que la page affiche 50 lignes, les KPI comptent bien les 500.

---

## 4. Rôles et permissions

L'application distingue **3 rôles**, chacun avec ses permissions.

### 4.1 Directeur

Le rôle le plus élevé. A **tous les droits** :

- Créer, modifier, affecter, clôturer n'importe quelle anomalie
- Modifier même les anomalies déjà résolues ou closes (voir §6.4)
- Valider ou refuser les rapports quotidiens
- Créer et modifier des plans de contrôle et PAC
- Configurer les directions et gérer les utilisateurs

### 4.2 Chef de département

Rôle manager intermédiaire. A la plupart des droits sauf certaines actions administratives :

- Créer, modifier, affecter, clôturer les anomalies
- Valider ou refuser les rapports quotidiens
- Créer et modifier plans de contrôle et PAC
- Ne peut pas modifier les anomalies déjà closes (réservé au Directeur)

### 4.3 Contrôleur

Rôle opérationnel de terrain. Périmètre restreint :

- Consulter toutes les anomalies (en cours et résolues)
- Créer une anomalie
- Modifier une anomalie **si** il en est le déclarant ou la personne affectée
- Clôturer une anomalie **si** il en est la personne affectée
- Saisir ses rapports quotidiens
- Consulter ses propres rapports (Mes rapports)
- Consulter les plans de contrôle et PAC dont il est responsable

### 4.4 Simulation de rôle

Les managers peuvent basculer temporairement vers un rôle inférieur via le **sélecteur de rôle** dans la topbar. C'est utile pour tester l'expérience d'un contrôleur. Cliquer sur « (mon rôle) » à côté du sélecteur pour revenir à son rôle réel.

---

## 5. Module Anomalies — En cours

### 5.1 Vue générale

Cette page affiche **uniquement les anomalies non résolues** (statuts Ouvert et En cours). Les anomalies déjà closes ou résolues sont consultables sur la page « Anomalies résolues ».

### 5.2 Consulter la liste

Le tableau présente les colonnes :

- **Numéro** : `T-{ID}` (identifiant SharePoint)
- **Statut** : Ouvert / En cours (pastille colorée)
- **Date de déclaration**
- **Criticité** : Faible / Moyenne / Haute / Critique
- **Délai** : pastille visuelle basée sur la date d'affectation + délai accordé

Colonnes **dépliables** (bouton +/-) : Déclarant, Auteur, Personne affectée, Cause, Classification, Agence, Réseau, Montant, Date régularisation.

### 5.3 Filtrer

La barre de filtres propose :

- **Recherche libre** : n°, titre, description, classification…
- **Date du / Date au** : période de déclaration
- **Agence / Réseau / Classification / Criticité / Domaine d'activité**
- **Statut** : Ouvert / En cours
- **Agent (auteur) / Personne affectée / Déclarant** : recherche partielle sur nom ou email

Cliquer sur **Rechercher** pour appliquer, **Réinitialiser** pour tout vider.

### 5.4 Créer une anomalie

1. Cliquer sur **+ Nouvelle anomalie** en haut à droite
2. Le formulaire s'organise en 5 sections :
   - **Identité** : le déclarant est pré-rempli avec votre nom (non modifiable). Renseigner l'auteur (personne à l'origine de l'anomalie).
   - **Caractérisation** : classification, criticité, domaine d'activité, type de risque
   - **Localisation** : agence, réseau
   - **Impact** : description, cause, montant estimé
   - **Pièce jointe** : optionnelle mais fortement recommandée (justificatif)
3. Cliquer sur **Enregistrer**

Après création, un **ticket de suivi** est ouvert automatiquement avec le statut « Ouvert ».

### 5.5 Affecter une anomalie (managers uniquement)

1. Cliquer sur le bouton **Affecter** d'une ligne
2. **Étape 1** : rechercher une personne dans Office 365 et la sélectionner
3. **Étape 2** : définir un **délai** (par défaut 3 jours) et saisir un **commentaire** d'affectation
4. Valider

La personne affectée reçoit automatiquement une **notification Microsoft Teams** (voir §13.1) avec le résumé de l'anomalie et le délai.

### 5.6 Suivre le ticket

Le bouton **Ticket** ouvre une vue synthétique avec les infos clés, le commentaire d'affectation, et les actions disponibles :

- **Détail** : voir tous les champs et modifier si autorisé
- **Changer le statut** : basculer Ouvert ↔ En cours
- **Clore la résolution** : formulaire complet de clôture (voir §5.7)

### 5.7 Clôturer une anomalie

Réservé aux managers et au contrôleur affecté. Cliquer sur **Clore la résolution** :

1. Choisir le **statut final** : Resolu ou Clos
2. Renseigner les **dates** : régularisation, clôture
3. **Auteur de la clôture** : modifiable si nécessaire
4. **Description structurée** :
   - Cause immédiate
   - Cause racine
   - Actions menées
   - Observations
5. **Mode de traitement** (optionnel) : Relance, Avertissement, Blâme, Suspension, Formation, etc.
6. **Pièce jointe** : rapport de résolution recommandé

Cliquer sur **Valider** pour enregistrer. L'anomalie disparaît de la liste « En cours » et bascule dans « Anomalies résolues ».

### 5.8 Modifier une anomalie

Bouton **Détail** puis **Modifier** dans la modale. Modifications possibles selon le rôle :

- Managers : tous les champs métier sauf les personnes (auteur, déclarant, affecté)
- Contrôleur déclarant ou affecté : idem

Pour changer l'auteur ou l'affecté, utiliser les flux dédiés (Affectation, formulaire de clôture).

---

## 6. Module Anomalies — Résolues

### 6.1 Vue générale

Cette page affiche **uniquement les anomalies clôturées** (statuts Resolu ou Clos), sous forme de **cartes** (bulletins). Chaque carte présente les informations essentielles avec un lien vers la fiche complète.

### 6.2 Consulter un bulletin

Cliquer sur une carte ouvre la **fiche récapitulative complète** en modale, avec les sections :

1. **En-tête** : numéro, titre, statut, criticité, délai
2. **Identification** : déclarant, auteur, personne affectée, agence, réseau, domaine
3. **Dates clés** : déclaration, ouverture ticket, régularisation, clôture, délai
4. **Caractérisation** : classification, criticité, type de risque, montant, mode de traitement, statut
5. **Description et causes** : description consolidée, cause immédiate, cause racine, observations
6. **Cycle de vie** : timeline chronologique des étapes clés (voir §6.5)
7. **Actions à mener** : liste des actions consignées (cochées si Clos)
8. **Pièces jointes** : liens vers les fichiers avec icônes par type

### 6.3 Filtrer

Les filtres sont **identiques** à ceux de la page « Anomalies en cours » : recherche libre, dates, agence, réseau, classification, criticité, domaine, statut, agent, personne affectée, déclarant. Cela permet de trouver rapidement un bulletin historique.

### 6.4 Modifier un bulletin (Directeur uniquement)

Le **Directeur** peut rectifier une anomalie déjà close, par exemple pour corriger une erreur de saisie ou compléter un champ. Sur la fiche récapitulative :

1. Cliquer sur **Modifier** (bouton visible uniquement pour le Directeur)
2. Une modale d'édition s'ouvre avec les champs autorisés :
   - Statut (Resolu ↔ Clos)
   - Date de régularisation, date de clôture
   - Montant
   - Mode de traitement
   - Description consolidée, actions menées
3. Modifier puis cliquer sur **Enregistrer**

Les champs identifiants (déclarant, agence, criticité, classification) ne sont pas éditables — ils sont considérés comme figés métier.

### 6.5 Comprendre le cycle de vie

La timeline affiche les étapes suivantes dans l'ordre :

1. **Déclaration** (date de création système dans SharePoint)
2. **Survenance** (date métier saisie par l'agent, où l'anomalie s'est réellement produite)
3. **Ouverture du ticket** (si différente)
4. **Régularisation**
5. **Clôture**
6. Notes libres du journal (si présentes)

### 6.6 Imprimer un bulletin

Le bouton **Imprimer** en haut de la fiche déclenche l'impression navigateur. La sidebar et la topbar sont automatiquement masquées grâce à une feuille de style dédiée. Choisir **Enregistrer en PDF** dans la boîte d'impression pour obtenir un fichier PDF.

---

## 7. Module Rapport quotidien

Ce module regroupe **3 vues** selon le rôle et le workflow :

- **Saisie du rapport** : contrôleurs (renseigner l'activité du jour)
- **Mes rapports** : contrôleurs (historique de leurs propres soumissions)
- **Validation reportings** : managers (valider ou refuser les rapports)

### 7.1 Saisie du rapport (contrôleur)

Chaque contrôleur doit soumettre **un rapport par jour ouvré** décrivant son activité.

1. Cliquer sur **Saisie du rapport** dans le menu
2. La date est pré-remplie au jour courant (modifiable pour rattraper un rapport oublié)
3. Renseigner les **lignes d'activité** : type d'activité, description, heure début/fin
4. **Statut de la journée** : Calme / Normal / Chargé / Très chargé
5. **Anomalies détectées** : compteur si des anomalies ont été identifiées ce jour-là
6. **Observations globales** : commentaire libre
7. **Pièces jointes** : optionnel (rapport complémentaire)
8. Cliquer sur **Soumettre**

Le rapport passe en statut « Soumis » et est visible par les managers dans « Validation reportings ».

**Brouillon** : les données sont automatiquement sauvegardées en local si vous quittez sans soumettre. Vous les retrouverez à votre prochaine ouverture.

### 7.2 Mes rapports (contrôleur)

Vue historique de tous vos rapports soumis :

- **Filtres** : date, statut
- **Statut** de chaque rapport :
  - **Soumis** : en attente de validation manager
  - **Valider** : approuvé
  - **Refuser** : rejeté (motif visible)
- **Total heures** cumulé sur la période

En cas de refus, cliquer sur le rapport pour voir le **motif de rejet** et pouvoir **re-soumettre** une version corrigée.

### 7.3 Validation reportings (managers)

Vue consolidée de tous les rapports soumis par les contrôleurs.

**Filtres** :
- Contrôleur (recherche partielle sur nom ou email)
- Période (Date du / Date au)
- Case **En attente uniquement** pour ne voir que les rapports non traités

**KPI en haut** :
- Total reportings, En attente, Validés, Refusés, Total heures

**Section Taux de soumission des rapports quotidiens** :
- Repliée par défaut, cliquer sur le header pour l'ouvrir
- Le **taux global** est visible même repliée (pastille noire à côté du titre)
- Une fois dépliée : bandeau taux global + tableau par contrôleur
- Cliquer sur une ligne contrôleur pour voir les **jours manquants** (chips rouges datées) sur la période

Le calcul du taux :
```
taux = (jours ouvrés couverts par un rapport non refusé)
       / (jours ouvrés dans la période)
```
Période par défaut : 30 derniers jours. Les weekends sont exclus du dénominateur. Les jours fériés ne sont pas exclus.

**Groupes de rapports** : chaque « carte » représente un couple (contrôleur, date). Cliquer dessus pour voir le détail et prendre la décision.

**Valider** :
1. Cliquer sur le bouton **Valider**
2. Note optionnelle
3. Le rapport passe en statut « Valider »

**Refuser** :
1. Cliquer sur **Invalider**
2. **Motif obligatoire** (textarea)
3. Le rapport passe en statut « Refuser ». Le contrôleur peut alors le corriger et le re-soumettre.

---

## 8. Module Reporting par Agent

Vue agrégée des anomalies **regroupées par personne**.

### 8.1 Modes de regroupement

Sélecteur en haut :

- **Par personne affectée** (défaut) : mesure la **charge** des contrôleurs — qui gère quoi
- **Par auteur** : mesure la **production** de signalements — qui a déclaré quoi

### 8.2 Vue principale

- **Cartes KPI** : nombre d'agents, total anomalies, montant total
- **Filtres** : dates, agence, réseau, classification, criticité
- **Tableau** : une ligne par agent avec compteurs par statut (Ouvert / En cours / Resolu / Clos), montant total, domaines d'activité (chips bleues), types de traitement (chips ambrées)

Cliquer sur un agent pour ouvrir sa **vue détail**.

### 8.3 Vue détail agent

- En-tête : nom + email + bouton **Imprimer le bulletin**
- Cartes KPI spécifiques à l'agent
- Filtres propres à la vue détail
- Tableau des anomalies de l'agent (paginé)
- Bouton **Retour à la liste**

### 8.4 Bulletin imprimable — Fiche de notation

Le bouton **Imprimer** ouvre la fiche de notation formatée pour impression papier. Selon le mode de regroupement, le titre s'adapte :

- Mode « affecté » : **Fiche de notation du contrôleur**
- Mode « auteur » : **Fiche de notation de l'auteur**

La fiche comporte 3 sections avec pour chacune un tableau Critères / Prévu / Réalisé / Écart et un score % :

1. **Surveillance des anomalies** — colonnes **Détectée** / **Corrigée** / Écart (par classification)
2. **Surveillance des plans de contrôle** — colonnes Prévu / Réalisé / Écart (1 ligne par contrôle assigné)
3. **Surveillance des plans d'action correctifs** — colonnes Prévu / Réalisé / Écart (1 ligne par PAC assigné)

Un **score global** en bas résume la performance de l'agent sur la période.

---

## 9. Module Plan de Contrôle

Gestion des **contrôles récurrents** planifiés (mensuels, trimestriels, semestriels, annuels).

### 9.1 Vue générale

- **KPI** : Total, À planifier, Planifiés, En cours, Réalisés, **Taux d'évolution global**
- **Filtres** : recherche libre, catégorie, fréquence, responsable, statut, année
- **Tableau** paginé avec barre de progression par contrôle

### 9.2 Créer un contrôle (managers uniquement)

1. Cliquer sur **+ Nouveau contrôle**
2. Renseigner : libellé, catégorie, nature d'activité, fréquence, année, responsable, objectif, objectif chiffré
3. Ajouter des pièces jointes (procédures, référentiels)
4. Enregistrer

Le contrôle passe en statut « À planifier ».

### 9.3 Affecter un responsable

Cliquer sur le bouton **Affectation** d'une ligne, sélectionner une personne O365. Le nouveau responsable reçoit une **notification Teams**.

### 9.4 Évaluer un contrôle

Cliquer sur le bouton **Évaluer** :

1. Date de l'évaluation
2. Constats
3. Actions correctives
4. Pièces jointes justificatives
5. Statut résultant : Réalisé / En cours / etc.

Chaque évaluation est **historisée**. Le compteur d'évaluations effectuées alimente le taux d'évolution global.

### 9.5 Taux d'évolution global

Le taux affiché en KPI est le ratio agrégé :

```
taux = (somme des évaluations effectuées) / (somme des évaluations attendues) × 100
```

Le nombre d'évaluations attendues dépend de la fréquence :
- Mensuelle : 12 par an
- Trimestrielle : 4
- Semestrielle : 2
- Annuelle : 1

---

## 10. Module Plan d'Action Correctif (PAC)

Gestion des **actions correctives** décidées suite à une anomalie ou un contrôle.

### 10.1 Vue générale

- **KPI** : Total PAC, En cours, Exécutées, Non Exécutées, **Taux d'évolution global**
- **Filtres** : recherche libre, statut, année, responsable, source
- **Tableau** paginé

### 10.2 Créer un PAC (managers uniquement)

1. Cliquer sur **+ Nouveau PAC**
2. Renseigner :
   - **Source PAC** : anomalie, contrôle, audit, autre
   - **Intitulé** et **description du problème**
   - **Cause immédiate** et **cause racine**
   - **Actions correctives** à mener
   - **Directions concernées** (multiselect)
   - **Échéance**, **année**, **KPI**
   - **Responsable de mise en œuvre**
3. Ajouter les pièces jointes
4. Enregistrer

Le responsable reçoit une notification Teams.

### 10.3 Évaluer un PAC

Cliquer sur **Évaluer** :

1. Date d'évaluation
2. Constats sur la mise en œuvre
3. Statut : Exécutée / Non Exécutée / En cours
4. Observations
5. Pièces jointes (preuves)

### 10.4 Taux d'évolution global

Le taux affiché est le ratio :

```
taux = (nb de PAC Exécutées) / (nb total de PAC visibles) × 100
```

Cohérent avec la fiche de notation des contrôleurs (chaque PAC = 1 tâche à exécuter).

---

## 11. Dashboards Power BI

Section **Dashboards** dans le menu latéral. Contient plusieurs tableaux de bord Power BI embarqués :

- **Numérisation des journées** : suivi de la numérisation quotidienne
- **Créance Hors Bilan** : indicateurs de créance
- **Apurement Comex** : suivi des dossiers Comex
- **Clôture des journées** (sous-groupe)

Les dashboards sont hébergés sur Power BI Service et s'affichent en **iframe** dans l'application. Les filtres et interactions sont ceux de Power BI (drill-down, exports, etc.).

Si un dashboard ne charge pas, vérifier :
- Votre accès Power BI (contacter la DSI)
- La connexion internet
- L'authentification Office 365 (déconnecter/reconnecter si expirée)

---

## 12. Configuration (managers uniquement)

Section **Configuration** dans le menu latéral.

### 12.1 Directions

Gère le **référentiel des directions** de la banque, utilisé notamment dans le champ « Directions concernées » d'un PAC.

- Ajouter une direction : bouton **+ Nouvelle**
- Modifier / Supprimer via les boutons dans le tableau

### 12.2 Gestion des utilisateurs

Gère la **liste des utilisateurs autorisés** à accéder à l'application, avec leur rôle.

**KPI** : Total, Directeurs, Chefs de département, Contrôleurs

**Ajouter un utilisateur** :
1. Bouton **+ Nouveau**
2. Nom, Email, Fonction (Directeur / Chef_Departement / Controleur)
3. Enregistrer

L'utilisateur doit exister dans Office 365 pour se connecter. Ajouter ici définit son **rôle applicatif**.

**Modifier / Supprimer** : boutons dans le tableau. **Attention** : supprimer un utilisateur lui coupe l'accès à l'application immédiatement.

---

## 13. Fonctionnalités transverses

### 13.1 Notifications Microsoft Teams

Certaines actions déclenchent une **notification Teams automatique** au destinataire :

- Affectation d'une anomalie
- Création d'une anomalie avec responsable pré-affecté
- Affectation d'un plan de contrôle
- Création d'un plan de contrôle avec responsable
- Affectation d'un PAC
- Création d'un PAC avec responsable

Le destinataire reçoit un message dans son chat Teams avec un résumé et éventuellement un lien vers l'application. La notification est en mode **fire-and-forget** : si l'envoi Teams échoue (Teams indisponible, workflow off), l'action métier (l'affectation) est quand même enregistrée en SharePoint.

### 13.2 Export Excel et PDF

Presque toutes les pages à tableau proposent deux boutons **Excel** et **PDF** dans l'en-tête :

- **Excel** : télécharge un fichier **CSV** (compatible Excel, séparateur `;`, encodage UTF-8 avec BOM). Excel l'ouvre nativement avec les accents corrects.
- **PDF** : ouvre une nouvelle fenêtre avec un tableau prêt à imprimer, puis déclenche l'impression. Choisir **Enregistrer au format PDF** dans la boîte d'impression.

**Important** :
- L'export inclut **toutes les données filtrées**, pas seulement la page courante
- Si aucun résultat, un message d'alerte vous prévient
- Pour le PDF, autoriser les **fenêtres popup** si votre navigateur les bloque

Disponible sur : Anomalies en cours, Anomalies résolues, Plan de Contrôle, PAC, Validation reportings, Mes rapports, Reporting par Agent.

### 13.3 Impression et enregistrement en PDF

Certaines pages ont un bouton **Imprimer** dédié qui produit un rendu papier optimisé (Bulletin d'anomalie, Fiche de notation). Choisir **Enregistrer au format PDF** dans la boîte d'impression du navigateur pour obtenir un PDF propre.

### 13.4 Filtres et pattern « Rechercher »

La plupart des filtres suivent le pattern **saisie / applied** :

- Vous tapez dans les champs — rien ne se passe encore
- Vous cliquez sur **Rechercher** — les filtres sont **snapshot** et appliqués
- Le bouton **Réinitialiser** vide tout et refait la recherche

Ce pattern évite les recharges intempestives à chaque frappe.

### 13.5 Pagination

- Taille par défaut : **50 lignes par page**
- Sélecteur en bas de tableau pour changer (25, 50, 100, 200)
- Navigation : première, précédente, page courante, suivante, dernière

Les KPI et exports portent sur **toutes les données filtrées**, pas seulement la page.

### 13.6 Pièces jointes

Plusieurs types de fichiers acceptés :
- **PDF**, **Word**, **Excel**, **PowerPoint**
- **Images** : JPG, PNG, GIF
- **Autres** : ZIP, TXT

Une icône visuelle par type est affichée à côté du nom du fichier. Cliquer sur le nom pour l'ouvrir dans un nouvel onglet.

**En cas d'échec d'upload** : un message d'erreur s'affiche. Le fichier n'est pas enregistré, mais l'anomalie/PAC/contrôle est créé quand même (à ce jour). Pour ajouter la pièce jointe après coup, ouvrir l'élément et refaire l'upload.

---

## 14. FAQ et dépannage

### 14.1 Je ne vois pas certains onglets dans le menu

Certains modules sont **réservés aux managers** (Validation reportings, Configuration). Si vous êtes contrôleur, c'est normal. Contactez votre manager si vous pensez qu'il y a une erreur de rôle.

### 14.2 La liste des anomalies est vide alors qu'il devrait y en avoir

Vérifier les filtres appliqués. Cliquer sur **Réinitialiser** pour tout vider.

Si le problème persiste, vérifier votre connexion internet et cliquer sur **Actualiser** si le bouton est présent.

### 14.3 Erreur 400 sur une page

Une erreur 400 signifie que SharePoint a rejeté la requête (souvent un filtre invalide). Actions :
1. **Réinitialiser** les filtres
2. Recharger la page (F5)
3. Si l'erreur persiste, contacter la DSI en précisant la page et l'action effectuée

### 14.4 Le rapport quotidien ne veut pas s'enregistrer

Vérifier que **tous les champs obligatoires** sont remplis (lignes d'activité au minimum). Si le problème persiste, essayer de rafraîchir la page — le brouillon sera conservé.

### 14.5 Je n'ai pas reçu la notification Teams

Causes possibles :
- Teams n'est pas installé ou pas connecté sur votre poste
- Votre email dans la fiche utilisateur ne correspond pas à votre compte Teams
- Le workflow Power Automate est indisponible temporairement

La notification est un **plus** — l'affectation elle-même est bien enregistrée dans l'application. Vérifiez directement la liste des anomalies qui vous sont affectées.

### 14.6 Comment corriger un rapport refusé ?

1. Aller dans **Mes rapports**
2. Ouvrir le rapport refusé (statut « Refuser »)
3. Lire le **motif de rejet**
4. Cliquer sur **Modifier** pour éditer les champs
5. **Re-soumettre**

Le rapport repasse en statut « Soumis ». Le manager le retrouve dans sa file de validation.

### 14.7 Le taux de soumission ne correspond pas à ce que j'attends

Vérifier :
- La **période** sélectionnée (par défaut 30 derniers jours si aucune date saisie)
- Les rapports refusés ne comptent **pas** dans le numérateur
- Les weekends sont exclus, mais **pas les jours fériés** (limite connue)
- Un contrôleur nouvellement arrivé aura un taux mécaniquement bas s'il n'a pas encore rattrapé toute la période

### 14.8 L'application semble lente

- Nombre d'items élevé (> 5000 par liste SharePoint) : les filtres serveur doivent être bien indexés
- Connexion internet dégradée
- Trop d'onglets Power BI ouverts en parallèle

Contactez la DSI si le problème est récurrent.

### 14.9 Où signaler un bug ou une amélioration ?

- Contacter le manager de la DCPO
- Ou envoyer un email à la DSI avec :
  - Une **capture d'écran** de l'écran concerné
  - Le **rôle** avec lequel vous êtes connecté
  - L'**action effectuée** juste avant le problème
  - L'**heure approximative** du problème

---

## Contacts

- **DCPO** — pour les questions métier et fonctionnelles
- **DSI Afriland First Bank** — pour les problèmes techniques, accès et authentification
- **Administrateur applicatif** — pour l'ajout d'utilisateurs et la configuration des rôles

---

*Ce guide est un document vivant. Il sera mis à jour au fil des évolutions de l'application. Toute suggestion d'amélioration est bienvenue.*
