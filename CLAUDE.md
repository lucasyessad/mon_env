# CLAUDE.md

Ce fichier guide Claude Code (claude.ai/code) lorsqu'il travaille sur ce dépôt.

## Vue d'ensemble

**Suivi d'Exploitation Multi-Squads** : désigne automatiquement chaque semaine, pour
chaque squad, **une personne par rôle** (1 rôle = solo, 2 = binôme, 3 = trio…), puis
envoie les mails d'annonce (vendredi) et de rappel (lundi).

Conçu pour **Windows** (PowerShell 5.1 + Excel), exécuté via le **Task Scheduler** sur un
partage réseau. **Pas de build ni de tests automatisés** : la validation se fait avec
`-Action Test` et des exécutions manuelles `Annonce`/`Rappel`.

Tout le code vit dans `suiviexploit/`. La référence fonctionnelle est
`suiviexploit/README.md` — **à consulter avant toute modification de comportement.**

## Structure

```
suiviexploit/
├── SuiviExploitation.ps1            tout le code (script unique)
├── config-suivi.json                configuration globale (partagée avec le moteur mail)
├── SendMailNotificationHTML.ps1     moteur d'envoi HTML (dépendance requise)
├── template-notification.html       gabarit HTML des mails
├── suivi_exploitation_template.xlsx gabarit d'un classeur de squad
├── suivi_exploitation_exemple.xlsx  exemple rempli
└── README.md                        guide d'installation/exploitation
```

Déploiement : ces fichiers à la racine du partage ; un dossier par squad contenant
`suivi_exploitation.xlsx` + `historique\` ; `_logs\` généré pour l'orchestrateur **et** le
journal daté de chaque squad (`_logs\<squad>_<date>.journal.log`).

## Organisation du script

`SuiviExploitation.ps1 -Action <Annonce|Rappel|Test|NouvelleSquad>` plus `-DossierRacine`,
`-Config`, `-Squad`.

- Outils communs : `Get-Opt`, `ConvertTo-Ht` (JSON → hashtable), `Resolve-Chemin`,
  `Resolve-MoteurMail`, `Get-Configuration`, `Send-Notification`.
- `Invoke-SquadJob` : traite UNE squad (verrou, lecture, désignation, mail, écriture
  transactionnelle, rétention) et **retourne un objet résultat** (`Squad`/`Statut`/`Detail`).
- `Invoke-Orchestration` : boucle sur les squads (try/catch par squad), journal dans
  `_logs\`, mail récapitulatif **uniquement** en cas d'anomalie.
- `Invoke-Test` : validation. `Invoke-NouvelleSquad` : onboarding.

## Concepts clés

### Configuration : un seul JSON, partagé avec le moteur mail
- `config-suivi.json` est lu par le suivi (`ConvertFrom-Json` → `ConvertTo-Ht`) **et**
  passé tel quel au moteur (`-ConfigFile`). Source unique.
- Clés **moteur** : `SmtpServer`, `Port`, `From`, `To`, `TemplatePath` (= `${TEMPLATE_PATH}`,
  fourni par le suivi via une variable d'environnement), `Subject`, `Statuses`
  (DESIGNATION/RAPPEL/ALERTE), `Environnement`, `EquipeNom`.
- Clés **suivi** : `EmailsAdmin`, `MoteurMail`, `Template` (HTML), `TemplateClasseur`
  (xlsx pour NouvelleSquad), `DossierSquads`, `DossierLogs`, `Squads`
  (par squad : `Nom`, `Classeur` — c'est tout).
- Obligatoires validées par `$script:ClesGlobales` (`SmtpServer`, `From`, `To`,
  `TemplatePath`, `EmailsAdmin`) et `$script:ClesSquad` (`Nom`, `Classeur`).
- `Classeur` absolu/UNC pris tel quel ; relatif résolu via `DossierSquads` puis
  `DossierRacine` (`Resolve-Chemin`).
- `-Action NouvelleSquad` **met à jour le JSON** (ajout au tableau `Squads`, sauvegarde
  `.bak`, vérification du rechargement).

### Configuration par squad : feuille `Parametres` du classeur
- **Colonne A** : rôles à désigner, ordonnés (1=solo, 2=binôme, 3=trio…). Alimente aussi
  la liste déroulante `Rôle` de `Membres`. Lue dans `$rolesADesigner`.
- **Colonnes C/D** : paires clé/valeur self-service de la squad, lues dans `$paramKV`
  (ex. `EmailsCopieSquad`). Zone extensible.

### Excel = données métier + config squad
Feuilles `Parametres`, `Membres`, `Congés`, `Historique`, `Log`. **Pas de VBA.**
`Historique` au **format long** : 1 ligne par rôle désigné (`SemaineLundi, Role, Nom,
Email, DateAnnonce, DateRappel, Statut, Note`). Listes déroulantes : `Rôle` =
`Parametres!$A$2:$A$50`, `Actif` = Oui/Non, `Congés!Membre` = noms de `Membres`. Lecture
tolérante : `ConvertTo-DateOuNull` / `ConvertTo-IntOuZero` ; lectures encadrées par `@(...)`.

### Désignation (par rôle)
Pour chaque rôle de `Parametres`, on désigne 1 personne : candidats actifs du rôle →
exclusion congés (semaine cible ou vendredi précédent) → exclusion du désigné précédent
de ce rôle (sauf seul candidat) → **tri** par ordre de priorité : 1) **dernière
désignation la plus ancienne** (ceux qui n'ont **jamais** fait d'abord), 2) **min
`NB_FOIS`** (le moins souvent désigné), 3) `Nom`. La « dernière désignation » est la
semaine la plus récente trouvée dans l'`Historique` pour la personne (jamais = priorité
absolue, modélisé par `[DateTime]::MinValue`). `NB_FOIS` = `Compteur` (graine de la
feuille `Membres`) + nb de désignations dans l'`Historique`. Les deux critères sont
**dérivés de l'historique** ; le script n'écrit jamais dans `Membres`.
**Recalcul idempotent** (`Invoke-DesignationSemaine`) partagé par Annonce et Rappel : on
supprime les lignes de la semaine cible puis on redésigne depuis les données à jour.
Donc `Rappel` (lundi, semaine en cours) reprend les congés ajoutés le week-end. Si un rôle
n'a aucun candidat → **ALERTE atomique** (aucune désignation commitée), mail aux `EmailsAdmin`.

### Envoi des mails : délégation au moteur CL
`Send-Notification` est le point d'envoi unique. Il **délègue toujours** à
`SendMailNotificationHTML.ps1` (lancé en **process séparé** : `pwsh`/`powershell.exe`
selon l'édition) avec `-ConfigFile <le même JSON>`, `-Status` (DESIGNATION/RAPPEL/ALERTE),
`-NomJob` (squad), `-OverrideTo`/`-OverrideCc`, `-KeyValues`, et un `-SectionFile`. SMTP,
expéditeur, sujet et template viennent du JSON. **Pas de repli** : si le moteur est
introuvable, l'envoi échoue explicitement.

Les sections (tableaux, notes) sont passées à `Send-Notification` en **objets** puis
sérialisées **une seule fois** via `ConvertTo-JsonSafe` (émetteur JSON maison). C'est
**volontaire** : `ConvertTo-Json` de PS 5.1 « déballe » les tableaux à un seul élément et
n'échappe pas comme PS 7, ce qui produisait un JSON invalide → sections silencieusement
perdues (mail réduit à la seule ligne « Désignation de la semaine »). Toute construction de
lignes de tableau utilise `… += , @(...)` puis `return , $rows` pour ne pas déballer une
ligne unique. **Ne pas réintroduire d'aller-retour `ConvertTo-Json`/`ConvertFrom-Json`.**

Le mail d'**annonce** inclut (via `-SectionsInline`) un tableau de **compteurs par rôle**
(`Get-CompteursRows <Role>`) et la **prévision des 3 prochaines semaines**
(`Get-PrevisionRows`, 1 colonne par rôle, simulation de l'algo sur copies en mémoire).

### Accès concurrent au classeur
Verrou applicatif `<squad>\suivi_exploitation.lock` **honoré** : si un verrou est présent et
récent (< `$script:VerrouStaleMinutes`, 120 min) le traitement de la squad est **abandonné**
(ERREUR) ; au-delà il est considéré orphelin (process tué) et ignoré avec WARN. Le `.lock`
n'est supprimé que par le process qui l'a créé (drapeau `$jaiLeVerrou`). On attend aussi le
verrou Excel `~$*`. Écriture **transactionnelle** : le résultat est calculé dans un `.tmp`,
on **copie** l'état courant dans `historique\` puis on **promeut** le `.tmp` (écrasement) —
pas de fenêtre où le classeur a disparu si la promotion échoue.

## Conventions de modification

- **Langue** : tout (code, messages, mails, doc) en **français**.
- **PowerShell 5.1** : `Set-StrictMode -Version Latest` + `$ErrorActionPreference='Stop'`.
  Pas de syntaxe PS7-only.
- **Style** : commentaires **limités aux parties complexes** (dates, désignation,
  délégation moteur, écriture transactionnelle…). Pas de commentaire qui paraphrase le
  code ; le détail fonctionnel va dans le README.
- **Encodage** : `SuiviExploitation.ps1` et `config-suivi.json` en **UTF-8 avec BOM**
  (accents sous PS 5.1). **Conserver le BOM** lors des éditions (les outils peuvent le
  retirer — vérifier après coup). Tout le pipeline est UTF-8 (accents validés de bout en bout).
- **Dépendances** : module `ImportExcel` (installé auto en `-Scope CurrentUser`) ; moteur
  `SendMailNotificationHTML.ps1`.
- **Rétention** : `$script:RetentionJours` (28) en haut du script — un seul endroit.
- **SMTP** : anonyme (le moteur utilise `Send-MailMessage`).
- Ne pas committer d'artefacts d'exécution (`_logs/`, `historique/`, `*.lock`, `~$*`,
  `*.journal.log`, `*.bak`) — couverts par `.gitignore`.

## Validation avant livraison

Pas de suite de tests. Pour valider un changement :
1. Lint : `pwsh -NoProfile -Command "Invoke-ScriptAnalyzer -Path ./suiviexploit -Recurse"`
   (le hook `.claude/hooks/session-start.sh` installe pwsh + PSScriptAnalyzer en session web).
   Référence : 0 erreur (warnings résiduels = `Write-Host` + quelques noms de fonctions au pluriel).
2. `-Action Test -DossierRacine <...>` doit être vert (hors connectivité SMTP).
3. Exécution manuelle `-Action Annonce` puis vérifier `_logs\`, l'`Historique`,
   `<squad>\historique\` et les mails.

> Les validations 2-3 nécessitent **Windows + Excel** ; non exécutables en conteneur Linux
> (mais l'enchaînement désignation → moteur → mail est testable sous pwsh + ImportExcel).
