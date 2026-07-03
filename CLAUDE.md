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
├── SuiviExploitation.ps1            tout le code (script unique, global)
├── SendMailNotificationHTML.ps1     moteur d'envoi HTML (dépendance partagée, NON modifiée)
├── template-notification.html       gabarit HTML des mails (partagé)
├── suivi_exploitation_template.xlsx gabarit d'un classeur de squad
├── suivi_exploitation_exemple.xlsx  exemple rempli
├── conf-suivi-squad.exemple.json    exemple commenté de config COMPLÈTE d'une squad
└── README.md                        guide d'installation/exploitation
```

**Pas de config globale.** Déploiement : ces fichiers partagés à la racine ; puis **un dossier
par squad** contenant `conf-suivi-squad.json` (config complète), `suivi_exploitation.xlsx` et
`historique\`. Les squads sont **découvertes en scannant les sous-dossiers** de `-DossierRacine`
(un dossier avec `conf-suivi-squad.json` = une squad ; dossiers `_…` ignorés). `_logs\` reçoit
le log d'orchestration **et** le journal daté de chaque squad (`_logs\<squad>_<date>.journal.log`).

## Organisation du script

`SuiviExploitation.ps1 -Action <Annonce|Rappel|Apercu|Rapport|Test|NouvelleSquad>` plus
`-DossierRacine`, `-Squad` (filtre par nom de dossier / `Nom`, correspondance **exacte** —
`Select-SquadsParFiltre`). **Plus de `-Config`** : chaque squad porte sa config.

- Outils communs : `Get-Opt`, `ConvertTo-Ht` (JSON → hashtable), `ConvertTo-Liste`,
  `Resolve-Chemin`, `Resolve-MoteurMail`, `ConvertTo-JsonSafe`, `Send-Notification`,
  `Get-JoursFeriesFR` (fériés France, Meeus/Butcher), `Test-ConfSquad` (validation
  présence + non-vacuité des clés).
- `Get-DossiersSquads` (scan des sous-dossiers) + `Get-ConfSquad` (charge le `conf-suivi-squad.json`
  d'une squad → hashtable). Remplacent l'ancien `Get-Configuration` global.
- `Invoke-SquadJob` : traite UNE squad (reçoit son dossier + sa config complète ; verrou, lecture,
  désignation, mail, écriture transactionnelle, rétention) et **retourne** `Squad`/`Statut`/`Detail`.
  Modes `Apercu`/`Rapport` = **lecture seule** (copie de travail en %TEMP%, pas de verrou,
  rien d'écrit) : `Apercu` affiche en console la désignation prévue/prévision/compteurs,
  `Rapport` envoie ces mêmes éléments aux `EmailsAdmin` (statut `RAPPORT`, injecté par
  défaut dans la config moteur s'il manque).
- `Invoke-Orchestration` : découvre les squads, boucle (try/catch par squad), journal dans
  `_logs\`. **Pas de mail récapitulatif global** (chaque squad alerte ses propres `EmailsAdmin`).
- `Invoke-Test` : validation par squad. `Get-ConfSquadModele` + `Invoke-NouvelleSquad` : onboarding.

## Concepts clés

### Configuration : une config COMPLÈTE par squad (pas de global)
- Chaque squad = un sous-dossier de `-DossierRacine` contenant `conf-suivi-squad.json`.
  `Get-DossiersSquads` scanne ces dossiers ; `Get-ConfSquad` lit le fichier
  (`ConvertFrom-Json` → `ConvertTo-Ht`). **Il n'y a pas de `config-suivi.json` global.**
- Ce fichier est **autonome et complet** : `SmtpServer`, `Port`, `From`, `To`, `Cc`,
  `EmailsAdmin`, `TemplatePath`, `MoteurMail`, `Classeur`, `Subject`, `Environnement`,
  `EquipeNom`, `Statuses` (DESIGNATION/RAPPEL/ALERTE/RAPPORT), `ToleranceCongesJours`
  (défaut 0), `JoursFeries` (`'FR'` défaut / `'Aucun'`), `SemainesPrevision` (défaut 3,
  borné 1..12), `AbsencesCsv` (import optionnel d'un CSV d'absences : `Chemin` +
  `Separateur`/`Encodage`/`ColonneNom`/`ColonneDebut`/`ColonneFin` ; fusion EN MÉMOIRE,
  noms inconnus filtrés, échec = WARN non bloquant), et `Nom` (affiché, défaut = dossier).
- Obligatoires validées par `Test-ConfSquad` (`$script:ClesConfSquad` : `SmtpServer`, `From`,
  `To`, `TemplatePath`, `EmailsAdmin`) — **présence ET non-vacuité** (`"To": []` refusé).
  Constantes : `$script:NomConfSquad` (`conf-suivi-squad.json`), `$script:NomClasseur`,
  `$script:FeuillesClasseur` (5 feuilles, exigées à l'exécution ET par `-Action Test` —
  qui contrôle aussi la **cohérence des congés** : noms connus, dates valides/ordonnées).
- Le **To** des mails de désignation reste **les désignés** ; `To` n'est qu'un **repli** si aucun
  désigné n'a d'adresse. `Cc` vide ⇒ tous les membres actifs. `EmailsAdmin` = alertes de la squad.
- Clés de **documentation** `_…` (ex. `_commentaire`) tolérées : ignorées par le code et **non
  transmises** au moteur. `TemplatePath`/`MoteurMail`/`Classeur` : relatifs (résolus via la racine,
  ou le dossier de squad pour `Classeur`) ou absolus/UNC (`Resolve-Chemin`).
- `-Action NouvelleSquad` crée le dossier, copie le classeur et **écrit un `conf-suivi-squad.json`
  complet** (`New-ConfSquadComplete` + saisies) ; il ne touche aucun fichier global.

### Données métier dans le classeur (feuille `Parametres`)
- **Colonne A** : rôles à désigner, ordonnés (1=solo, 2=binôme, 3=trio…). Alimente aussi la liste
  déroulante `Rôle` de `Membres`. Lue dans `$rolesADesigner`.
- **Colonnes C/D** : paires clé/valeur self-service, lues dans `$paramKV` : `Note <Rôle>` (texte
  affiché dans le mail), `EmailsCopieSquad` (*legacy* Cc, repli seulement) et `ToleranceCongesJours`
  (règle de désignation, **prioritaire** ici sur le JSON ; cf. exclusion congés). Zone extensible.
  **La config mail ne se met PAS ici** : elle est dans `conf-suivi-squad.json`. Des commentaires
  de cellule (A1/C1) le rappellent dans le classeur.

### Excel = données métier + config squad
Feuilles `Parametres`, `Membres`, `Congés`, `Historique`, `Log`. **Pas de VBA.**
`Historique` au **format long** : 1 ligne par rôle désigné (`SemaineLundi, Role, Nom,
Email, DateAnnonce, DateRappel, Statut, Note`). Listes déroulantes : `Rôle` =
`Parametres!$A$2:$A$50`, `Actif` = Oui/Non, `Congés!Membre` = noms de `Membres`. Lecture
tolérante : `ConvertTo-DateOuNull` / `ConvertTo-IntOuZero` ; lectures encadrées par `@(...)` ;
parcours jusqu'à la **fin de zone utilisée** (`Dimension.End.Row`, lignes vides intercalées
ignorées — un congé « effacé » ne masque plus la suite) ; noms **nettoyés** (`.Trim()`) dans
Membres, Congés et Historique (l'ajout dans Historique/Log se fait donc en fin de zone, pas
au premier emplacement vide).

### Désignation (par rôle)
Pour chaque rôle de `Parametres`, on désigne 1 personne : candidats actifs du rôle →
exclusion congés → exclusion du désigné précédent
de ce rôle (sauf seul candidat) → **tri** par ordre de priorité : 1) **min `NB_FOIS`**
(le moins souvent désigné — équité réelle sur le nombre de fois), 2) à égalité, **dernière
désignation la plus ancienne** (ceux qui n'ont **jamais** fait d'abord), 3) `Nom`. La « dernière désignation » est la
semaine la plus récente trouvée dans l'`Historique` pour la personne ; en repli (pas encore
d'historique) la graine `DateDernierSuivi` (col D de `Membres`, cf. `Get-DerniereDe`) ; sinon
jamais désigné (`[DateTime]::MinValue`, priorité absolue).
**Exclusion congés** : congés du classeur + absences CSV importées (cf. `AbsencesCsv`). Sur
la **semaine cible**, exclu si le nb de **jours ouvrés** (lundi→vendredi) en congé
**dépasse** `ToleranceCongesJours` (feuille `Parametres` C/D en priorité, repli sur
`conf-suivi-squad.json`, défaut 0 = exclu dès 1 jour ; cf. `Get-NbJoursCongesSemaine`).
Pondération : ligne `Demi…` (col D `Type` de `Congés`) = **0,5/jour** (poids max si
chevauchement) ; un **jour férié** (cf. `Test-JourFerie`) ne compte jamais. Le **vendredi
précédent** (jour de l'annonce) reste une exclusion **stricte** (`Test-MembreEnConges`),
non soumise à la tolérance — **neutralisée si ce vendredi est férié**. Même logique
répliquée dans `Select-Sim` (prévision). Qualité : noms de congés inconnus et dates
inversées → WARN (dates inversées : échangées).
**`NB_FOIS` = colonne `Compteur` (col E) = CUMUL VIVANT** : `Get-NbFoisDe` la lit telle quelle ;
le script l'**incrémente (+1 par désignation) et la réécrit dans `Membres`** (seule colonne
écrite dans cette feuille), **mais UNIQUEMENT au `Rappel`** : la désignation d'`Annonce`
(vendredi) est **prévisionnelle** (congés/changements possibles le week-end), donc elle
**ne touche pas au `Compteur`** ; seul le `Rappel` du lundi confirme et incrémente. La valeur
de départ saisie sert de point de comptage initial.
**Recalcul idempotent** (`Invoke-DesignationSemaine`) partagé par Annonce et Rappel : on
supprime les lignes de la semaine cible puis on redésigne depuis les données à jour. Pour le
`Compteur` (touché au seul `Rappel`), avant recalcul on **annule** (-1) l'incrément d'un
`Rappel` précédent (lignes dont la `DateRappel` est renseignée) puis on ré-incrémente → pas
de double comptage sur relance. Donc `Rappel` (lundi, semaine en cours)
reprend les congés ajoutés le week-end. Si un rôle n'a aucun candidat → **ALERTE atomique**
(aucune désignation commitée), mail aux `EmailsAdmin`.

### Envoi des mails : délégation au moteur CL
`Send-Notification` est le point d'envoi unique. Il **délègue toujours** à
`SendMailNotificationHTML.ps1` (lancé en **process séparé** : `pwsh`/`powershell.exe`
selon l'édition). **Le moteur (outil commun) n'est PAS modifié** : il lit TOUTE sa config
dans `-ConfigFile`. On lui passe donc la **config complète de la squad** (`$cfgSquad` =
son `conf-suivi-squad.json`, `TemplatePath` résolu, clés `_…` retirées) **sérialisée via
`ConvertTo-JsonSafe`** dans un fichier temporaire. Ainsi `From`/`Subject`/`SmtpServer`/
`Statuses`… sont propres **à la squad** sans toucher au moteur. Restent passés en ligne :
`-Status` (DESIGNATION/RAPPEL/ALERTE), `-NomJob` (squad), `-OverrideTo` (les **désignés**),
`-OverrideCc` (copies), `-KeyValues`, `-SectionFile`.
**Pas de repli** : si le moteur est introuvable, l'envoi échoue explicitement.
Le moteur ne connaît pas la **semaine cible** : `Send-Notification -Remplacements` injecte les
placeholders maison `{{SEMAINE}}` / `{{SEMAINE_NUM}}` / `{{PERIODE}}` (calculés depuis `$Lundi`)
dans le JSON sérialisé **avant** envoi (utilisables dans `Subject` et les messages `Statuses`).

Les sections (tableaux, notes) sont passées à `Send-Notification` en **objets** puis
sérialisées **une seule fois** via `ConvertTo-JsonSafe` (émetteur JSON maison). C'est
**volontaire** : `ConvertTo-Json` de PS 5.1 « déballe » les tableaux à un seul élément et
n'échappe pas comme PS 7, ce qui produisait un JSON invalide → sections silencieusement
perdues (mail réduit à la seule ligne « Désignation de la semaine »). Toute construction de
lignes de tableau utilise `… += , @(...)` puis `return , $rows` pour ne pas déballer une
ligne unique. **Ne pas réintroduire d'aller-retour `ConvertTo-Json`/`ConvertFrom-Json`.**

Le mail d'**annonce** inclut (via `-SectionsInline`) un tableau de **compteurs par rôle**
(`Get-CompteursRows <Role> <Lundi>`) et la **prévision des `SemainesPrevision` prochaines
semaines** (`Get-PrevisionRows`, 1 colonne par rôle, simulation de l'algo sur copies en
mémoire). Si la prévision montre `(aucun disponible)` pour un rôle, une **alerte
préventive de couverture** part aux `EmailsAdmin` (envoi non bloquant, annonce seulement).
**Règle d'affichage des compteurs** : une personne indisponible pour la semaine cible
(inactive, ou en congé au-delà de la tolérance / le vendredi de passation) ne doit **jamais**
apparaître en tête du tableau, même avec le compteur le plus bas — les désignables d'abord,
puis les indisponibles annotées « (en congé) ». L'exclusion congés s'applique donc PARTOUT :
désignation, prévision (`Select-Sim`) ET affichage des compteurs.

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
- **Encodage** : `SuiviExploitation.ps1` en **UTF-8 avec BOM** (accents sous PS 5.1) ; les
  `conf-suivi-squad.json` en **UTF-8** (BOM recommandé). **Conserver le BOM** lors des éditions
  (les outils peuvent le retirer — vérifier après coup). Le moteur `SendMailNotificationHTML.ps1`
  est en **ANSI/Windows-1252** (pas de BOM) — **ne pas le réencoder**. Pipeline UTF-8 de bout en bout.
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
