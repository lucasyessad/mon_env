# CLAUDE.md

Ce fichier guide Claude Code (claude.ai/code) lorsqu'il travaille sur ce dépôt.

## Vue d'ensemble

**Suivi d'Exploitation Multi-Squads** : système qui désigne automatiquement, chaque
semaine, un développeur et un PO « de suivi d'exploitation » pour chaque squad, et
envoie les mails d'annonce (vendredi) et de rappel (lundi).

Le projet est conçu pour **Windows** (PowerShell 5.1 natif + Excel/VBA), exécuté
via le **Task Scheduler** sur un partage réseau. Il n'y a **pas de build ni de tests
automatisés** : la validation se fait avec `Test-Configuration.ps1` et des exécutions
manuelles en mode `Annonce`/`Rappel`.

Tout le code applicatif vit dans `suiviexploit/`. La documentation fonctionnelle et
opérationnelle de référence est `suiviexploit/README.md` — **la consulter avant toute
modification de comportement.**

## Architecture d'exécution

```
<DossierRacine>\               (partage réseau, ex. \\srv\suivi-exploitation)
├── _config\                   parametres_globaux.xlsx (+ template)
├── _scripts\                  les .ps1
├── _logs\                     logs orchestrateur (généré)
└── <squad>\                   1 dossier par squad : suivi_exploitation.xlsm + historique\
```

Les dossiers commençant par `_` sont **ignorés** par l'orchestrateur (réservés à
config / logs / scripts / archives). Pour archiver une squad : renommer son dossier
avec un préfixe `_`.

### Flux

1. **Task Scheduler** lance `Invoke-AllSquads.ps1 -Mode <Annonce|Rappel> -DossierRacine <...>`.
2. L'orchestrateur boucle sur les squads et appelle, dans un `try/catch` isolé,
   `Invoke-SuiviExploitation.ps1` pour chacune.
3. Chaque job : fusionne la config, désigne, envoie le mail, archive le classeur,
   met à jour membres/historique.
4. L'orchestrateur compile un rapport (texte + HTML), l'écrit dans `_logs\`, et
   envoie un mail récapitulatif **uniquement** en cas d'erreur ou de warning.

## Concepts clés

### Configuration globale + locale (fusion)
- Globale : `_config\parametres_globaux.xlsx` (SMTP, sujets de mail par défaut, alertes).
- Locale : feuille **Paramètres** du `.xlsm` de chaque squad (NomSquad, expéditeur,
  responsables, copies).
- **Priorité : local > global.** Une clé obligatoire absente des *deux* fichiers fait
  échouer le job avec une erreur explicite. Les clés obligatoires sont listées dans
  `$script:ClesObligatoires` (`Invoke-SuiviExploitation.ps1`) et `$ClesGlobalesAttendues`
  (`Test-Configuration.ps1`) — **garder ces deux listes cohérentes**.

### Algorithme de désignation (`Get-DesignePourRole`)
Pour un rôle (`Dev` ou `PO`) et un lundi cible :
1. Candidats = membres actifs (`Actif = Oui`) du rôle.
2. Exclusion des membres en congé sur la semaine cible **ou** le vendredi précédent.
3. Exclusion du désigné de la semaine précédente (sauf s'il est le seul candidat).
4. Tri par `DateDernierSuivi` ↑, puis `Compteur` ↑, puis `Nom` ↑ ; le premier est choisi.
Les nouveaux membres (sans `DateDernierSuivi`) sont initialisés à la **médiane** des
dates du rôle (`Initialize-NouveauxMembres`) pour ne pas être désignés en priorité.

### Sécurité d'accès concurrent au classeur
- Verrou `<squad>\suivi_exploitation.lock` posé par le job pendant le traitement.
- Le job attend la disparition du verrou Excel `~$*.xlsm` (retry 6× / 30s).
- Côté Excel, `Workbook_Open` (dans `ThisWorkbook_CodeAColler.txt`) détecte le `.lock`
  et ferme le fichier avec un message si un traitement est en cours.
- Écriture transactionnelle : copie `.tmp` → modifs → archive l'ancien → renomme `.tmp`.

### VBA côté squad
- `mdlSuiviExploitation.bas` : macro `AjouterConge` (raccourci `Ctrl+Shift+C`),
  `SupprimerCongesPasses`, init/désactivation des raccourcis.
- `ThisWorkbook_CodeAColler.txt` : `Workbook_Open` / `Workbook_BeforeClose`.
- Ces deux éléments doivent être importés manuellement dans chaque classeur squad
  (voir étape 5 de `suiviexploit/README.md`), puis le fichier enregistré en `.xlsm`.

## Conventions de modification

- **Langue** : tout le code, les commentaires, les messages utilisateur et les mails
  sont en **français**. S'y conformer.
- **PowerShell** : tous les scripts utilisent `Set-StrictMode -Version Latest` et
  `$ErrorActionPreference = 'Stop'`. Préserver ces garde-fous ; le code doit rester
  compatible **PowerShell 5.1** (pas de syntaxe PS 7 uniquement).
- **Dépendance** : module `ImportExcel` (installé automatiquement en `-Scope CurrentUser`
  au premier run). Toute lecture/écriture de classeur passe par lui.
- **Rétention** : 4 semaines, codée en dur via `(Get-Date).AddDays(-28)` dans
  `Invoke-SuiviExploitation.ps1` et `Invoke-AllSquads.ps1` — modifier les deux ensemble.
- **SMTP** : anonyme par défaut. L'auth n'est pas gérée (voir « Limitations » du README).
- Ne pas committer d'artefacts d'exécution (`_logs/`, `historique/`, `*.lock`, `~$*`,
  `*.journal.log`) — ils sont déjà couverts par `.gitignore`.

## Validation avant livraison

Il n'existe pas de suite de tests. Pour valider un changement de comportement :
1. `Test-Configuration.ps1 -DossierRacine <...>` doit être **vert** (structure, clés,
   SMTP, ≥1 Dev actif et ≥1 PO actif par squad ; code de sortie 0).
2. Exécution manuelle : `Invoke-AllSquads.ps1 -Mode Annonce -DossierRacine <...>` et
   vérifier le log dans `_logs\`, les lignes ajoutées dans `Historique`, la copie datée
   dans `<squad>\historique\` et les mails reçus.

> Ces validations nécessitent un environnement **Windows + Excel** ; elles ne sont pas
> exécutables dans un conteneur Linux.
