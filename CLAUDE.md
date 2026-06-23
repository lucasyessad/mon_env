# CLAUDE.md

Ce fichier guide Claude Code (claude.ai/code) lorsqu'il travaille sur ce dépôt.

## Vue d'ensemble

**Suivi d'Exploitation Multi-Squads** : système qui désigne automatiquement, chaque
semaine, un développeur et un PO « de suivi d'exploitation » pour chaque squad, et
envoie les mails d'annonce (vendredi) et de rappel (lundi).

Le projet est conçu pour **Windows** (PowerShell 5.1 natif + Excel), exécuté via le
**Task Scheduler** sur un partage réseau. Il n'y a **pas de build ni de tests
automatisés** : la validation se fait avec `-Action Test` et des exécutions manuelles
en mode `Annonce`/`Rappel`.

Tout le code vit dans `suiviexploit/`. La documentation de référence est
`suiviexploit/README.md` — **la consulter avant toute modification de comportement.**

## Structure (volontairement minimale)

```
suiviexploit/
├── SuiviExploitation.ps1            tout le code (script unique)
├── config.psd1                      configuration unique (exemple/template)
├── suivi_exploitation_template.xlsx gabarit Excel d'une squad
└── README.md                        guide d'installation/exploitation
```

Déploiement sur le partage : le script, `config.psd1` et le template à la racine ;
un dossier par squad contenant `suivi_exploitation.xlsx` + `historique\` ; `_logs\`
généré pour les logs orchestrateur.

## Le script unique

`SuiviExploitation.ps1 -Action <Annonce|Rappel|Test|NouvelleSquad>` plus
`-DossierRacine`, `-Config`, `-Squad`. Organisation interne :
- Fonctions communes : `Get-Configuration`, `Send-Mail`, `Format-Sujet`, `Build-Corps*`.
- `Invoke-SquadJob` : traite UNE squad (lock, lecture, désignation, mail, swap, rétention)
  et **retourne un objet résultat** (plus de protocole texte parsé par regex).
- `Invoke-Orchestration` : boucle sur les squads (try/catch par squad), rapport dans
  `_logs\`, mail récap **uniquement** en cas d'anomalie.
- `Invoke-Test` : validation. `Invoke-NouvelleSquad` : onboarding.

## Concepts clés

### Configuration unique (`config.psd1`)
- Un seul fichier `.psd1` (PowerShell Data File), chargé par `Import-PowerShellDataFile`.
- Contient le SMTP global, les sujets de mail, les destinataires de l'orchestrateur, et
  la liste `Squads` (un bloc par squad : `Nom`, `Classeur`, expéditeur, responsables, copies).
- **Plus de fusion global/local** : une squad peut surcharger un sujet en ajoutant la clé
  dans son bloc. Clés obligatoires validées par `$script:ClesGlobales` et `$script:ClesSquad`.
- La liste des squads vient de la config (plus de scan de dossiers `_`).

### Excel = données métier seulement
Feuilles `Membres`, `Congés`, `Historique`, `Log`. **Plus de feuille Paramètres lue**,
**plus de VBA**. Les congés se saisissent directement (feuille Congés, données dès la
ligne 6 ; l'en-tête est en ligne 5, héritage du gabarit). Mappage des colonnes inchangé
par rapport à l'historique du projet — **ne pas le casser** (compat avec les classeurs
existants).

### Algorithme de désignation (`Get-DesignePourRole`, inchangé)
1. Candidats = membres actifs (`Actif = Oui`) du rôle.
2. Exclusion des membres en congé sur la semaine cible **ou** le vendredi précédent.
3. Exclusion du désigné de la semaine précédente (sauf s'il est le seul candidat).
4. Tri `DateDernierSuivi` ↑, `Compteur` ↑, `Nom` ↑ ; le premier gagne.
Nouveaux membres initialisés à la **médiane** des dates du rôle (`Initialize-NouveauxMembres`).

### Accès concurrent au classeur
Verrou `<squad>\suivi_exploitation.lock` pendant le traitement ; attente du verrou Excel
`~$*` (6× / 30 s) ; écriture transactionnelle (`.tmp` → archive l'ancien → renomme).
Note : il n'y a plus de protection côté Excel (VBA supprimé) ; le `.lock` + l'attente `~$`
restent la seule protection.

## Conventions de modification

- **Langue** : tout (code, commentaires, messages, mails) en **français**.
- **PowerShell 5.1** : `Set-StrictMode -Version Latest` + `$ErrorActionPreference='Stop'`.
  Pas de syntaxe PS7-only.
- **Encodage** : `SuiviExploitation.ps1` et `config.psd1` sont en **UTF-8 avec BOM**
  (indispensable aux accents sous PS 5.1). **Conserver le BOM** lors des éditions
  (les outils d'édition peuvent le retirer — vérifier après coup).
- **Dépendance** : module `ImportExcel` (installé auto en `-Scope CurrentUser`).
- **Rétention** : `$script:RetentionJours` (28) en haut du script — **un seul endroit**
  désormais (archives squad + logs orchestrateur).
- **SMTP** : anonyme par défaut, auth non gérée.
- Ne pas committer d'artefacts d'exécution (`_logs/`, `historique/`, `*.lock`, `~$*`,
  `*.journal.log`) — couverts par `.gitignore`.

## Validation avant livraison

Pas de suite de tests. Pour valider un changement :
1. Lint : `pwsh -NoProfile -Command "Invoke-ScriptAnalyzer -Path ./suiviexploit"` (le hook
   `.claude/hooks/session-start.sh` installe pwsh + PSScriptAnalyzer en session web).
   Référence actuelle : 0 erreur ; warnings résiduels = `Write-Host` (sortie console
   interactive, assumé) + quelques noms de fonctions au pluriel (helpers métier).
2. `-Action Test -DossierRacine <...>` doit être **vert** (code 0).
3. Exécution manuelle `-Action Annonce` puis vérifier `_logs\`, `Historique`,
   `<squad>\historique\`, et les mails.

> Les validations 2-3 nécessitent **Windows + Excel** ; non exécutables en conteneur Linux.
