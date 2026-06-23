# mon_env

Dépôt du système **Suivi d'Exploitation Multi-Squads** : désignation hebdomadaire
automatique d'un développeur et d'un PO pour le suivi d'exploitation, scalable à
plusieurs squads dans une même entreprise.

Le projet repose sur **PowerShell 5.1+** (orchestration, désignation, envoi de mails,
lecture/écriture des classeurs) et **Excel** (membres, congés, historique). Il est
conçu pour **Windows** + Task Scheduler.

## Contenu du dépôt

Tout vit dans [`suiviexploit/`](suiviexploit/) :

| Fichier | Rôle |
|---|---|
| `SuiviExploitation.ps1` | **Script unique** : `-Action Annonce\|Rappel\|Test\|NouvelleSquad` |
| `config.psd1` | **Configuration unique** (SMTP + liste des squads), éditable au bloc-notes |
| `suivi_exploitation_template.xlsx` | Gabarit Excel pour onboarder une squad |
| `README.md` | Guide d'installation et d'exploitation complet |

## Démarrage

Guide complet : [`suiviexploit/README.md`](suiviexploit/README.md).

```powershell
# Valider la configuration
.\SuiviExploitation.ps1 -Action Test -DossierRacine "\\srv\suivi-exploitation"

# Onboarder une squad (crée le dossier + le classeur)
.\SuiviExploitation.ps1 -Action NouvelleSquad -DossierRacine "\\srv\suivi-exploitation"
```

Deux tâches planifiées suffisent : `-Action Annonce` le vendredi, `-Action Rappel` le lundi.
