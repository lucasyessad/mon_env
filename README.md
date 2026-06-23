# mon_env

Dépôt du système **Suivi d'Exploitation Multi-Squads** : désignation hebdomadaire
automatique d'un développeur et d'un PO pour le suivi d'exploitation, scalable à
plusieurs squads dans une même entreprise.

Le projet repose sur **PowerShell 5.1+** (orchestration, envoi de mails, lecture/écriture
des classeurs) et **Excel + VBA** (saisie des membres, congés, historique côté squad).

## Contenu du dépôt

Tout le code et la documentation détaillée se trouvent dans [`suiviexploit/`](suiviexploit/) :

| Fichier | Rôle |
|---|---|
| `Invoke-AllSquads.ps1` | Orchestrateur appelé par le Task Scheduler ; boucle sur les squads |
| `Invoke-SuiviExploitation.ps1` | Job d'une squad : fusion config, désignation, mail, historique |
| `New-SquadFolder.ps1` | Onboarding d'une nouvelle squad |
| `Test-Configuration.ps1` | Vérification de la configuration avant mise en prod |
| `mdlSuiviExploitation.bas` | Module VBA (macro `AjouterConge`, raccourcis) |
| `ThisWorkbook_CodeAColler.txt` | Code VBA `ThisWorkbook` (verrou + raccourcis) |
| `parametres_globaux.xlsx` | Config globale entreprise (SMTP, sujets de mail) |
| `suivi_exploitation_template.xlsx` | Template Excel pour onboarder une squad |
| `SuiviExploitation_AllSquads_*.xml` | Définitions des tâches planifiées (Annonce / Rappel) |

## Démarrage

Voir le guide d'installation et d'exploitation complet : [`suiviexploit/README.md`](suiviexploit/README.md).

Étapes principales :
1. Créer la structure de dossiers sur le partage réseau (`_config`, `_scripts`, `_logs`).
2. Déployer les fichiers et configurer `parametres_globaux.xlsx`.
3. Onboarder une squad avec `New-SquadFolder.ps1`, puis importer le VBA.
4. Valider avec `Test-Configuration.ps1`.
5. Enregistrer les 2 tâches planifiées (vendredi *Annonce*, lundi *Rappel*).
