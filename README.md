# mon_env — Suivi d'Exploitation Multi-Squads

Système qui désigne **automatiquement chaque semaine**, pour chaque squad, la ou les
personnes « de suivi d'exploitation » (1 personne par rôle : **solo / binôme / trio…**),
puis envoie le mail d'**annonce** (vendredi) et le mail de **rappel** (lundi).

- **PowerShell 5.1+** pour l'orchestration, la désignation et la lecture/écriture Excel.
- **Excel** pour les données métier (membres, congés, historique) et la configuration
  propre à chaque squad.
- Envoi des mails **délégué** au moteur commun `SendMailNotificationHTML.ps1` (charte HTML).
- Conçu pour **Windows + Task Scheduler**, sur un partage réseau.

## Contenu du dépôt

Tout vit dans [`suiviexploit/`](suiviexploit/) :

| Fichier | Rôle |
|---|---|
| `SuiviExploitation.ps1` | Script unique : `-Action Annonce \| Rappel \| Test \| NouvelleSquad` |
| `config-suivi.json` | Configuration **globale**, partagée avec le moteur mail |
| `SendMailNotificationHTML.ps1` | Moteur d'envoi de mails HTML (dépendance requise) |
| `template-notification.html` | Gabarit HTML des mails |
| `suivi_exploitation_template.xlsx` | Gabarit d'un classeur de squad |
| `suivi_exploitation_exemple.xlsx` | Exemple rempli |
| `README.md` | Guide d'installation et d'exploitation détaillé |

## Démarrage rapide

Guide complet : [`suiviexploit/README.md`](suiviexploit/README.md).

```powershell
# Créer une squad (dossier + classeur + ajout dans la config)
.\SuiviExploitation.ps1 -Action NouvelleSquad -DossierRacine "\\srv\suivi-exploitation"

# Valider la configuration et les classeurs
.\SuiviExploitation.ps1 -Action Test -DossierRacine "\\srv\suivi-exploitation"
```

Deux tâches planifiées suffisent : `-Action Annonce` le vendredi, `-Action Rappel` le lundi.
