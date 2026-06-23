#!/bin/bash
# Hook SessionStart — prépare l'environnement de validation pour Claude Code sur le web.
#
# Ce projet est un système PowerShell 5.1 + Excel/VBA conçu pour Windows.
# Il n'a pas de tests automatisés ; la seule validation outillable dans un
# conteneur Linux est le linter PowerShell (PSScriptAnalyzer), qui nécessite pwsh.
#
# Le hook installe donc, en session distante uniquement :
#   - PowerShell 7 (pwsh)
#   - le module PSScriptAnalyzer
# Il est idempotent : il ne réinstalle pas ce qui est déjà présent.
set -euo pipefail

# Ne rien faire en local : ce projet tourne nativement sous Windows.
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

PWSH_VERSION="7.4.6"

# --- PowerShell ---
if ! command -v pwsh >/dev/null 2>&1; then
  echo "Installation de PowerShell ${PWSH_VERSION}..."
  TMP_DEB="$(mktemp --suffix=.deb)"
  curl -fsSL -o "$TMP_DEB" \
    "https://github.com/PowerShell/PowerShell/releases/download/v${PWSH_VERSION}/powershell_${PWSH_VERSION}-1.deb_amd64.deb"
  apt-get install -y "$TMP_DEB"
  rm -f "$TMP_DEB"
else
  echo "PowerShell déjà présent : $(pwsh --version)"
fi

# --- PSScriptAnalyzer (linter) ---
pwsh -NoProfile -Command '
  if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer)) {
    Write-Host "Installation de PSScriptAnalyzer..."
    Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -ErrorAction Stop
  } else {
    Write-Host ("PSScriptAnalyzer déjà présent : " + (Get-Module -ListAvailable PSScriptAnalyzer)[0].Version)
  }
'

echo "Environnement de validation prêt."
echo "Linter : pwsh -NoProfile -Command \"Invoke-ScriptAnalyzer -Path ./suiviexploit -Recurse\""
