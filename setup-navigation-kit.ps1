<#
.SYNOPSIS
  Installe le kit « chantier navigation » d'Hermes et prépare Playwright avec le Edge du poste.

.DESCRIPTION
  Étapes : vérifications (git, node, npm, Edge) -> branche git -> copie du kit -> .kilocodeignore
  -> fichiers Playwright dans le projet Angular -> npm install -> scripts package.json -> .gitignore
  -> contrôles finaux. Ne committe jamais : affiche la commande à lancer.
  Compatible Windows PowerShell 5.1 et PowerShell 7.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\setup-navigation-kit.ps1
.EXAMPLE
  .\setup-navigation-kit.ps1 -RepoPath C:\code\hermes -ZipPath C:\temp\hermes-navigation-kit-v2.zip -RunTests
.EXAMPLE
  .\setup-navigation-kit.ps1 -WhatIfOnly      # affiche ce qui serait fait, sans rien modifier
#>
[CmdletBinding()]
param(
  [string]$RepoPath = "$env:USERPROFILE\code\acode\hermes",
  [string]$ZipPath  = "$env:USERPROFILE\Downloads\hermes-navigation-kit-v2.zip",
  [string]$Branch   = "fix/navigation",
  [switch]$SkipBranch,      # rester sur la branche courante
  [switch]$SkipNpm,         # ne pas installer @playwright/test
  [switch]$PersistEnv,      # enregistrer PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD et NO_PROXY pour l'utilisateur
  [switch]$RunTests,        # lancer npm run e2e:nav à la fin (le back doit tourner)
  [switch]$AllowDirty,      # autoriser un dépôt avec des modifications non committées
  [switch]$WhatIfOnly       # simulation
)

$ErrorActionPreference = 'Stop'
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:Warnings = New-Object System.Collections.Generic.List[string]

function Step($msg)  { Write-Host "`n==> $msg" -ForegroundColor Cyan }
function Ok($msg)    { Write-Host "    [OK] $msg" -ForegroundColor Green }
function Info($msg)  { Write-Host "    $msg" -ForegroundColor Gray }
function Warn($msg)  { Write-Host "    [!] $msg" -ForegroundColor Yellow; $script:Warnings.Add($msg) }
function Fail($msg)  { Write-Host "    [X] $msg" -ForegroundColor Red; exit 1 }
function Do-Action([string]$label, [scriptblock]$action) {
  if ($WhatIfOnly) { Info "(simulation) $label" } else { & $action; Ok $label }
}
function Ensure-Lines([string]$file, [string[]]$lines, [string]$header) {
  $existing = @()
  if (Test-Path $file) { $existing = [IO.File]::ReadAllLines($file) }
  $missing = @($lines | Where-Object { $existing -notcontains $_ })
  if ($missing.Count -eq 0) { Ok "$(Split-Path $file -Leaf) déjà à jour"; return }
  Do-Action "$(Split-Path $file -Leaf) : ajout de $($missing.Count) ligne(s)" {
    $text = ""
    if ($existing.Count -gt 0 -and -not ([IO.File]::ReadAllText($file).EndsWith("`n"))) { $text += "`r`n" }
    if ($header) { $text += "`r`n$header`r`n" }
    $text += ($missing -join "`r`n") + "`r`n"
    [IO.File]::AppendAllText($file, $text, $Utf8NoBom)
  }
}

# ---------------------------------------------------------------------------
Step "1/9 Vérification des outils"
foreach ($tool in 'git','node','npm') {
  if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) { Fail "$tool introuvable dans le PATH" }
}
$nodeVersion = (node --version).TrimStart('v')
$nodeMajor = [int]($nodeVersion.Split('.')[0])
if ($nodeMajor -lt 18) { Fail "Node $nodeVersion trop ancien : Playwright demande Node 18 ou plus" }
Ok "git, node $nodeVersion, npm $(npm --version)"

$edgePaths = @(
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
)
$edge = $edgePaths | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if ($edge) { Ok "Edge trouvé : $((Get-Item $edge).VersionInfo.ProductVersion)" }
else { Warn "Edge introuvable : passer channel à 'chrome' dans playwright.config.ts" }

$registry = (npm config get registry).Trim()
Info "Registre npm : $registry"
if ($registry -like '*registry.npmjs.org*') { Warn "Registre npm public : en entreprise, l'installation peut être bloquée (vérifier le registre interne)" }

# ---------------------------------------------------------------------------
Step "2/9 Dépôt et branche"
if (-not (Test-Path (Join-Path $RepoPath '.git'))) { Fail "Pas de dépôt git dans $RepoPath (utiliser -RepoPath)" }
Set-Location $RepoPath
Ok "Dépôt : $RepoPath"
$dirty = git status --porcelain
if ($dirty -and -not $AllowDirty) { Fail "Modifications non committées. Committer/stasher, ou relancer avec -AllowDirty" }
if (-not $SkipBranch) {
  $current = (git rev-parse --abbrev-ref HEAD).Trim()
  if ($current -eq $Branch) { Ok "Déjà sur $Branch" }
  else {
    git show-ref --verify --quiet "refs/heads/$Branch"
    if ($LASTEXITCODE -eq 0) { Do-Action "Bascule sur la branche existante $Branch" { git checkout $Branch | Out-Null } }
    else { Do-Action "Création de la branche $Branch (depuis $current)" { git checkout -b $Branch | Out-Null } }
  }
}

# ---------------------------------------------------------------------------
Step "3/9 Extraction du kit"
if (-not (Test-Path $ZipPath)) { Fail "Zip introuvable : $ZipPath (utiliser -ZipPath)" }
$tmp = Join-Path $env:TEMP ("navkit_" + [guid]::NewGuid().ToString('N').Substring(0,8))
Expand-Archive -Path $ZipPath -DestinationPath $tmp -Force
$kit = Join-Path $tmp 'hermes-navigation-kit'
if (-not (Test-Path (Join-Path $kit 'docs\specs\navigation'))) { Fail "Contenu inattendu dans le zip (dossier hermes-navigation-kit/docs/specs/navigation absent)" }
Ok "Kit extrait dans $tmp"

# ---------------------------------------------------------------------------
Step "4/9 Copie des spécifications et de la règle Kilo"
Do-Action "docs/specs/navigation copié" {
  New-Item -ItemType Directory -Force -Path 'docs\specs' | Out-Null
  Copy-Item (Join-Path $kit 'docs\specs\navigation') 'docs\specs\' -Recurse -Force
}
Do-Action ".kilocode/rules/hermes-navigation.md copié" {
  New-Item -ItemType Directory -Force -Path '.kilocode\rules' | Out-Null
  Copy-Item (Join-Path $kit '.kilocode\rules\hermes-navigation.md') '.kilocode\rules\' -Force
}

# ---------------------------------------------------------------------------
Step "5/9 .kilocodeignore (fusion, jamais d'écrasement)"
$kiloIgnore = Join-Path $RepoPath '.kilocodeignore'
if (-not (Test-Path $kiloIgnore)) {
  Do-Action ".kilocodeignore créé depuis le kit" { Copy-Item (Join-Path $kit '.kilocodeignore') $kiloIgnore }
} else {
  Info ".kilocodeignore existant conservé"
}
Ensure-Lines $kiloIgnore @(
  '!docs/specs/maquettes/*.png',
  '!docs/specs/navigation/maquettes/*.png',
  'playwright-report/',
  'test-results/'
) '# --- ajouté par setup-navigation-kit.ps1 ---'

# ---------------------------------------------------------------------------
Step "6/9 Projet Angular et fichiers Playwright"
$angularJson = Get-ChildItem -Path $RepoPath -Filter angular.json -Recurse -Depth 4 -ErrorAction SilentlyContinue |
  Where-Object { $_.FullName -notmatch '\\node_modules\\' } | Select-Object -First 1
if (-not $angularJson) { Fail "angular.json introuvable sous $RepoPath" }
$front = $angularJson.DirectoryName
Ok "Front Angular : $front"

$srcConfig = Join-Path $kit 'frontend-e2e\playwright.config.ts'
$dstConfig = Join-Path $front 'playwright.config.ts'
if (Test-Path $dstConfig) { Warn "playwright.config.ts existe déjà : non remplacé (modèle disponible dans $srcConfig)" }
else { Do-Action "playwright.config.ts installé" { Copy-Item $srcConfig $dstConfig } }

$srcSpec = Join-Path $kit 'frontend-e2e\e2e\navigation'
$dstSpecDir = Join-Path $front 'e2e\navigation'
if (Test-Path (Join-Path $dstSpecDir 'smoke.spec.ts')) { Warn "e2e/navigation/smoke.spec.ts existe déjà : non remplacé" }
else { Do-Action "e2e/navigation/smoke.spec.ts installé" {
  New-Item -ItemType Directory -Force -Path $dstSpecDir | Out-Null
  Copy-Item (Join-Path $srcSpec '*') $dstSpecDir -Recurse -Force
} }

# ---------------------------------------------------------------------------
Step "7/9 Installation de Playwright (sans télécharger de navigateur)"
$env:PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD = '1'
if ($env:NO_PROXY) { if ($env:NO_PROXY -notmatch 'localhost') { $env:NO_PROXY = "$env:NO_PROXY,localhost,127.0.0.1" } }
else { $env:NO_PROXY = 'localhost,127.0.0.1' }
if ($PersistEnv) {
  Do-Action "Variables enregistrées pour l'utilisateur" {
    [Environment]::SetEnvironmentVariable('PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD', '1', 'User')
    [Environment]::SetEnvironmentVariable('NO_PROXY', $env:NO_PROXY, 'User')
  }
}
Push-Location $front
try {
  if ($SkipNpm) { Info "Installation npm ignorée (-SkipNpm)" }
  elseif (Test-Path 'node_modules\@playwright\test') { Ok "@playwright/test déjà installé" }
  else {
    Do-Action "@playwright/test installé" {
      npm install --save-dev @playwright/test
      if ($LASTEXITCODE -ne 0) { throw "npm install a échoué (registre interne ? proxy ?)" }
    }
  }

  Step "8/9 Scripts npm et .gitignore"
  $scripts = [ordered]@{
    'e2e:nav' = 'playwright test e2e/navigation'
    'e2e:ui'  = 'playwright test --ui'
    'e2e:rec' = 'playwright codegen --channel=msedge http://localhost:4200'
  }
  foreach ($k in $scripts.Keys) {
    $current = (npm pkg get "scripts.$k" | Out-String).Trim()
    if ($current -and $current -ne '{}' -and $current -ne '') { Ok "script $k déjà présent" }
    else { Do-Action "script $k ajouté" { npm pkg set "scripts.$k=$($scripts[$k])" | Out-Null } }
  }
} finally { Pop-Location }

Ensure-Lines (Join-Path $RepoPath '.gitignore') @('playwright-report/', 'test-results/', 'playwright/.auth/') '# Playwright'

# ---------------------------------------------------------------------------
Step "9/9 Contrôles finaux"
if (Test-Path '.kilocode\rules\hermes-navigation.md') { Ok "Règle Kilo présente (recharger VS Code si Kilo ne la voit pas)" }
$png = 'docs\specs\navigation\maquettes\navigation.png'
if (Test-Path $png) { Ok "Maquette présente" }
else { Warn "Maquette absente : exporter l'écran « Navigation » du canevas vers $png" }

if ($RunTests -and -not $WhatIfOnly) {
  Push-Location $front
  try {
    Info "Lancement des tests (le back doit tourner en auth de dev)…"
    npm run e2e:nav
    if ($LASTEXITCODE -ne 0) { Warn "Des tests échouent : normal avant le chantier, c'est le diagnostic de départ (npx playwright show-report)" }
  } finally { Pop-Location }
}

Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n================ Résumé ================" -ForegroundColor Cyan
if ($script:Warnings.Count -eq 0) { Write-Host "Aucun avertissement." -ForegroundColor Green }
else { $script:Warnings | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow } }
Write-Host @"

Étapes suivantes :
  1. Démarrer le back en auth de dev (HERMES_DEV_ROLE=manager), puis :  cd "$front"; npm run e2e:nav
  2. Vérifier le diff :  git status ; git diff
  3. Committer :         git add -A ; git commit -m "chore: kit chantier navigation + Playwright (Edge)"
  4. Kilo, nouvelle tâche, mode Architect :  Exécute @docs/specs/navigation/N0_diagnostic.md
"@
