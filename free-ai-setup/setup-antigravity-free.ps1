# Antigravity and Antigravity IDE: add the OpenRouter free models.
# Reads the key already saved for Claude Code. Does not print the key.
$ErrorActionPreference = "Stop"

function Get-OpenRouterKey {
  $settingsPath = Join-Path $env:USERPROFILE ".claude\settings.json"
  if (Test-Path $settingsPath) {
    $settings = Get-Content -Raw -Encoding UTF8 $settingsPath | ConvertFrom-Json
    $saved = $settings.env.ANTHROPIC_AUTH_TOKEN
    if ($saved -match '^sk-or-v1-[A-Za-z0-9]+$') { return $saved }
  }

  $candidates = @(
    (Join-Path $env:USERPROFILE "Downloads\Context7 Secret Key.txt"),
    "D:\ai\AI 셋팅\Context7 Secret Key.txt",
    "D:\ai\Context7 Secret Key.txt"
  )
  foreach ($path in $candidates) {
    if (-not (Test-Path $path)) { continue }
    $raw = Get-Content -Raw -Encoding UTF8 $path
    if ($raw -match 'sk-or-v1-[A-Za-z0-9]+') { return $Matches[0] }
  }
  throw "OpenRouter key not found in $settingsPath or Context7 Secret Key.txt"
}

function Find-InstallDirs {
  $roots = @(
    (Join-Path $env:LOCALAPPDATA "Programs"),
    $env:ProgramFiles,
    ${env:ProgramFiles(x86)},
    $env:LOCALAPPDATA
  ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

  $found = New-Object System.Collections.Generic.List[string]
  foreach ($root in $roots) {
    $dirs = Get-ChildItem -Path $root -Directory -ErrorAction SilentlyContinue
    foreach ($dir in $dirs) {
      if ($dir.Name -notmatch 'Antigravity') { continue }
      if (Test-Path (Join-Path $dir.FullName "resources")) { [void]$found.Add($dir.FullName) }
    }
  }
  return @($found | Select-Object -Unique)
}

function Test-NodeVersion {
  $node = Get-Command node -ErrorAction SilentlyContinue
  if (-not $node) { throw "Node.js 22.13 or newer is required. Install it from https://nodejs.org" }
  $version = node -p "process.versions.node"
  $parts = $version.Split(".")
  $major = [int]$parts[0]
  $minor = [int]$parts[1]
  if ($major -lt 22 -or ($major -eq 22 -and $minor -lt 13)) {
    throw "Node.js 22.13 or newer is required. Current: $version"
  }
  "NODE=$version"
}

$key = Get-OpenRouterKey
Test-NodeVersion
if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "Git is required." }

$repo = Join-Path $env:USERPROFILE "antigravity-add-model"
if (Test-Path (Join-Path $repo ".git")) {
  git -C $repo pull --ff-only
  if ($LASTEXITCODE -ne 0) { throw "Could not update $repo" }
} else {
  git clone https://github.com/vahapogut/antigravity-add-model.git $repo
  if ($LASTEXITCODE -ne 0) { throw "Could not clone antigravity-add-model" }
}

Push-Location $repo
npm ci --ignore-scripts
if ($LASTEXITCODE -ne 0) { throw "npm ci failed" }
npm run build
if ($LASTEXITCODE -ne 0) { throw "npm run build failed" }

Get-Process -ErrorAction SilentlyContinue |
  Where-Object { $_.ProcessName -like "Antigravity*" } |
  Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

$installs = @(Find-InstallDirs)
$applied = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
foreach ($dir in $installs) {
  "CHECK=$dir"
  npm run ide:check -- --resources $dir
  if ($LASTEXITCODE -eq 0) {
    npm run ide:apply -- --resources $dir
    if ($LASTEXITCODE -ne 0) { throw "IDE apply failed for $dir" }
    [void]$applied.Add("IDE $dir")
    continue
  }
  powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo "deploy.ps1") --check --resources $dir
  if ($LASTEXITCODE -eq 0) {
    powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo "deploy.ps1") --resources $dir
    if ($LASTEXITCODE -ne 0) { throw "Desktop apply failed for $dir" }
    [void]$applied.Add("DESKTOP $dir")
    continue
  }
  [void]$skipped.Add($dir)
}
Pop-Location

$dataDir = Join-Path $env:USERPROFILE ".gemini\antigravity"
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }

$specs = @(
  @{ name = "models/or-free"; display = "OpenRouter Free"; model = "openrouter/free"; context = 131072 },
  @{ name = "models/gemma-4-31b"; display = "Gemma 4 31B"; model = "google/gemma-4-31b-it:free"; context = 131072 },
  @{ name = "models/gemma-4-26b"; display = "Gemma 4 26B"; model = "google/gemma-4-26b-a4b-it:free"; context = 131072 },
  @{ name = "models/nemotron-ultra"; display = "Nemotron 3 Ultra"; model = "nvidia/nemotron-3-ultra-550b-a55b:free"; context = 1000000 },
  @{ name = "models/nemotron-super"; display = "Nemotron 3.5 Super"; model = "nvidia/nemotron-3.5-super-vl:free"; context = 1000000 },
  @{ name = "models/nemotron-lightning"; display = "Nemotron 3.5 Lightning"; model = "nvidia/nemotron-3.5-lightning:free"; context = 1000000 },
  @{ name = "models/nemotron-nano"; display = "Nemotron 3 Nano"; model = "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free"; context = 131072 },
  @{ name = "models/north-mini-code"; display = "North Mini Code"; model = "cohere/north-mini-code:free"; context = 131072 },
  @{ name = "models/flint"; display = "Flint"; model = "recursive/flint-v1.0:free"; context = 131072 },
  @{ name = "models/xing-4"; display = "Xing 4.0"; model = "xingchen-agi/xing4.0-29b-a4b:free"; context = 131072 },
  @{ name = "models/dots3-note"; display = "Dots3 Note"; model = "dots-studio/dots-3-note-preview:free"; context = 131072 }
)
$models = foreach ($spec in $specs) {
  $entry = [ordered]@{
    name = $spec.name
    displayName = $spec.display
    provider = "openrouter"
    apiFormat = "openai"
    apiUrl = "https://openrouter.ai/api/v1/chat/completions"
    apiKey = $key
    externalModelName = $spec.model
    enabled = $true
    contextWindow = $spec.context
    maxOutputTokens = 8192
    timeout = 120000
    maxRetries = 2
  }
  if ($spec.name -ne "models/or-free") { $entry.fallbackModels = @("models/or-free") }
  [pscustomobject]$entry
}
$importPath = Join-Path $dataDir "openrouter-free-import.json"
$json = @{ models = @($models) } | ConvertTo-Json -Depth 8
$utf8 = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($importPath, $json, $utf8)

if ($applied.Count -gt 0) {
  Start-Process powershell -ArgumentList @(
    "-NoExit",
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-Command",
    "Set-Location '$repo'; npm run proxy:start -- --open"
  )
}

"KEY_SET=True"
"REPO=$repo"
"IMPORT_FILE=$importPath"
"INSTALLS=$($installs.Count)"
"APPLIED=$($applied -join ' | ')"
"SKIPPED=$($skipped -join ' | ')"
if ($applied.Count -eq 0) {
  "NEXT=No supported Antigravity install was changed. Close both apps and rerun after confirming the install folder, or use Cline inside Antigravity IDE."
} else {
  "NEXT=In the browser window, choose Import and select the import file. Then pick a saved model in Antigravity. Delete the import file after Import succeeds."
}
