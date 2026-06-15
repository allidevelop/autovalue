param(
  [string]$AppName = "Realtify",
  [switch]$OneFile,
  [switch]$SkipBrowserInstall
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $Python)) {
  throw "Virtual environment not found: $Python"
}

& $Python -m pip install pyinstaller==6.11.1
if ($LASTEXITCODE -ne 0) {
  throw "PyInstaller install failed"
}

$browserRoot = Join-Path $Root "tools\ms-playwright"
if (-not $SkipBrowserInstall) {
  $chromiumExists = Test-Path -LiteralPath $browserRoot
  if ($chromiumExists) {
    $chromiumExists = [bool](Get-ChildItem -LiteralPath $browserRoot -Directory -Filter "chromium*" -ErrorAction SilentlyContinue | Select-Object -First 1)
  }
  if (-not $chromiumExists) {
    New-Item -ItemType Directory -Force -Path $browserRoot | Out-Null
    $oldBrowserPath = $env:PLAYWRIGHT_BROWSERS_PATH
    try {
      $env:PLAYWRIGHT_BROWSERS_PATH = $browserRoot
      & $Python -m playwright install chromium
      if ($LASTEXITCODE -ne 0) {
        throw "Playwright Chromium install failed"
      }
    } finally {
      $env:PLAYWRIGHT_BROWSERS_PATH = $oldBrowserPath
    }
  }
}

$dist = Join-Path $Root "dist"
$work = Join-Path $Root "build\pyinstaller"
$entry = Join-Path $Root "scripts\run_windows_app.py"

$args = @(
  "-m", "PyInstaller",
  "--name", $AppName,
  "--noconfirm",
  "--noconsole",
  "--distpath", $dist,
  "--workpath", $work,
  "--paths", (Join-Path $Root "src"),
  "--collect-data", "playwright",
  "--collect-submodules", "playwright",
  "--hidden-import", "win32timezone",
  "--add-data", "$(Join-Path $Root 'config');config",
  "--add-data", "$(Join-Path $Root 'tools\poppler');tools\poppler",
  "--add-data", "$(Join-Path $Root 'tools\tessdata');tools\tessdata"
)

if (Test-Path -LiteralPath $browserRoot) {
  $args += @("--add-data", "${browserRoot};tools\ms-playwright")
}

$defaultExcelTemplate = Get-ChildItem -LiteralPath $Root -File -Filter "*.xls" |
  Where-Object { $_.Name -like "*15*.xls" } |
  Select-Object -First 1
if ($defaultExcelTemplate) {
  $args += @("--add-data", "$($defaultExcelTemplate.FullName);.")
}

if ($OneFile) {
  $args += "--onefile"
} else {
  $args += "--onedir"
}

$args += $entry
& $Python @args

Write-Host "Build complete. Output: $dist"
