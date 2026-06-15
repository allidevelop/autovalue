$ErrorActionPreference = "Stop"

$Root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")
$Tools = Join-Path $Root "tools"
$Downloads = Join-Path $Tools "_downloads"
$Tessdata = Join-Path $Tools "tessdata"
$PopplerDir = Join-Path $Tools "poppler"

New-Item -ItemType Directory -Force -Path $Downloads, $Tessdata, $PopplerDir | Out-Null

$tesseractCandidates = @(
  "C:\Program Files\Tesseract-OCR\tesseract.exe",
  "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe"
)
$tesseract = $tesseractCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $tesseract) {
  winget install --id UB-Mannheim.TesseractOCR -e --source winget --silent --accept-package-agreements --accept-source-agreements
  $tesseract = $tesseractCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}

$popplerZip = Join-Path $Downloads "poppler-windows-25.07.0-0.zip"
$popplerMarker = Join-Path $PopplerDir ".extracted"
if (-not (Test-Path -LiteralPath $popplerZip)) {
  Invoke-WebRequest `
    -Uri "https://github.com/oschwartz10612/poppler-windows/releases/download/v25.07.0-0/Release-25.07.0-0.zip" `
    -OutFile $popplerZip
}
if (-not (Test-Path -LiteralPath $popplerMarker)) {
  Expand-Archive -LiteralPath $popplerZip -DestinationPath $PopplerDir -Force
  Set-Content -Path $popplerMarker -Value "ok" -Encoding ASCII
}

foreach ($lang in @("ukr", "rus")) {
  $out = Join-Path $Tessdata "$lang.traineddata"
  if (-not (Test-Path -LiteralPath $out)) {
    Invoke-WebRequest `
      -Uri "https://github.com/tesseract-ocr/tessdata_fast/raw/main/$lang.traineddata" `
      -OutFile $out
  }
}

$globalTessdata = "C:\Program Files\Tesseract-OCR\tessdata"
foreach ($lang in @("eng", "osd")) {
  $source = Join-Path $globalTessdata "$lang.traineddata"
  $dest = Join-Path $Tessdata "$lang.traineddata"
  if ((Test-Path -LiteralPath $source) -and -not (Test-Path -LiteralPath $dest)) {
    Copy-Item -LiteralPath $source -Destination $dest
  }
}

Write-Output "Tools ready:"
Write-Output "  Tesseract: $tesseract"
Write-Output "  Tessdata:  $Tessdata"
Write-Output "  Poppler:   $PopplerDir"
