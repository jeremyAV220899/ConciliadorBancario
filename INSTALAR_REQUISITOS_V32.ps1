$ErrorActionPreference = 'Stop'
Write-Host 'CONCILIADOR V32 - REQUISITOS OCR' -ForegroundColor Cyan
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { Write-Host 'WinGet no disponible. Instale App Installer desde Microsoft Store.' -ForegroundColor Red; Read-Host 'Pulse Enter'; exit 1 }
$packages = @(@{ exe='tesseract.exe'; id='UB-Mannheim.TesseractOCR' }, @{exe='pdftoppm.exe'; id='oschwartz10612.Poppler'})
foreach ($package in $packages) {
  $known = Get-Command $package.exe -ErrorAction SilentlyContinue
  if ($known) { Write-Host "Disponible: $($package.exe)" -ForegroundColor Green; continue }
  Write-Host "Comprobando/instalando $($package.id)..." -ForegroundColor Yellow
  & winget install --id $package.id --exact --accept-package-agreements --accept-source-agreements
  if ($LASTEXITCODE -ne 0) { Write-Warning "WinGet devolvió código $LASTEXITCODE. Revise si ya estaba instalado." }
}
Write-Host 'Instalación finalizada. Regrese al conciliador y pulse Verificar nuevamente.' -ForegroundColor Green
Read-Host 'Pulse Enter para cerrar'
