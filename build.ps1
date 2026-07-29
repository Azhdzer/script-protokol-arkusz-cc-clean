# build.ps1 — jednokomendowa budowa ProtokolCC.exe
# Uruchom:  powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$py = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) { throw "Brak .venv — utworz virtualenv i zainstaluj requirements." }

Write-Host "==> Instaluje/aktualizuje PyInstaller..." -ForegroundColor Cyan
& $py -m pip install --upgrade pip pyinstaller | Out-Host

Write-Host "==> Czyszcze poprzednia budowe..." -ForegroundColor Cyan
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

Write-Host "==> Buduje ProtokolCC.exe..." -ForegroundColor Cyan
& $py -m PyInstaller ProtokolCC.spec --noconfirm --clean | Out-Host

$exe = Join-Path $PSScriptRoot "dist\ProtokolCC.exe"
if (Test-Path $exe) {
    $mb = [math]::Round((Get-Item $exe).Length / 1MB, 1)
    Write-Host "`n==> GOTOWE: $exe  ($mb MB)" -ForegroundColor Green
} else {
    throw "Budowa nie powiodla sie — brak $exe"
}
