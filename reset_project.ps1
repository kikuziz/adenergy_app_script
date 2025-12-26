# reset_project.ps1
# DĖMESIO: Šis skriptas sukurs NAUJĄ Google Apps Script projektą.
# Tai padeda apeiti 200 versijų limitą, pradedant istoriją nuo nulio.
# Pasikeis Script ID ir Deployment URL.

# 1. Išsaugome seną .clasp.json kaip atsarginę kopiją
if (Test-Path .clasp.json) {
    Rename-Item .clasp.json .clasp.json.bak -Force
    Write-Host "Senas .clasp.json pervadintas į .clasp.json.bak" -ForegroundColor Yellow
}

# 2. Sukuriame naują projektą (Standalone tipas, kaip ir buvo naudojama)
Write-Host "Kuriamas naujas Apps Script projektas..." -ForegroundColor Cyan
clasp create --type standalone --title "Ad Energy Program (v2)"

# 3. Įkeliame failus į naują projektą
Write-Host "Įkeliami failai į naują projektą..." -ForegroundColor Cyan
clasp push --force

Write-Host "✅ Projektas sėkmingai perkurtas!" -ForegroundColor Green
Write-Host "Dabar galite naudoti 'deploy2.ps1' darbui arba 'deploy.ps1' naujai versijai sukurti." -ForegroundColor Green