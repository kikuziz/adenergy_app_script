# Nuskaitome Script ID iš .clasp.json failo
$claspConfig = Get-Content .clasp.json -Raw | ConvertFrom-Json
$scriptId = $claspConfig.scriptId

# Naudojame /dev nuorodą (Test Deployment). Ji nekuria naujų versijų, todėl nepasiekiamas 200 versijų limitas.
# Tai leidžia greičiau testuoti pakeitimus.
$execUrl = "https://script.google.com/macros/s/$scriptId/dev"

# Atidarome nuorodą Google Chrome naršyklėje
Write-Host "Atidaroma nuoroda: $execUrl"
Start-Process chrome $execUrl