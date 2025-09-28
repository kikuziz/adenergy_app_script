# deploy.ps1

# Išvalome ekraną, kad būtų aiškiau
#Clear-Host

clasp push

$deployOutput = clasp deploy
$deployments = clasp deployments
write-Host $deployOutput

# Ieškome eilutės, kuri neturi "@HEAD" ir paimame pirmą tokią
#$latestDeploymentLine = $deployments | Where-Object { $_ -notmatch "@HEAD" } | Select-Object -First 1




# Iš eilutės ištraukiame tik ID
$deploymentId = ($deployOutput -split ' ')[1]

if (-not $deploymentId) {
    Write-Host "❌ Nepavyko išgauti įdiegimo ID." -ForegroundColor Red
    exit 1
}

# 4. Sukonstruojame ir parodome galutinę nuorodą
$execUrl = "https://script.google.com/macros/s/$deploymentId/exec"



# Atidarome nuorodą Google Chrome naršyklėje
Write-Host "Atidaroma nuoroda: $execUrl"
Start-Process chrome $execUrl


