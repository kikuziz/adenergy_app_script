# undeploy-all.ps1
# This script removes all versioned deployments for the current clasp project.

Write-Host "Fetching all deployments..." -ForegroundColor Cyan

# Get all deployments in a reliable JSON format
# Get all deployments as an array of strings (for older clasp versions)
$deploymentsOutput = clasp deployments
$deploymentLines = $deploymentsOutput | Where-Object { $_.Trim().StartsWith('-') }

# Check if there are any deployments to remove
if ($null -eq $deploymentLines -or $deploymentLines.Count -eq 0) {
    Write-Host "No deployments found to remove."
    exit
}

Write-Host "Found $($deploymentLines.Count) deployments. Starting removal..."

# Loop through each deployment and undeploy it
foreach ($line in $deploymentLines) {
    # The @HEAD deployment is a special pointer to the latest code and cannot be undeployed.
    if ($line -match "@HEAD") {
        Write-Host "Skipping special @HEAD deployment." -ForegroundColor Gray
        continue
    }
    
    # Extract the deployment ID, which is the second token on the line
    $deploymentId = ($line.Trim() -split ' ')[1]
    
    Write-Host "Removing deployment (ID: $deploymentId)..." -ForegroundColor Yellow
    clasp undeploy $deploymentId
}

Write-Host "✅ All versioned deployments have been successfully removed." -ForegroundColor Green
