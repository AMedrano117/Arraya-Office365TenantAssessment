# get all enterprise apps

# Connect to Graph with needed scopes
Connect-MgGraph -Scopes "Application.Read.All", "AuditLog.Read.All", "Directory.Read.All"

# Get all Enterprise Applications (Service Principals)
$apps = Get-MgServicePrincipal -All

# Prepare a list to hold the output
$results = @()
$progress = 0
$total = $apps.Count
$startTime = Get-Date
foreach ($app in $apps) {
    $protocol = $null
    $lastSignIn = $null
    $progress++
    Write-ProgressHelper -Activity "Getting SSO protocol and last sign-in" -CurrentOperation "Processing $($app.DisplayName)" -TotalCount $total -ProgressCounter $progress -StartTime $startTime

    # Try to get the SSO protocol
    try {
        $ssoDetails = Get-MgServicePrincipal -ServicePrincipalId $app.Id | Select-Object -ExpandProperty PreferredSingleSignOnMode
        $protocol = if ($ssoDetails) { $ssoDetails } else { "Not set" }
    } catch {
        $protocol = "Error retrieving"
    }

    # Get latest sign-in from signIn logs (limited retention)
    $signin = Get-MgAuditLogSignIn -Filter "AppId eq '$($app.AppId)'" -Top 1 -Sort "createdDateTime desc" -ErrorAction SilentlyContinue
    $lastSignIn = if ($signin) { $signin[0].CreatedDateTime } else { "No sign-ins found" }

    # Build output object
    $results += [PSCustomObject]@{
        DisplayName = $app.DisplayName
        ResourceDisplayName = $app.ResourceDisplayName
        AppId       = $app.AppId
        Protocol    = $protocol
        LastSignIn  = $lastSignIn
        LastUserSignIn = $signin.UserDisplayName
        LastUserUPN = $signin.UserPrincipalName
        ConditionalAccessStatus = $signin.ConditionalAccessStatus
        ClientAppUsed = $signin.ClientAppUsed
    }

    Write-Host "Processed $($app.DisplayName) using $($protocol). Last Sign In: $($lastSignIn)" -ForegroundColor Green
}

# Display results in table
$results | Sort-Object LastSignIn -Descending | Format-Table -AutoSize
