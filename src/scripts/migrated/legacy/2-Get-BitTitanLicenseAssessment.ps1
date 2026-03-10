<#
.SYNOPSIS
    BitTitan License Assessment for Exchange Online Mailboxes

.DESCRIPTION
    This script gathers Exchange Online mailbox data and calculates the required 
    BitTitan MigrationWiz licenses needed for migration.
    
    Supports both interactive authentication and certificate-based authentication.
    
    License Types:
    - MigrationWiz-Mailbox: For mailboxes up to 50 GB
    - MigrationWiz-Mailbox (x2): For mailboxes 51-100 GB (requires 2 licenses)
    - User Migration Bundle: For mailboxes with Archive enabled (any size)
    
    Outputs an HTML report with executive summary and detailed breakdown.

.PARAMETER ClientId
    The Application (Client) ID from the app registration

.PARAMETER TenantId
    The Tenant ID for the client's Microsoft 365 tenant

.PARAMETER CertificateThumbprint
    The thumbprint of the certificate used for app-based authentication

.PARAMETER OutputPath
    Optional path for the HTML report. Defaults to Desktop.

.EXAMPLE
    # Interactive authentication (prompts for login)
    .\2-Get-BitTitanLicenseAssessment.ps1

.EXAMPLE
    # Certificate-based authentication (no prompts)
    .\2-Get-BitTitanLicenseAssessment.ps1 -ClientId "xxx" -TenantId "xxx" -CertificateThumbprint "xxx"

.NOTES
    Author: Migration Assessment Tool
    Version: 1.3 (Enhanced & Optimized)
    Date: 2025-01-08
    Requires: ExchangeOnlineManagement PowerShell module
    
    ENHANCEMENTS IN v1.3:
    - Fixed critical syntax errors and variable references
    - Optimized parallel processing for PowerShell 7
    - Adaptive batch sizing based on tenant size and performance
    - Standardized error messaging throughout script
    - Improved code organization with clear functional sections
    - Added comprehensive function documentation
    - Consolidated redundant lookup table creation
    - Inline license calculation logic for clarity
    - Consistent type labels: ActiveMailbox, ActiveArchive, InactiveMailbox, InactiveArchive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$CertificateThumbprint,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [switch]$GraphVerbose
)

# ========================================
#region SECTION 1: CORE UTILITY FUNCTIONS

function Write-ProgressHelper {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][int]$Total,
        [int]$Index,
        [string]$Activity = 'Processing',
        [string]$Operation,
        [int]$Id = 1,
        [int]$ParentId,
        [switch]$Completed
    )

    if ($ProgressPreference -eq 'SilentlyContinue') { $ProgressPreference = 'Continue' }

    # Always ensure required progress dictionaries exist before any use
    if (-not $script:ProgressStartTimes) { $script:ProgressStartTimes = @{} }
    if (-not $script:ProgressIndices)    { $script:ProgressIndices    = @{} }
    if (-not $script:ProgressTotals)     { $script:ProgressTotals     = @{} }

    if (-not $script:ProgressStartTimes.ContainsKey($Id)) { $script:ProgressStartTimes[$Id] = Get-Date }
    if (-not $script:ProgressIndices.ContainsKey($Id))    { $script:ProgressIndices[$Id]    = 0 }
    if (-not $script:ProgressTotals.ContainsKey($Id))     { $script:ProgressTotals[$Id]     = $Total }

    # If Total changed for this Id, reset timer/index so batches don't bleed together
    if ($script:ProgressTotals[$Id] -ne $Total) {
        $script:ProgressStartTimes[$Id] = Get-Date
        $script:ProgressIndices[$Id]    = 0
        $script:ProgressTotals[$Id]     = $Total
    }

    # Auto-increment when Index isn't provided
    if (-not $PSBoundParameters.ContainsKey('Index')) {
        $script:ProgressIndices[$Id]++
        $Index = $script:ProgressIndices[$Id]
    } else {
        $script:ProgressIndices[$Id] = $Index
    }

    $startTime = $script:ProgressStartTimes[$Id]
    $elapsed   = (Get-Date) - $startTime

    # Clamp index within [0, Total] so percent never exceeds 100
    if ($Total -gt 0) {
        if     ($Index -lt 0)      { $Index = 0 }
        elseif ($Index -gt $Total) { $Index = $Total }
    }

    $percent = if ($Total -gt 0) {
        $raw = (($Index / $Total) * 100)
        [math]::Round([math]::Min(100,[math]::Max(0,$raw)), 2)
    } else { $null }

    $etaSec  = if ($Index -gt 0 -and $Total -ge $Index) {
        $rate = $elapsed.TotalSeconds / $Index
        [int][math]::Max(0, [math]::Round($rate * ($Total - $Index)))
    } else { $null }

    $status = if ($Total -gt 0) { "[${Index} / ${Total}] $($elapsed.ToString('hh\:mm\:ss')) elapsed" }
              else              { "$($elapsed.ToString('hh\:mm\:ss')) elapsed" }

    $splat = @{ Activity=$Activity; Status=$status; Id=$Id }
    if ($percent -ne $null)   { $splat.PercentComplete  = $percent }
    if ($etaSec -ne $null)    { $splat.SecondsRemaining = $etaSec }
    if ($Operation)           { $splat.CurrentOperation = $Operation }
    if ($PSBoundParameters.ContainsKey('ParentId')) { $splat.ParentId = $ParentId }
    if ($Completed.IsPresent) { $splat.Completed = $true }

    Write-Progress @splat

    if ($Completed.IsPresent) {
        $script:ProgressStartTimes.Remove($Id) | Out-Null
        $script:ProgressIndices.Remove($Id)    | Out-Null
        $script:ProgressTotals.Remove($Id)     | Out-Null
    }
}

<#
.SYNOPSIS
    Writes standardized error, warning, or info messages

.DESCRIPTION
    Provides consistent formatting for all script messages with emoji indicators
    and appropriate colors based on severity

.PARAMETER Message
    The main message to display

.PARAMETER Details
    Optional additional details to show below the message

.PARAMETER Severity
    Message severity: Error, Warning, or Info
#>
function Write-StandardMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Details,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('Error', 'Warning', 'Info', 'Success')]
        [string]$Severity = 'Info'
    )
    
    $emoji = switch ($Severity) {
        'Error'   { '❌' }
        'Warning' { '⚠️ ' }
        'Info'    { 'ℹ️ ' }
        'Success' { '✅' }
    }
    
    $color = switch ($Severity) {
        'Error'   { 'Red' }
        'Warning' { 'Yellow' }
        'Info'    { 'Cyan' }
        'Success' { 'Green' }
    }
    
    Write-Host "$emoji $Message" -ForegroundColor $color
    if ($Details) {
        Write-Host "   $Details" -ForegroundColor Gray
    }
}

function Get-PowerShellVersion {
    return $PSVersionTable.PSVersion.Major
}

function Write-GraphVerbose {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )
    if ($GraphVerbose.IsPresent) {
        Write-Verbose $Message
    }
}

<# Get-SafeThrottleLimit Function
.SYNOPSIS
    Calculates optimal throttle limit based on processor count

.DESCRIPTION
    Determines the number of parallel threads to use based on logical CPU cores,
    constrained between minimum and maximum values for stability

.PARAMETER Default
    Default throttle limit if calculation fails (default: 10)

.PARAMETER Min
    Minimum throttle limit (default: 4)

.PARAMETER Max
    Maximum throttle limit (default: 20)

.OUTPUTS
    Int32 - Optimal throttle limit for parallel processing

.EXAMPLE
    Get-SafeThrottleLimit -Default 10 -Min 4 -Max 20
#>
function Get-SafeThrottleLimit {
    param(
        [int]$Default = 10,
        [int]$Min = 4,
        [int]$Max = 20
    )
    try {
        $logicalCores = [Environment]::ProcessorCount
        $throttle = $logicalCores
        if ($throttle -lt $Min) { $throttle = $Min }
        if ($throttle -gt $Max) { $throttle = $Max }
        return $throttle
    }
    catch {
        return $Default
    }
}

<# Get-OptimalBatchSize Function
.SYNOPSIS
    Calculates optimal batch size based on tenant size and performance testing

.DESCRIPTION
    Determines the best batch size for processing mailboxes based on total count
    Uses tiered approach: larger batches for larger tenants

.PARAMETER TotalCount
    Total number of items to process

.PARAMETER MinBatchSize
    Minimum batch size (default: 100)

.PARAMETER MaxBatchSize
    Maximum batch size (default: 2000)

.EXAMPLE
    Get-OptimalBatchSize -TotalCount 50000
#>
function Get-OptimalBatchSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [int]$TotalCount,
        
        [Parameter(Mandatory=$false)]
        [int]$MinBatchSize = 100,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxBatchSize = 2000
    )
    
    # Tiered batch sizing based on tenant size
    $batchSize = if ($TotalCount -lt 500) {
        100   # Small tenant: smaller batches for faster feedback
    }
    elseif ($TotalCount -lt 2000) {
        250   # Medium tenant: balanced approach
    }
    elseif ($TotalCount -lt 10000) {
        500   # Large tenant: larger batches for efficiency
    }
    elseif ($TotalCount -lt 50000) {
        1000  # Very large tenant: optimize for throughput
    }
    else {
        2000  # Enterprise tenant: maximum batch size
    }
    
    # Ensure within bounds
    $batchSize = [math]::Max($MinBatchSize, [math]::Min($MaxBatchSize, $batchSize))
    
    Write-Verbose "Calculated batch size: $batchSize for $TotalCount items"
    return $batchSize
}

#endregion
# ========================================

#region SECTION 2: MODULE & CONNECTION FUNCTIONS

function Test-RequiredModules {
    $requiredModules = @('ExchangeOnlineManagement')
    $missingModules = @()
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            $missingModules += $module
        }
    }
    
    if ($missingModules.Count -gt 0) {
        Write-Host "Missing required modules:" -ForegroundColor Red
        $missingModules | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
        Write-Host "`nInstall with: Install-Module -Name $($missingModules -join ', ') -Scope CurrentUser -Force" -ForegroundColor Cyan
        return $false
    }
    
    return $true
}

function Connect-ToExchangeOnline {
    param(
        [string]$ClientId,
        [string]$TenantId,
        [string]$CertificateThumbprint
    )
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Connecting to Exchange Online" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    try {
        # Determine authentication method
        $usingCertificate = $ClientId -and $TenantId -and $CertificateThumbprint
        $usingAppBased = $ClientId -and $TenantId -and -not $CertificateThumbprint
        
        if ($usingCertificate) {
            Write-Host "`n🔐 Using Certificate-Based Authentication" -ForegroundColor Yellow
            Write-Host "  Client ID: $ClientId" -ForegroundColor Gray
            Write-Host "  Tenant ID: $TenantId" -ForegroundColor Gray
            #Write-Host "  Certificate: $CertificateThumbprint" -ForegroundColor Gray
            
            # Convert Tenant ID (GUID) to domain name using certificate auth for Graph
            Write-Host "`nConnecting to Microsoft Graph with certificate..." -ForegroundColor Cyan
            
            try {
                # Connect to Graph using the same certificate
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
                
                $org = Get-MgOrganization -ErrorAction Stop
                $tenantDomain = ($org.VerifiedDomains | Where-Object { $_.IsInitial -eq $true }).Name
                
                if (-not $tenantDomain) {
                    # Fallback to any verified domain
                    $tenantDomain = $org.VerifiedDomains[0].Name
                }
                
                Write-Host "  ✅ Connected to Graph via certificate" -ForegroundColor Green
                Write-Host "  Tenant Domain: $tenantDomain" -ForegroundColor Gray
                
                # Keep Graph connection open for potential future use
                # Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }
            catch {
                Write-StandardMessage -Message "Failed to connect to Graph with certificate" `
                    -Details $_.Exception.Message `
                    -Severity 'Error'
                
                Write-Host "`n⚠️  Could not auto-detect tenant domain" -ForegroundColor Yellow
                $tenantDomain = Read-Host "Please enter your tenant domain (e.g., contoso.onmicrosoft.com)"
                
                if ([string]::IsNullOrWhiteSpace($tenantDomain)) {
                    throw "Tenant domain is required for certificate-based authentication"
                }
            }
            
            Write-Host "`nConnecting to Exchange Online with certificate..." -ForegroundColor Cyan
            
            try {
                # Certificate-based authentication requires domain, not GUID
                Connect-ExchangeOnline `
                    -AppId $ClientId `
                    -Organization $tenantDomain `
                    -CertificateThumbprint $CertificateThumbprint `
                    -ShowBanner:$false `
                    -ErrorAction Stop
                
                Write-Host "✅ Connected to Exchange Online via certificate!" -ForegroundColor Green
            }
            catch {
                Write-Host "`n❌ Certificate-based authentication to Exchange Online failed" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
                
                Write-Host "`n💡 Common Issues:" -ForegroundColor Yellow
                Write-Host "  1. Certificate not uploaded to app registration in Azure Portal" -ForegroundColor White
                Write-Host "  2. Certificate not in your store: Get-ChildItem Cert:\CurrentUser\My" -ForegroundColor White
                Write-Host "  3. Admin consent not granted for the app" -ForegroundColor White
                Write-Host "  4. App missing Exchange.ManageAsApp permission" -ForegroundColor White
                
                # Ask if user wants to fall back to interactive
                Write-Host ""
                $fallback = Read-Host "Would you like to try interactive authentication instead? (Y/N)"
                
                if ($fallback -eq 'Y' -or $fallback -eq 'y') {
                    Write-Host "`nFalling back to interactive authentication..." -ForegroundColor Cyan
                    
                    # Disconnect certificate-based Graph connection
                    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                    
                    # Connect to Exchange Online interactively
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                }
                else {
                    throw "Certificate-based authentication failed and user declined fallback"
                }
            }
        }
        elseif ($usingAppBased) {
            Write-Host "`n⚠️  App ID and Tenant ID provided but no certificate" -ForegroundColor Yellow
            Write-Host "Certificate-based auth requires all three parameters.`n" -ForegroundColor Yellow
            
            Write-Host "Falling back to interactive authentication..." -ForegroundColor Cyan
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        else {
            Write-Host "`n👤 Using Interactive Authentication" -ForegroundColor Yellow
            Write-Host "You will be prompted to sign in...`n" -ForegroundColor White
            
            # Interactive authentication
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        
        # Verify connection by getting organization config
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            
            Write-Host "`n✅ Successfully connected to Exchange Online!" -ForegroundColor Green
            Write-Host "   Organization: $($orgConfig.DisplayName)" -ForegroundColor White
            Write-Host "   Tenant: $($orgConfig.Identity)" -ForegroundColor Gray
            
            return $orgConfig
            
        }
        catch {
            Write-Host "`n❌ Connected but failed to retrieve organization configuration" -ForegroundColor Red
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
            throw "Failed to verify Exchange Online connection"
        }
        
    }
    catch {
        Write-Host "`n❌ Failed to connect to Exchange Online" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        
        # This will cause the script to terminate in the main try-catch
        throw
    }
}

#endregion
# ========================================

function Get-AllMailboxData {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Gathering Mailbox Data" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    
    $mailboxData = @{
        AllMailboxes = @() # All Mailboxes Including Inactive - An Array
        ActiveMailboxes = @() # Active Mailboxes - An Array
        ActiveArchiveMailboxes = @() # Active Archive Mailboxes - An Array
        InactiveMailboxes = @() # Inactive Mailboxes - An Array
        InactiveArchiveMailboxes = @() # Inactive Archive Mailboxes - An Array
        AllMailboxStats = @{} # Mailbox Stats - A Hash Table
        AllArchiveStats = @{} # Active Archive Stats - A Hash Tabl
    }
    
    # Get all mailboxes
    try {
        #$throttleLimit = Get-SafeThrottleLimit
        #Write-Host "  Using ThrottleLimit: $throttleLimit (based on processor count)" -ForegroundColor Gray

        # Test if there are any inactive mailboxes
        $inactiveMBXTest = Get-EXOMailbox -InactiveMailboxOnly -ResultSize 1 -ErrorAction SilentlyContinue
        $MailboxesStartTime = Get-Date

        if ($inactiveMBXTest) {
            Write-Host "`nGetting all mailboxes (including inactive)..." -ForegroundColor Cyan
            Write-Host "This will take a while..." -ForegroundColor Gray
            $mailboxData.AllMailboxes = Get-EXOMailbox -IncludeInactiveMailbox -ResultSize Unlimited `
                    -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientTypeDetails,ExchangeGuid, `
                        ArchiveStatus,ArchiveGuid,WhenMailboxCreated,IsInactiveMailbox,WhenSoftDeleted,LitigationHoldEnabled -ErrorAction Stop |
                    Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' }
        } else {
            Write-Host "`nGetting all mailboxes (excluding inactive)..." -ForegroundColor Cyan
            Write-Host "This will take a while..." -ForegroundColor Gray
            $mailboxData.AllMailboxes = Get-EXOMailbox -ResultSize Unlimited `
                    -Properties DisplayName,UserPrincipalName,PrimarySmtpAddress,RecipientTypeDetails,ExchangeGuid, `
                        ArchiveStatus,ArchiveGuid,WhenMailboxCreated,IsInactiveMailbox,WhenSoftDeleted,LitigationHoldEnabled -ErrorAction Stop |
                    Where-Object { $_.RecipientTypeDetails -ne 'DiscoveryMailbox' }
        }
    }
    catch {
        Write-Host "  ⚠️  Error in bulk All Mailbox statistics retrieval: $($_.Exception.Message)" -ForegroundColor Yellow
        throw
    }

    # Check and Build Mailbox Tables
    if ($mailboxData.AllMailboxes.Count -eq 0) {
        Write-Host "`n⚠️  Warning: No mailboxes found!" -ForegroundColor Yellow
        throw "No mailboxes found - stopping assessment"
    }
    # Get active mailboxes
    $elapsed = ((Get-Date) - $MailboxesStartTime).TotalMinutes
    Write-Host "  ✅ Found $($mailboxData.AllMailboxes.Count) mailboxes ($([math]::Round($elapsed,1))m)" -ForegroundColor Green
    Write-Host "  Active Mailboxes: $(($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false }).Count)" -ForegroundColor Gray
    Write-Host "  Inactive Mailboxes: $(($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $true }).Count)" -ForegroundColor Gray
    Write-Host "  Active Archive Mailboxes: $(($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false -and $_.ArchiveStatus -eq 'Active' }).Count)" -ForegroundColor Gray
    Write-Host "  Inactive Archive Mailboxes: $(($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $true -and $_.ArchiveStatus -eq 'Active' }).Count)" -ForegroundColor Gray
    
    # Get mailbox statistics - Try Graph API first (much faster), fall back to EXO cmdlets
    Write-Host "`nGathering Active Mailbox Statistics..." -ForegroundColor Cyan
    $statsStartTime = Get-Date
    $useGraphSuccess = $false
    $graphDurationSec = $null
    $exoMailboxStatsSeconds = $null
    $exoArchiveStatsSeconds = $null
    $activeMailboxes = $mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false }
    $activeMailboxCount = $activeMailboxes.Count
    try {
        # Method 1: Try Graph API Mailbox Usage Report (FAST - bulk data)
        Write-GraphVerbose "  Attempting Graph API method (fastest)..."

        # Check if we're connected to Graph
        $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        
        if ($graphContext) {
            #Write-Host "  Graph connection detected, fetching mailbox report..." -ForegroundColor Gray
            
            # Get mailbox usage report from Graph - save to temp file
            $mailboxUsageUri = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')"
            $tempCsvFile = Join-Path $env:TEMP "MailboxUsageReport-$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            
            try {
                Write-GraphVerbose "Graph API: Starting CSV download from $mailboxUsageUri"
                
                # Show progress during download
                Write-Progress -Activity "Gathering Mailbox Statistics" -Status "Downloading report from Microsoft Graph..." -PercentComplete 25
                $ProgressPreference = 'SilentlyContinue'
                
                # Download the CSV report to a temp file
                Invoke-MgGraphRequest -Uri $mailboxUsageUri -Method GET -OutputFilePath $tempCsvFile -ErrorAction Stop | Out-Null
                
                # ✅ CRITICAL FIX: Clear progress bar immediately after download
                Write-Progress -Activity "Gathering Mailbox Statistics" -Completed
                $ProgressPreference = 'Continue'
                
                Write-GraphVerbose "Graph API: CSV downloaded to $tempCsvFile"
                
                # Verify file was created
                if (-not (Test-Path $tempCsvFile)) {
                    throw "CSV file was not created"
                }
                
                $fileSize = (Get-Item $tempCsvFile).Length
                Write-GraphVerbose "Graph API: File size is $fileSize bytes"
                
                # Read and parse the CSV
                Write-GraphVerbose "Graph API: Importing CSV..."
                $graphReportData = Import-Csv -Path $tempCsvFile -ErrorAction Stop
                Write-GraphVerbose "Graph API: CSV imported successfully, $($graphReportData.Count) rows"
                
                # Clean up temp file
                Remove-Item -Path $tempCsvFile -Force -ErrorAction SilentlyContinue
                
                # Validate we got data
                if (-not $graphReportData -or $graphReportData.Count -eq 0) {
                    Write-Host "  ⚠️  Graph report returned no data" -ForegroundColor Yellow
                    throw "Graph report is empty"
                }
                
                Write-Host "  ✅ Retrieved $($graphReportData.Count) mailboxes from Graph" -ForegroundColor Green
                
                # Create hash table from Graph data for quick lookup by UPN
                Write-GraphVerbose "Graph API: Building hash table by UPN..."
                $graphMailboxHash = @{}
                $rowNumber = 0
                foreach ($item in $graphReportData) {
                    $rowNumber++
                    
                    try {
                        $upn = $item.'User Principal Name'
                        
                        if ([string]::IsNullOrWhiteSpace($upn)) {
                            Write-GraphVerbose "Graph API: Row $rowNumber has null/empty UPN, skipping"
                            continue
                        }
                        
                        $graphMailboxHash[$upn] = $item
                    }
                    catch {
                        Write-Warning "Graph API: Error processing row $rowNumber - $($_.Exception.Message)"
                        continue
                    }
                }
                
                Write-GraphVerbose "Graph API: Hash table built with $($graphMailboxHash.Count) entries"

                # Map Graph data to our mailboxes: Match by UPN, store by GUID
                Write-GraphVerbose "Graph API: Starting mailbox mapping..."
                $matchedCount = 0
                $mailboxNumber = 0
                
                foreach ($mailbox in $activeMailboxes) {
                    $mailboxNumber++
                    
                    try {
                        # Match by UPN from Graph report
                        if (-not $graphMailboxHash.ContainsKey($mailbox.UserPrincipalName)) {
                            Write-GraphVerbose "Graph API: Mailbox $mailboxNumber ($($mailbox.UserPrincipalName)) - No Graph data found"
                            continue
                        }
                        
                        $graphData = $graphMailboxHash[$mailbox.UserPrincipalName]
                        Write-GraphVerbose "Graph API: Mailbox $mailboxNumber ($($mailbox.UserPrincipalName)) - Processing..."
                        
                        # Safely extract storage bytes
                        $storageBytes = 0
                        try {
                            $storageBytesRaw = $graphData.'Storage Used (Byte)'
                            Write-GraphVerbose "  Storage raw value: '$storageBytesRaw'"
                            
                            if (-not [string]::IsNullOrWhiteSpace($storageBytesRaw)) {
                                $storageBytes = [double]$storageBytesRaw
                                Write-GraphVerbose "  Storage parsed: $storageBytes bytes"
                            }
                        }
                        catch {
                            Write-Warning "  Could not parse storage: $($_.Exception.Message)"
                        }
                        
                        # Safely extract deleted item size
                        $deletedBytes = 0
                        try {
                            $deletedBytesRaw = $graphData.'Deleted Item Size (Byte)'
                            Write-GraphVerbose "  Deleted raw value: '$deletedBytesRaw'"
                            
                            if (-not [string]::IsNullOrWhiteSpace($deletedBytesRaw)) {
                                $deletedBytes = [double]$deletedBytesRaw
                                Write-GraphVerbose "  Deleted parsed: $deletedBytes bytes"
                            }
                        }
                        catch {
                            Write-Warning "  Could not parse deleted size: $($_.Exception.Message)"
                        }
                        
                        # Get item count
                        $itemCount = "0"
                        try {
                            $itemCountRaw = $graphData.'Item Count'
                            Write-GraphVerbose "  Item count raw: '$itemCountRaw'"
                            
                            if (-not [string]::IsNullOrWhiteSpace($itemCountRaw)) {
                                $itemCount = $itemCountRaw
                            }
                        }
                        catch {
                            Write-Warning "  Could not get item count: $($_.Exception.Message)"
                        }
                        
                        # Get display name
                        $displayName = $mailbox.DisplayName
                        try {
                            $displayNameRaw = $graphData.'Display Name'
                            if (-not [string]::IsNullOrWhiteSpace($displayNameRaw)) {
                                $displayName = $displayNameRaw
                            }
                        }
                        catch {
                            Write-Warning "  Could not get display name: $($_.Exception.Message)"
                        }
                        
                        # Validate critical mailbox properties
                        if (-not $mailbox.ExchangeGuid -and -not $mailbox.Guid) {
                            Write-Warning "  Mailbox has null ExchangeGuid, skipping: $($mailbox.UserPrincipalName)"
                            continue
                        }

                        $guidKey = if ($mailbox.ExchangeGuid) { $mailbox.ExchangeGuid.ToString() } else { $mailbox.Guid.ToString() }
                        Write-GraphVerbose "  GUID key: $guidKey"
                        
                        # Get mailbox type safely
                        $mailboxType = "Unknown"
                        if ($mailbox.RecipientTypeDetails) {
                            $mailboxType = $mailbox.RecipientTypeDetails
                        }
                        Write-GraphVerbose "  Mailbox type: $mailboxType"
                        
                        # Create mailbox statistics object from Graph data
                        Write-GraphVerbose "  Creating stats object..."
                        $stats = [PSCustomObject]@{
                            DisplayName = $displayName
                            TotalItemSize = "$([math]::Round($storageBytes / 1GB, 4)) GB ($storageBytes bytes)"
                            ItemCount = $itemCount
                            TotalDeletedItemSize = "$([math]::Round($deletedBytes / 1GB, 4)) GB ($deletedBytes bytes)"
                            MailboxType = $mailboxType
                            MailboxGuid = $mailbox.Guid
                        }
                        
                        # Store using GUID as the key for consistency
                        Write-GraphVerbose "  Storing with GUID key: $guidKey"
                        $mailboxData.AllMailboxStats[$guidKey] = $stats
                        $matchedCount++
                        
                        Write-GraphVerbose "  Success! Matched count: $matchedCount"
                    }
                    catch {
                        Write-Warning "Graph API: Error processing mailbox $mailboxNumber ($($mailbox.UserPrincipalName)) - $($_.Exception.Message)"
                        Write-GraphVerbose "  Stack trace: $($_.ScriptStackTrace)"
                    }
                }
                
                $elapsed = ((Get-Date) - $statsStartTime).TotalSeconds
                $graphDurationSec = $elapsed
                Write-Host "  ✅ Mapped $matchedCount mailboxes ($([math]::Round($elapsed,1))s)" -ForegroundColor Green
                
                if ($activeMailboxCount -gt 0 -and $matchedCount -ge ($activeMailboxCount * 0.8)) {
                    $useGraphSuccess = $true
                    #Write-Host "  Graph API method successful!" -ForegroundColor Green
                }
                else {
                    Write-Host "  ⚠️  Only matched $matchedCount of $activeMailboxCount, using EXO supplement..." -ForegroundColor Yellow
                }
            }
            catch {
                # Clean up temp file on error
                if (Test-Path $tempCsvFile) {
                    Remove-Item -Path $tempCsvFile -Force -ErrorAction SilentlyContinue
                }
                
                # ✅ CRITICAL FIX: Clear any stuck progress bars
                Write-Progress -Activity "Gathering Mailbox Statistics" -Completed
                
                Write-Host "  Graph API failed: $($_.Exception.Message)" -ForegroundColor Yellow
                
                # Check if it's a permissions issue
                if ($_.Exception.Message -like "*Insufficient privileges*" -or $_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*Forbidden*") {
                    Write-Host "  💡 Add Reports.Read.All permission for faster stats gathering" -ForegroundColor Gray
                }
                
                Write-Host "  Falling back to Exchange Online..." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "    No Graph connection - skipping Graph method" -ForegroundColor Gray
        }
    }
    catch {
        Write-Verbose "Graph API method failed: $($_.Exception.Message)"
        # ✅ CRITICAL FIX: Clear any stuck progress bars
        Write-Progress -Activity "Gathering Mailbox Statistics" -Status "Downloading report from Microsoft Graph..." -Completed
    }

    # Fetch Mailbox Statistics from EXO only for gaps (or full set if Graph failed)
    $mailboxesNeedingStats = if ($useGraphSuccess) {
        $mailboxData.AllMailboxes | Where-Object { -not $mailboxData.AllMailboxStats.ContainsKey($_.ExchangeGuid.ToString()) }
    } else {
        $mailboxData.AllMailboxes
    }

    if ($mailboxesNeedingStats.Count -gt 0) {
        if ($useGraphSuccess) {
            if ($inactiveMBXTest) {
                Write-Host "Graph report covered $($mailboxData.AllMailboxStats.Count) mailboxes; fetching EXO stats for $($mailboxesNeedingStats.Count) missing/inactive." -ForegroundColor Gray
            } else {
                Write-Host "Graph report covered $($mailboxData.AllMailboxStats.Count) mailboxes; fetching EXO stats for $($mailboxesNeedingStats.Count) missing." -ForegroundColor Gray
            }
        }
        Write-Host "Fetching Mailbox Statistics from Exchange Online..." -ForegroundColor Cyan -NoNewline
        Write-Host "This will take a while..." -ForegroundColor Gray

        $validMailboxStatsInput = $mailboxesNeedingStats | Where-Object {
            $_.UserPrincipalName -and $_.UserPrincipalName -match ".+@.+"
        }
        $invalidMailboxStatsInput = $mailboxesNeedingStats | Where-Object {
            -not $_.UserPrincipalName -or $_.UserPrincipalName -notmatch ".+@.+"
        }

        if ($invalidMailboxStatsInput.Count -gt 0) {
            Write-Host "  ⚠️  Skipping $($invalidMailboxStatsInput.Count) mailboxes without a valid UserPrincipalName; attempting by ExchangeGuid/PrimarySmtpAddress where possible." -ForegroundColor Yellow
        }


        $starttime2 = Get-Date
        $allMBXStats = @()
        foreach ($mbx in $validMailboxStatsInput) {
            try {
                if ($inactiveMBXTest) {
                    $stat = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -IncludeSoftDeletedRecipient
                } else {
                    $stat = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName
                }
                if ($stat) { $allMBXStats += $stat }
            }
            catch {
                $msg = $_.Exception.Message
                if ($msg -match "DatabaseNotFoundException|database with ID .* couldn't be found") {
                    Write-Host "  ⚠️  Skipping mailbox due to missing database: $($mbx.UserPrincipalName)" -ForegroundColor Yellow
                } else {
                    Write-Host "  ⚠️  Failed to get stats for $($mbx.UserPrincipalName): $msg" -ForegroundColor Yellow
                }
            }
        }

        $elapsed2   = ((Get-Date) - $starttime2)
        $exoMailboxStatsSeconds = $elapsed2.TotalSeconds
        Write-Host "Pipeline All MBX Stats Found $($allMBXStats.Count) mailboxes ($([math]::Round($elapsed2.TotalHours,3))h)" -ForegroundColor Green
        $allMBXStats | ForEach-Object {
            $mailboxData.AllMailboxStats[$_.MailboxGuid.ToString()] = $_
        }
        Write-Host "All Mailbox Statistics Fetched in $([math]::Round($elapsed2.TotalHours,3))h" -ForegroundColor Green

        foreach ($mbx in $invalidMailboxStatsInput) {
            $identity = if ($mbx.ExchangeGuid) { $mbx.ExchangeGuid.ToString() }
                        elseif ($mbx.Guid) { $mbx.Guid.ToString() }
                        elseif ($mbx.PrimarySmtpAddress -and $mbx.PrimarySmtpAddress -match ".+@.+") { $mbx.PrimarySmtpAddress }
                        elseif ($mbx.UserPrincipalName -and $mbx.UserPrincipalName -match ".+@.+") { $mbx.UserPrincipalName }
                        else { $null }

            if (-not $identity) {
                Write-Host "  ⚠️  Skipping mailbox without usable identity: $($mbx.DisplayName)" -ForegroundColor Yellow
                continue
            }

            try {
                if ($inactiveMBXTest) {
                    $stat = Get-EXOMailboxStatistics -Identity $identity -IncludeSoftDeletedRecipient
                } else {
                    $stat = Get-EXOMailboxStatistics -Identity $identity
                }
                if ($stat) {
                    $mailboxData.AllMailboxStats[$stat.MailboxGuid.ToString()] = $stat
                }
            }
            catch {
                Write-Host "  ⚠️  Failed to get stats for ${identity}: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Host "Skipping EXO mailbox statistics (Graph covered all active mailboxes)" -ForegroundColor Green
    }

    # Fetch All Archive Mailbox Statistics
    Write-Host "`nFetching All Archive Mailbox Statistics..." -ForegroundColor Cyan -NoNewline
    Write-Host "This will take a while..." -ForegroundColor Gray

    $archiveMailboxes = $mailboxData.AllMailboxes | Where-Object {
        $_.ArchiveStatus -eq 'Active' -and $_.UserPrincipalName -and $_.UserPrincipalName -match ".+@.+"
    }
    $skippedArchives = ($mailboxData.AllMailboxes | Where-Object { $_.ArchiveStatus -eq 'Active' }).Count - $archiveMailboxes.Count
    if ($skippedArchives -gt 0) {
        Write-Host "  ⚠️  Skipping $skippedArchives archive mailbox(es) without a valid UserPrincipalName." -ForegroundColor Yellow
    }

    $starttime3 = Get-Date
    $allArchiveMBXStats = @()
    foreach ($mbx in $archiveMailboxes) {
        try {
            if ($inactiveMBXTest) {
                $stat = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -Archive -IncludeSoftDeletedRecipient
            } else {
                $stat = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -Archive
            }
            if ($stat) { $allArchiveMBXStats += $stat }
        }
        catch {
            $msg = $_.Exception.Message
            if ($msg -match "DatabaseNotFoundException|database with ID .* couldn't be found") {
                Write-Host "  ⚠️  Skipping archive stats due to missing database: $($mbx.UserPrincipalName)" -ForegroundColor Yellow
            } else {
                Write-Host "  ⚠️  Failed to get archive stats for $($mbx.UserPrincipalName): $msg" -ForegroundColor Yellow
            }
        }
    }

    $elapsed3   = ((Get-Date) - $starttime3)
    $exoArchiveStatsSeconds = $elapsed3.TotalSeconds
    Write-Host "Pipeline All Archive MBX Stats Found $($allArchiveMBXStats.Count) mailboxes ($([math]::Round($elapsed3.TotalHours,3))h)" -ForegroundColor Green
    $allArchiveMBXStats | ForEach-Object {
        $mailboxData.AllArchiveStats[$_.MailboxGuid.ToString()] = $_
    }
    Write-Host "All Archive Mailbox Statistics Fetched in $([math]::Round($elapsed3.TotalHours,3))h" -ForegroundColor Green
    function Set-MailboxAndArchiveStats {
        param (
            [Parameter(Mandatory)]
            [object]$Mailbox,
            [Parameter(Mandatory)]
            [string]$MailboxGuidKey,
            [Parameter(Mandatory)]
            [string]$ArchiveGuidKey,
            [Parameter(Mandatory)]
            [hashtable]$MailboxStats,
            [Parameter(Mandatory)]
            [hashtable]$ArchiveStats,
            [switch]$IsInactive
        )

        $MailboxLabel = if ($IsInactive) { "inactive mailbox" } else { "mailbox" }

        # Choose element if input is array, else use as-is
        function Get-FirstIfArray { param($obj) if ($obj -is [System.Collections.IEnumerable] -and $obj -isnot [string]) { if ($obj.Count -gt 0) { return $obj[0] } } return $obj }

        # Retrieve the mailbox statistics object (if present)
        $MailboxStat = if ($MailboxStats.ContainsKey($MailboxGuidKey)) {
            Get-FirstIfArray $MailboxStats[$MailboxGuidKey]
        } elseif ($MailboxGuidKey) {
            if ($inactiveMBXTest) {
                Get-EXOMailboxStatistics -Identity $MailboxGuidKey -IncludeSoftDeletedRecipient
            } else {
                Get-EXOMailboxStatistics -Identity $MailboxGuidKey
            }
        } else {
            $null
        }
        # Skip if ArchiveGuidKey is all zeros
        if ($ArchiveGuidKey -eq "00000000-0000-0000-0000-000000000000" -or [string]::IsNullOrEmpty($ArchiveGuidKey)) {
            $ArchiveStat = $null
        } elseif ($ArchiveStats.ContainsKey($ArchiveGuidKey)) {
            $ArchiveStat = Get-FirstIfArray $ArchiveStats[$ArchiveGuidKey]
        } elseif ($ArchiveGuidKey) {
            $ArchiveStat = Get-EXOMailboxStatistics -Identity $ArchiveGuidKey -Archive -IncludeSoftDeletedRecipient
        } else {
            $ArchiveStat = $null
        }

        # Add property members for mailbox statistics (if do not exist already, Force ensures update)
        $Mailbox | Add-Member -MemberType NoteProperty -Name TotalItemSize        -Value $null -Force
        $Mailbox | Add-Member -MemberType NoteProperty -Name ItemCount            -Value $null -Force
        $Mailbox | Add-Member -MemberType NoteProperty -Name TotalDeletedItemSize -Value $null -Force
        $Mailbox | Add-Member -MemberType NoteProperty -Name DeletedItemCount     -Value $null -Force

        # Add property members for archive mailbox statistics
        $Mailbox | Add-Member -MemberType NoteProperty -Name ArchiveTotalItemSize        -Value $null -Force
        $Mailbox | Add-Member -MemberType NoteProperty -Name ArchiveItemCount            -Value $null -Force
        $Mailbox | Add-Member -MemberType NoteProperty -Name ArchiveTotalDeletedItemSize -Value $null -Force
        $Mailbox | Add-Member -MemberType NoteProperty -Name ArchiveDeletedItemCount     -Value $null -Force

        # Assign mailbox stats if present
        if ($MailboxStat) {
            $Mailbox.TotalItemSize        = $MailboxStat.TotalItemSize
            $Mailbox.ItemCount            = $MailboxStat.ItemCount
            $Mailbox.TotalDeletedItemSize = $MailboxStat.TotalDeletedItemSize
            $Mailbox.DeletedItemCount     = $MailboxStat.DeletedItemCount
        }

        # Assign archive stats if present
        if ($ArchiveStat) {
            $Mailbox.ArchiveTotalItemSize        = $ArchiveStat.TotalItemSize
            $Mailbox.ArchiveItemCount            = $ArchiveStat.ItemCount
            $Mailbox.ArchiveTotalDeletedItemSize = $ArchiveStat.TotalDeletedItemSize
            $Mailbox.ArchiveDeletedItemCount     = $ArchiveStat.DeletedItemCount
        }
    }

    # Map Mailbox Stats to Mailboxes - add stat properties to each matched mailbox object
    $mailboxData.AllMailboxes | ForEach-Object {
        Write-ProgressHelper -Total $mailboxData.AllMailboxes.Count -Activity "Setting Mailbox and Archive Stats" -Operation "Processing mailbox $($_.UserPrincipalName)"
        Set-MailboxAndArchiveStats  -Mailbox $_ `
                                    -MailboxGuidKey $_.ExchangeGuid.ToString() `
                                    -ArchiveGuidKey $_.ArchiveGuid.ToString() `
                                    -MailboxStats $mailboxData.AllMailboxStats `
                                    -ArchiveStats $mailboxData.AllArchiveStats `
    }
    Write-ProgressHelper -Activity "Setting Mailbox and Archive Stats" -Total $mailboxData.AllMailboxes.Count -Completed

    # Add to Mailbox Data
    $mailboxData.ActiveMailboxes = $mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false }
    $mailboxData.InactiveMailboxes = $mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $true }
    $mailboxData.ActiveArchiveMailboxes = $mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false -and $_.ArchiveStatus -eq 'Active' }
    $mailboxData.InactiveArchiveMailboxes = $mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $true -and $_.ArchiveStatus -eq 'Active' }

    # Print data collection summary
    Write-Host "`n📊 Data Collection Summary:" -ForegroundColor Cyan
    Write-Host "   Active Mailboxes: $((($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false})).Count)" -ForegroundColor White
    Write-Host "   Active Mailbox Statistics: $((($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false -and $null -ne $_.ItemCount })).Count)" -ForegroundColor White
    Write-Host "   Active Archive Statistics: $((($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $false -and $null -ne $_.ArchiveItemCount })).Count)" -ForegroundColor White
    Write-Host "   Inactive Mailboxes: $(($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $true}).Count)" -ForegroundColor White
    Write-Host "   Inactive Mailbox Statistics: $((($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $true -and $null -ne $_.ItemCount })).Count)" -ForegroundColor White
    Write-Host "   Inactive Archive Statistics: $((($mailboxData.AllMailboxes | Where-Object { $_.IsInactiveMailbox -eq $true -and $null -ne $_.ArchiveItemCount })).Count)" -ForegroundColor White
    
    Write-Host "`n⏱️  Timing Summary:" -ForegroundColor Cyan
    if ($graphDurationSec -ne $null) {
        Write-Host ("   Graph report mapping: {0}s" -f ([math]::Round($graphDurationSec, 1))) -ForegroundColor White
    } else {
        Write-Host "   Graph report mapping: Not used or failed" -ForegroundColor Yellow
    }
    if ($exoMailboxStatsSeconds -ne $null) {
        Write-Host ("   EXO mailbox stats: {0}s" -f ([math]::Round($exoMailboxStatsSeconds, 1))) -ForegroundColor White
    }
    if ($exoArchiveStatsSeconds -ne $null) {
        Write-Host ("   EXO archive stats: {0}s" -f ([math]::Round($exoArchiveStatsSeconds, 1))) -ForegroundColor White
    }
    return $mailboxData

}

#endregion
# ========================================

#region SECTION 4: DATA TRANSFORMATION FUNCTIONS

<# Format-DataSize Function
.SYNOPSIS
    Formats data size in GB or TB based on magnitude

.DESCRIPTION
    Converts GB values to TB format when size exceeds 1000 GB for better readability

.PARAMETER SizeInGB
    Size in gigabytes to format

.OUTPUTS
    String - Formatted size with unit (GB or TB)

.EXAMPLE
    Format-DataSize -SizeInGB 5247.5
    # Returns: "5.12 TB"

.EXAMPLE
    Format-DataSize -SizeInGB 145.8
    # Returns: "145.8 GB"
#>
function Format-DataSize {
    param([double]$SizeInGB)
    
    # If over 1000 GB (1 TB), show in TB
    if ($SizeInGB -ge 1000) {
        $sizeInTB = [math]::Round($SizeInGB / 1024, 2)
        return "$sizeInTB TB"
    }
    else {
        return "$([math]::Round($SizeInGB, 2)) GB"
    }
}

#endregion
# ========================================

#region SECTION 5: ANALYSIS & CALCULATION FUNCTIONS
# ========================================

function Get-BitTitanLicenseCalculation {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$MailboxData
    )

    # Data structure to aggregate mailbox data
    $analysis = @{
        ActiveMailboxes = @{
            Total = 0
            ByType = @{}
            WithArchive = 0
            Over50GB = 0
            Over100GB = 0
            ArchiveOver100GB = 0
            Over50GBWithDeleted = 0
            TotalMailboxDataGB = 0.0
            TotalDeletedItemsGB = 0.0
            TotalArchiveDataGB = 0.0
            TotalArchiveDeletedItemsGB = 0.0 
            LicenseCounts = @{
                MigrationWizMailbox = 0
                UserMigrationBundle = 0
                MailboxesOver50GB = 0
            }
        }
        InactiveMailboxes = @{
            Total = 0
            ByType = @{}
            WithArchive = 0
            Over50GB = 0
            Over100GB = 0
            ArchiveOver100GB = 0
            Over50GBWithDeleted = 0
            TotalMailboxDataGB = 0.0
            TotalDeletedItemsGB = 0.0
            TotalArchiveDataGB = 0.0
            TotalArchiveDeletedItemsGB = 0.0
            LicenseCounts = @{
                MigrationWizMailbox = 0
                UserMigrationBundle = 0
                MailboxesOver50GB = 0
            }
        }
        LicenseRequirements = @{
            MigrationWizMailbox = 0
            UserMigrationBundle = 0
            MailboxesOver50GB = 0
            ArchiveOver100GB = 0
        }
        DetailedBreakdown = @()
    }

    function Convert-SizeStringToGB {
        param($value)
        if ([string]::IsNullOrWhiteSpace($value)) { return 0 }
        # If value is like "139.6 MB (146,427,928 bytes)" or "9.986 KB (10,226 bytes)"
        if ($value -match "^([\d.,]+)\s*(B|KB|MB|GB|TB)") {
            $num = [double]::Parse($matches[1])
            $unit = $matches[2].ToUpper()
            switch ($unit) {
                'TB' { return $num * 1024 }
                'GB' { return $num }
                'MB' { return $num / 1024 }
                'KB' { return $num / 1MB }
                'B'  { return $num / 1GB }
                default { return 0 }
            }
        } else {
            return 0
        }
    }

    # Helper function for mailbox stat aggregation and license logic
    function Process-Mailbox {
        param(
            [object]$mailbox,
            [bool]$IsInactive = $false
        )

        # Determine mailbox group based on IsInactive
        $group = if ($IsInactive) { $analysis.InactiveMailboxes } else { $analysis.ActiveMailboxes }

        $type = $mailbox.RecipientTypeDetails.ToString()
        if (-not $group.ByType.ContainsKey($type)) {
            $group.ByType[$type] = 0
        }
        $group.ByType[$type]++

        # Use mailbox stats from the mailbox object directly
        $mailboxSizeGB = Convert-SizeStringToGB $mailbox.TotalItemSize
        $deletedItemSizeGB = Convert-SizeStringToGB $mailbox.TotalDeletedItemSize

        $archiveSizeGB = 0
        $archiveDeletedItemSizeGB = 0

        $hasArchive = $false
        if ($null -ne $mailbox.ArchiveStatus -and $mailbox.ArchiveStatus -eq 'Active') {
            $hasArchive = $true
            $archiveSizeGB = Convert-SizeStringToGB $mailbox.ArchiveTotalItemSize
            $archiveDeletedItemSizeGB = Convert-SizeStringToGB $mailbox.ArchiveTotalDeletedItemSize
        }

        $group.TotalMailboxDataGB += $mailboxSizeGB
        $group.TotalDeletedItemsGB += $deletedItemSizeGB
        $group.TotalArchiveDataGB += $archiveSizeGB
        $group.TotalArchiveDeletedItemsGB += $archiveDeletedItemSizeGB # <-- NEW: accumulate archive deleted items

        $totalSizeWithDeletedGB = $mailboxSizeGB + $deletedItemSizeGB

        if ($hasArchive) {
            $group.WithArchive++
            if ($archiveSizeGB -gt 100) {
                $group.ArchiveOver100GB++
                $analysis.LicenseRequirements.ArchiveOver100GB++
            }
        }

        if (($totalSizeWithDeletedGB -gt 50) -and ($mailboxSizeGB -le 50)) {
            $group.Over50GBWithDeleted++
        }

        # License calculation and stats
        if ($hasArchive) {
            $group.LicenseCounts.UserMigrationBundle++
            $analysis.LicenseRequirements.UserMigrationBundle++
        } else {
            $licenseCount = if ($mailboxSizeGB -le 50) { 1 } else { 2 }
            $group.LicenseCounts.MigrationWizMailbox += $licenseCount
            $analysis.LicenseRequirements.MigrationWizMailbox += $licenseCount
            if ($mailboxSizeGB -gt 50) {
                $group.LicenseCounts.MailboxesOver50GB++
                $analysis.LicenseRequirements.MailboxesOver50GB++
                $group.Over50GB++
            }
            if ($mailboxSizeGB -gt 100) {
                $group.Over100GB++
            }
        }

        # Add breakdown row
        $analysis.DetailedBreakdown += [PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            MailboxType = $type
            MailboxSizeGB = $mailboxSizeGB
            HasArchive = $hasArchive
            ArchiveSizeGB = $archiveSizeGB
            ArchiveDeletedItemSizeGB = $archiveDeletedItemSizeGB # <-- NEW: add to breakdown
            IsInactive = $IsInactive
            LicenseType = if ($hasArchive) { "User Migration Bundle" } 
                        elseif ($mailboxSizeGB -le 50) { "MigrationWiz-Mailbox" }
                        else { "MigrationWiz-Mailbox x2" }
            LicenseCount = if ($hasArchive) { 1 } 
                        elseif ($mailboxSizeGB -le 50) { 1 }
                        else { 2 }
        }
    }

    # --- Process Active and Inactive Mailboxes ---
    $analysis.ActiveMailboxes.Total = $MailboxData.ActiveMailboxes.Count
    foreach ($mailbox in $MailboxData.ActiveMailboxes) {
        Process-Mailbox -mailbox $mailbox -IsInactive:$false
    }

    $analysis.InactiveMailboxes.Total = $MailboxData.InactiveMailboxes.Count
    foreach ($mailbox in $MailboxData.InactiveMailboxes) {
        Process-Mailbox -mailbox $mailbox -IsInactive:$true
    }

    return $analysis
}

#endregion
# ========================================

#region SECTION 6: REPORTING FUNCTIONS

function New-HtmlReport {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Analysis,

        [Parameter(Mandatory = $true)]
        [hashtable]$MailboxData,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [object]$OrgConfig
    )

    $reportDate = Get-Date -Format "MMMM dd, yyyy 'at' hh:mm tt"
    $orgName = $OrgConfig.DisplayName

    # Calculate license requirements for both active and inactive mailboxes
    $active_MigWiz = $Analysis.ActiveMailboxes.LicenseCounts.MigrationWizMailbox
    $active_Bundle = $Analysis.ActiveMailboxes.LicenseCounts.UserMigrationBundle
    $active_Over50 = $Analysis.ActiveMailboxes.LicenseCounts.MailboxesOver50GB

    $inactive_MigWiz = $Analysis.InactiveMailboxes.LicenseCounts.MigrationWizMailbox
    $inactive_Bundle = $Analysis.InactiveMailboxes.LicenseCounts.UserMigrationBundle
    $inactive_Over50 = $Analysis.InactiveMailboxes.LicenseCounts.MailboxesOver50GB

    # For summary as a rollup
    $totalMigrationWizLicenses = ($active_MigWiz + $inactive_MigWiz)
    $totalBundleLicenses = ($active_Bundle + $inactive_Bundle)
    $totalLicensesNeeded = $totalMigrationWizLicenses + $totalBundleLicenses

    $mailboxesOver50GB = ($active_Over50 + $inactive_Over50)

    # Gather Archive Deleted Item data
    $activeArchiveDeletedGB   = [double]($Analysis.ActiveMailboxes.TotalArchiveDeletedItemsGB)
    $inactiveArchiveDeletedGB = [double]($Analysis.InactiveMailboxes.TotalArchiveDeletedItemsGB)

    # Replace call to Get-HtmlStyle with inline CSS
    $inlineCss = @"
body {
    font-family: 'Segoe UI', Arial, sans-serif;
    background: #f4f6fa;
    margin: 0;
    padding: 0;
    color: #222;
}
.container {
    max-width: 1100px;
    margin: 32px auto;
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 0 10px 1px #c7d1e8;
    padding: 28px 30px 40px 30px;
}
.header {
    border-bottom: 2px solid #e3e7ee;
    margin-bottom: 18px;
    padding-bottom: 12px;
}
.header h1 {
    margin: 0;
    font-size: 2.2em;
    letter-spacing: 1px;
}
.subtitle {
    color: #506690;
    font-size: 1.08em;
    margin-top: 2px;
}
.report-info > p {
    font-size: 1em;
    margin: 5px 0 0 8px;
}
.content {
    margin-top: 15px;
}
.executive-summary {
    padding: 15px 20px 24px 20px;
    background: linear-gradient(120deg, #eaf6fe 55%, #eef3f7 100%);
    border-radius: 10px;
    box-shadow: 0 2px 8px -3px #b5bedc55;
    margin-bottom: 23px;
}
.summary-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 22px;
}
.summary-card {
    background: #fff;
    padding: 22px 18px;
    border-radius: 10px;
    flex: 1 0 220px;
    box-shadow: 0 2px 10px -6px #b5bedc8c;
    min-width: 210px;
    margin-bottom: 12px;
}
.summary-card.success {
    background: linear-gradient(120deg, #dbf2e6 60%, #f2fcfe 100%);
    border-left: 5px solid #48bb78;
}
.summary-card.info {
    border-left: 5px solid #3182ce;
    background: linear-gradient(120deg, #eaedfc 60%, #f2fcfe 100%);
}
.summary-card.warning {
    border-left: 5px solid #ffc107;
    background: linear-gradient(120deg, #fff6d1 60%, #fdfcff 100%);
}
.summary-card .value {
    font-size: 2.1em;
    font-weight: 700;
    margin: 6px 0 0 0;
    color: #305e90;
}
.section-title {
    font-size: 1.3em;
    color: #3e5c88;
    margin-bottom: 10px;
    margin-top: 12px;
}
.section {
    margin-top: 44px;
    margin-bottom: 30px;
}
.data-breakdown table {
    font-size: 1em;
    width: 100%;
    box-shadow: 0 1px 2px #b5bedc1e;
    margin-bottom: 0;
}
.data-breakdown th, .data-breakdown td {
    font-size: 1em;
    padding: 10px 8px;
}
.data-breakdown th {
    background: #f0f3f7;
    color: #275899;
    text-align: left;
}
.data-breakdown tr:nth-child(even) td {
    background: #f9fafc;
}
.data-breakdown tr:last-child td {
    font-weight: 600;
}
.license-breakdown {
    background: #f5f8fc;
    border-radius: 7px;
    margin: 13px 0 15px 0;
    padding: 18px 12px 8px 12px;
}
.license-breakdown-header {
    background: #42569F;
    color: #fff;
    padding: 8px 17px;
    border-radius: 6px 6px 0 0;
    font-size: 1.1em;
    font-weight: 600;
}
.license-breakdown-content {
    margin-top: 0px;
    padding: 8px 10px 12px 15px;
}
.license-item {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    padding: 14px 8px;
    border-bottom: 1px solid #e3e5e9;
}
.license-item:last-child { border-bottom: none; }
.license-item .name { font-weight: bold; font-size: 1.05em;}
.license-item .description { font-size: 0.97em; color: #555; margin-bottom: 2px;}
.license-item .count { font-size: 1.3em; font-weight: bold; color: #225574; }
.warning-box {
    background: #fffbe8;
    border-left: 5px solid #ffc107;
    margin: 18px 0 10px 0;
    padding: 16px;
    font-size: 1.03em;
    line-height: 1.45em;
    border-radius: 7px;
}
.inactive-section {
    margin-top: 29px;
    margin-bottom: 15px;
    padding: 19px 20px 18px 20px;
    background: linear-gradient(120deg, #fff5db 50%, #f9f5ee 100%);
    border-radius: 12px;
    box-shadow: 0 2px 8px -4px #ffdca355;
}
.note {
    margin-top: 10px;
    font-size: 1.03em;
    color: #504a1f;
    padding: 9px 18px;
    background: #f7f6f2;
    border-radius: 7px;
}
table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 20px;
    margin-top: 8px;
}
th, td {
    padding: 11px 10px;
    text-align: left;
    border-bottom: 1px solid #dde3ef;
}
th {
    background: #f7fafc;
    color: #233e65;
}
tr:nth-child(even) td {
    background: #f6fbff;
}
.footer {
    text-align: center;
    color: #65748a;
    margin-top: 48px;
    font-size: 1.04em;
    border-top: 1px solid #e6e9ef;
    padding-top: 22px;
}
@media (max-width: 900px) {
    .container { padding: 15px; }
    .summary-grid { flex-direction: column; gap: 12px;}
}
"@

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BitTitan License Assessment - $orgName</title>
    <style>
    $inlineCss
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>BitTitan License Assessment</h1>
            <div class="subtitle">Exchange Online Migration Analysis</div>
        </div>
        
        <div class="report-info">
            <p><strong>Organization:</strong> $orgName</p>
            <p><strong>Report Generated:</strong> $reportDate</p>
            <p><strong>Total Mailboxes Analyzed:</strong> $($Analysis.ActiveMailboxes.Total + $Analysis.InactiveMailboxes.Total)</p>
        </div>
        
        <div class="content">
            <!-- Executive Summary -->
            <div class="executive-summary">
                <h2 style="margin-bottom: 20px; color: #333;">Executive Summary</h2>
                
                <div class="summary-grid">
                    <div class="summary-card success">
                        <h3>Total Licenses Required </h3>
                        <div class="value">$totalLicensesNeeded</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">User Migration Bundle + MigrationWiz-Mailbox (includes Inactive Mailboxes)</p>
                    </div>

                    <div class="summary-card">
                        <h3>User Migration Bundles</h3>
                        <div class="value">$totalBundleLicenses</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Mailboxes with Archive (Active + Inactive)</p>
                    </div>
                    
                    <div class="summary-card">
                        <h3>MigrationWiz-Mailbox</h3>
                        <div class="value">$totalMigrationWizLicenses</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Mailboxes up to 50 GB (1 license per mailbox) (Active + Inactive)</p>
                    </div>
                    
                    <div class="summary-card warning">
                        <h3>Mailboxes Over 50 GB</h3>
                        <div class="value">$($Analysis.ActiveMailboxes.Over50GBWithDeleted + $Analysis.InactiveMailboxes.Over50GBWithDeleted)</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;"> May require additional licenses if using MigrationWiz-Mailbox (x2) (Active + Inactive) (includes deleted items)</p>
                    </div>

                </div>
                
                <div class="note" style="margin-top: 30px;">
                    <strong>Note:</strong> This assessment includes all active and inactive (soft-deleted) mailboxes. Totals above include both mailbox states. This is an estimate and would advise also including a small buffer for additional mailboxes or data that may not be included in the assessment.
                </div>
            </div>
            
            <!-- Mailbox Data Migration Summary -->
            <div class="section">
                <h2 class="section-title">Data Migration Summary</h2>
                
                <div class="summary-grid">
                    <div class="summary-card info">
                        <h3>Total Mailbox Data to Migrate</h3>
                        <div class="value">$(Format-DataSize -SizeInGB ($Analysis.ActiveMailboxes.TotalMailboxDataGB + $Analysis.ActiveMailboxes.TotalDeletedItemsGB + $Analysis.ActiveMailboxes.TotalArchiveDataGB + $activeArchiveDeletedGB + $Analysis.InactiveMailboxes.TotalMailboxDataGB + $Analysis.InactiveMailboxes.TotalDeletedItemsGB + $Analysis.InactiveMailboxes.TotalArchiveDataGB + $inactiveArchiveDeletedGB))</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Mailboxes + Archives (includes deleted items)</p>
                    </div>
                    
                    <div class="summary-card">
                        <h3>Total Mailbox Deleted Items</h3>
                        <div class="value">$(Format-DataSize -SizeInGB ($Analysis.ActiveMailboxes.TotalDeletedItemsGB + $Analysis.InactiveMailboxes.TotalDeletedItemsGB) )</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Mailboxes (active + inactive)</p>
                    </div>
                    
                    <div class="summary-card">
                        <h3>Total Archive Deleted Items</h3>
                        <div class="value">$(Format-DataSize -SizeInGB ($activeArchiveDeletedGB + $inactiveArchiveDeletedGB))</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Archives (active + inactive)</p>
                    </div>

                    <div class="summary-card">
                        <h3>Average Mailbox Size</h3>
                        <div class="value">$([math]::Round((($Analysis.ActiveMailboxes.TotalMailboxDataGB + $Analysis.InactiveMailboxes.TotalMailboxDataGB) / [math]::Max(($Analysis.ActiveMailboxes.Total + $Analysis.InactiveMailboxes.Total), 1)), 2)) GB</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Across all mailboxes (active + inactive)</p>
                    </div>
                </div>
                
                <div class="data-breakdown" style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 8px;">
                    <h3 style="margin-bottom: 15px; color: #333;">Detailed Data Breakdown</h3>
                    <table style="width: 100%; border-collapse: collapse;">
                        <thead>
                            <tr style="background-color: #e9ecef;">
                                <th style="padding: 10px; text-align: left; border-bottom: 2px solid #dee2e6;">Category</th>
                                <th style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;">Mailbox Data</th>
                                <th style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;">Deleted Items</th>
                                <th style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;">Archive Data</th>
                                <th style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;">Archive Deleted Items</th>
                                <th style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="padding: 10px; border-bottom: 1px solid #dee2e6;"><strong>Active Mailboxes</strong></td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $Analysis.ActiveMailboxes.TotalMailboxDataGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $Analysis.ActiveMailboxes.TotalDeletedItemsGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $Analysis.ActiveMailboxes.TotalArchiveDataGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $activeArchiveDeletedGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;"><strong>$(Format-DataSize -SizeInGB ($Analysis.ActiveMailboxes.TotalMailboxDataGB + $Analysis.ActiveMailboxes.TotalDeletedItemsGB + $Analysis.ActiveMailboxes.TotalArchiveDataGB + $activeArchiveDeletedGB))</strong></td>
                            </tr>
                            <tr>
                                <td style="padding: 10px; border-bottom: 1px solid #dee2e6;"><strong>Inactive Mailboxes</strong></td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $Analysis.InactiveMailboxes.TotalMailboxDataGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $Analysis.InactiveMailboxes.TotalDeletedItemsGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $Analysis.InactiveMailboxes.TotalArchiveDataGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;">$(Format-DataSize -SizeInGB $inactiveArchiveDeletedGB)</td>
                                <td style="padding: 10px; text-align: right; border-bottom: 1px solid #dee2e6;"><strong>$(Format-DataSize -SizeInGB ($Analysis.InactiveMailboxes.TotalMailboxDataGB + $Analysis.InactiveMailboxes.TotalDeletedItemsGB + $Analysis.InactiveMailboxes.TotalArchiveDataGB + $inactiveArchiveDeletedGB))</strong></td>
                            </tr>
                            <tr style="background-color: #e3f2fd;">
                                <td style="padding: 10px; border-bottom: 2px solid #dee2e6;"><strong>GRAND TOTAL</strong></td>
                                <td style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;"><strong>$(Format-DataSize -SizeInGB ($Analysis.ActiveMailboxes.TotalMailboxDataGB + $Analysis.InactiveMailboxes.TotalMailboxDataGB))</strong></td>
                                <td style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;"><strong>$(Format-DataSize -SizeInGB ($Analysis.ActiveMailboxes.TotalDeletedItemsGB + $Analysis.InactiveMailboxes.TotalDeletedItemsGB))</strong></td>
                                <td style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;"><strong>$(Format-DataSize -SizeInGB ($Analysis.ActiveMailboxes.TotalArchiveDataGB + $Analysis.InactiveMailboxes.TotalArchiveDataGB))</strong></td>
                                <td style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;"><strong>$(Format-DataSize -SizeInGB ($activeArchiveDeletedGB + $inactiveArchiveDeletedGB))</strong></td>
                                <td style="padding: 10px; text-align: right; border-bottom: 2px solid #dee2e6;"><strong>$(Format-DataSize -SizeInGB ($Analysis.ActiveMailboxes.TotalMailboxDataGB + $Analysis.ActiveMailboxes.TotalDeletedItemsGB + $Analysis.ActiveMailboxes.TotalArchiveDataGB + $activeArchiveDeletedGB + $Analysis.InactiveMailboxes.TotalMailboxDataGB + $Analysis.InactiveMailboxes.TotalDeletedItemsGB + $Analysis.InactiveMailboxes.TotalArchiveDataGB + $inactiveArchiveDeletedGB))</strong></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                
                $(
                    if (($Analysis.ActiveMailboxes.Over50GBWithDeleted + $Analysis.InactiveMailboxes.Over50GBWithDeleted) -gt 0) {
                        "<div class='warning-box' style='margin-top: 20px; padding: 15px; background-color: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;'>
                            <strong>⚠️ Important:</strong> $($Analysis.ActiveMailboxes.Over50GBWithDeleted + $Analysis.InactiveMailboxes.Over50GBWithDeleted) mailbox(es) exceed 50GB in size when deleted items are included. 
                            Consider whether deleted items need to be migrated, as this may affect licensing requirements. If Deleted Items are not needed, consider purging them to reduce the size of the mailbox. $mailboxesOver50GB mailboxes exceed 50GB when deleted items are not included.
                        </div>"
                    }
                )
            </div>
            
            <!-- License Breakdown -->
            <div class="section">
                <h2 class="section-title">License Requirements Breakdown</h2>

                <!-- Active Mailboxes Breakdown -->
                <div class="license-breakdown">
                    <div class="license-breakdown-header">Active Mailboxes - License Requirements</div>
                    <div class="license-breakdown-content">
                        <div class="license-item">
                                <div>
                                    <div class="name">User Migration Bundle</div>
                                    <div class="description">For any active mailbox with archive enabled (regardless of size)</div>
                                </div>
                                <div class="count">$active_Bundle</div>
                        </div>
                        <div class="license-item">
                            <div>
                                <div class="name">MigrationWiz-Mailbox</div>
                                <div class="description">Total licenses needed for active mailboxes without archives (1 license per mailbox up to 50 GB, 2 licenses per mailbox over 50 GB)</div>
                            </div>
                            <div class="count">$active_MigWiz</div>
                        </div>
                        <div class="license-item" style="background-color: #fff3cd;">
                            <div>
                                <div class="name">↳ Mailboxes Over 50 GB</div>
                                <div class="description">Informational: These active mailboxes each require 2 licenses (already counted above) (does not include deleted items)</div>
                            </div>
                            <div class="count" style="color: #856404;">$active_Over50 mailboxes</div>
                        </div>
                        
"@
    if ($Analysis.ActiveMailboxes.LicenseCounts.ArchiveOver100GB -gt 0) {
        $html += @"
                        <div class="license-item">
                            <div>
                                <div class="name" style="color: #ffc107;">⚠️ Archives Over 100 GB</div>
                                <div class="description">These active mailbox archives may require special handling or additional User Migration Bundles</div>
                            </div>
                            <div class="count" style="color: #ffc107;">$($Analysis.ActiveMailboxes.LicenseCounts.ArchiveOver100GB)</div>
                        </div>
"@
    }
    $html += @"
                    </div>
                </div>
                <!-- Inactive Mailboxes Breakdown -->
                <div class="license-breakdown" style="margin-top: 30px;">
                    <div class="license-breakdown-header" style="background: #856404;">Inactive Mailboxes - License Requirements</div>
                    <div class="license-breakdown-content">
                        <div class="license-item">
                            <div>
                                <div class="name">User Migration Bundle</div>
                                <div class="description">For any inactive mailbox with archive enabled (regardless of size)</div>
                            </div>
                            <div class="count">$inactive_Bundle</div>
                        </div>
                        <div class="license-item">
                            <div>
                                <div class="name">MigrationWiz-Mailbox</div>
                                <div class="description">Total licenses needed for inactive mailboxes without archives (1 license per mailbox up to 50 GB, 2 licenses per mailbox over 50 GB)</div>
                            </div>
                            <div class="count">$inactive_MigWiz</div>
                        </div>
                        <div class="license-item" style="background-color: #fff3cd;">
                            <div>
                                <div class="name">↳ Mailboxes Over 50 GB</div>
                                <div class="description">Informational: These inactive mailboxes each require 2 licenses (already counted above) (does not include deleted items)</div>
                            </div>
                            <div class="count" style="color: #856404;">$inactive_Over50 mailboxes</div>
                        </div>
                        
"@
    if ($Analysis.InactiveMailboxes.LicenseCounts.ArchiveOver100GB -gt 0) {
        $html += @"
                        <div class="license-item">
                            <div>
                                <div class="name" style="color: #ffc107;">⚠️ Archives Over 100 GB</div>
                                <div class="description">These inactive mailbox archives may require special handling or additional User Migration Bundles</div>
                            </div>
                            <div class="count" style="color: #ffc107;">$($Analysis.InactiveMailboxes.LicenseCounts.ArchiveOver100GB)</div>
                        </div>
"@
    }
    $html += @"
                    </div>
                </div>
            </div>
            
            <!-- Active Mailboxes Statistics -->
            <div class="section">
                <h2 class="section-title">Active Mailboxes Statistics</h2>
                
                <div class="summary-grid">
                    <div class="summary-card">
                        <h3>Total Active Mailboxes</h3>
                        <div class="value">$($Analysis.ActiveMailboxes.Total)</div>
                    </div>
                    
                    <div class="summary-card">
                        <h3>Mailboxes with Archive</h3>
                        <div class="value">$($Analysis.ActiveMailboxes.WithArchive)</div>
                    </div>
                    
                    <div class="summary-card warning">
                        <h3>Mailboxes Over 50 GB</h3>
                        <div class="value">$($Analysis.ActiveMailboxes.Over50GB)</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Does not include deleted items</p>
                    </div>
                    
                    <div class="summary-card warning">
                        <h3>Archives Over 100 GB</h3>
                        <div class="value">$($Analysis.ActiveMailboxes.ArchiveOver100GB)</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Does not include deleted items</p>
                    </div>
                </div>
                
                <h3 style="margin-top: 30px; margin-bottom: 15px; color: #667eea;">Mailbox Types Breakdown</h3>
                <table>
                    <thead>
                        <tr>
                            <th>Mailbox Type</th>
                            <th>Count</th>
                        </tr>
                    </thead>
                    <tbody>
"@
    foreach ($type in ($Analysis.ActiveMailboxes.ByType.Keys | Sort-Object)) {
        $html += @"
                        <tr>
                            <td>$type</td>
                            <td><strong>$($Analysis.ActiveMailboxes.ByType[$type])</strong></td>
                        </tr>
"@
    }
    $html += @"
                    </tbody>
                </table>
            </div>
            
"@
    # Inactive Mailboxes Section
    if ($Analysis.InactiveMailboxes.Total -gt 0) {
        $html += @"
            <!-- Inactive Mailboxes Section -->
            <div class="inactive-section">
                <h2 class="section-title">Inactive (Soft-Deleted) Mailboxes</h2>
                
                <div class="note">
                    <strong>Important:</strong> These mailboxes are soft-deleted and retained in the tenant. 
                    If these need to be migrated, additional BitTitan licenses will be required.
                </div>
                
                <h3 style="margin-top: 30px; margin-bottom: 15px; color: #856404;">Statistics</h3>
                <div class="summary-grid" style="margin-top: 20px;">
                    <div class="summary-card">
                        <h3>Total Inactive Mailboxes</h3>
                        <div class="value">$($Analysis.InactiveMailboxes.Total)</div>
                    </div>
                    
                    <div class="summary-card">
                        <h3>With Archive</h3>
                        <div class="value">$($Analysis.InactiveMailboxes.WithArchive)</div>
                    </div>
                    
                    <div class="summary-card warning">
                        <h3>Exceed 50GB with Deleted</h3>
                        <div class="value">$($Analysis.InactiveMailboxes.Over50GBWithDeleted)</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Mailboxes exceed 50GB when deleted items are included</p>
                    </div>
                    
                    <div class="summary-card warning">
                        <h3>Archives Over 100 GB</h3>
                        <div class="value">$($Analysis.InactiveMailboxes.ArchiveOver100GB)</div>
                    </div>
                </div>
                $(
                    if ($Analysis.InactiveMailboxes.Over50GBWithDeleted -gt 0) {
                        "<div class='warning-box' style='margin-top: 20px; padding: 15px; background-color: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;'>
                            <strong>⚠️ Important:</strong> $($Analysis.InactiveMailboxes.Over50GBWithDeleted) inactive mailbox(es) exceed 50GB in size when deleted items are included. 
                            Consider whether deleted items need to be migrated, as this may affect licensing requirements.
                        </div>"
                    }
                )
                
                <h3 style="margin-top: 30px; margin-bottom: 15px; color: #856404;">Data Migration Summary</h3>
                <div class="summary-grid">
                    <div class="summary-card info">
                        <h3>Total Data to Migrate</h3>
                        <div class="value">$(Format-DataSize -SizeInGB ($Analysis.InactiveMailboxes.TotalMailboxDataGB + $Analysis.InactiveMailboxes.TotalDeletedItemsGB + $Analysis.InactiveMailboxes.TotalArchiveDataGB + $inactiveArchiveDeletedGB))</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Mailboxes + Archives (includes archive deleted items)</p>
                    </div>
                    
                    <div class="summary-card">
                        <h3>Total Deleted Items</h3>
                        <div class="value">$(Format-DataSize -SizeInGB $Analysis.InactiveMailboxes.TotalDeletedItemsGB)</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Across all inactive mailboxes</p>
                    </div>
                    <div class="summary-card">
                        <h3>Archive Data to Migrate</h3>
                        <div class="value">$(Format-DataSize -SizeInGB $Analysis.InactiveMailboxes.TotalArchiveDataGB)</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Archives (includes archive deleted items)</p>
                    </div>
                    <div class="summary-card">
                        <h3>Total Archive Deleted Items</h3>
                        <div class="value">$(Format-DataSize -SizeInGB $inactiveArchiveDeletedGB)</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Across all inactive archives</p>
                    </div>
                    <div class="summary-card">
                        <h3>Average Mailbox Size</h3>
                        <div class="value">$(Format-DataSize -SizeInGB ($Analysis.InactiveMailboxes.TotalMailboxDataGB / [math]::Max($Analysis.InactiveMailboxes.Total, 1)))</div>
                        <p style="margin-top: 10px; color: #666; font-size: 0.9em;">Inactive mailboxes only</p>
                    </div>

                </div>
            </div>
"@
    }
    $html += @"
        </div>
        <div class="footer">
            <p>Generated by BitTitan License Assessment Tool</p>
            <p>Report Date: $reportDate</p>
        </div>
    </div>
</body>
</html>
"@
    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "`nHTML report generated successfully!" -ForegroundColor Green
        Write-Host "Report saved to: $OutputPath" -ForegroundColor Cyan
        return $true
    } catch {
        Write-Host "Error creating HTML report:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $false
    }
}
# endregion

# ========================================
#region SECTION 7: MAIN SCRIPT EXECUTION

Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║       BitTitan License Assessment Tool                     ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

$scriptStartTime = Get-Date

try {
    Write-Host "Checking required modules..." -ForegroundColor Cyan
    if (-not (Test-RequiredModules)) {
        Write-Host "`n❌ Cannot proceed without required modules" -ForegroundColor Red
        Write-Host "`nPlease install the ExchangeOnlineManagement module:" -ForegroundColor Yellow
        Write-Host "Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force`n" -ForegroundColor White
        #exit 1
    }
    Write-Host "✅ All required modules are installed`n" -ForegroundColor Green

    if (-not $ClientId -and -not $TenantId -and -not $CertificateThumbprint) {
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║           Choose Authentication Method                    ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan
        
        Write-Host "1. Certificate-Based (Unattended/Automated)" -ForegroundColor White
        Write-Host "   - No login prompts" -ForegroundColor Gray
        Write-Host "   - Requires app registration with certificate" -ForegroundColor Gray
        Write-Host "   - Best for scheduled/automated runs`n" -ForegroundColor Gray
        
        Write-Host "2. Interactive (Manual Login)" -ForegroundColor White
        Write-Host "   - Login prompt required" -ForegroundColor Gray
        Write-Host "   - Uses your user credentials" -ForegroundColor Gray
        Write-Host "   - Best for ad-hoc assessments`n" -ForegroundColor Gray
        
        $authChoice = Read-Host "Select authentication method (1 or 2)"
        if ($authChoice -eq '1') {
            Write-Host "`n📋 Certificate-Based Authentication Selected" -ForegroundColor Cyan
            Write-Host "Please provide your app registration details:`n" -ForegroundColor White
            $ClientId = Read-Host "Application (Client) ID"
            $TenantId = Read-Host "Tenant ID"
            $CertificateThumbprint = Read-Host "Certificate Thumbprint"
            if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
                Write-Host "`n❌ All three values are required for certificate-based auth" -ForegroundColor Red
                Write-Host "#exiting...`n" -ForegroundColor Yellow
                #exit 1
            }
        } elseif ($authChoice -eq '2') {
            Write-Host "`n👤 Interactive Authentication Selected" -ForegroundColor Cyan
            Write-Host "You will be prompted to sign in...`n" -ForegroundColor White
        } else {
            Write-Host "`n❌ Invalid choice. Please run the script again and select 1 or 2.`n" -ForegroundColor Red
            #exit 1
        }
    }

    try {
        $orgConfig = Connect-ToExchangeOnline -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
    }
    catch {
        Write-Host "`n❌ Cannot proceed without Exchange Online connection" -ForegroundColor Red
        Write-Host "Script terminated.`n" -ForegroundColor Yellow
        #exit 1
    }
    
    if (-not $orgConfig) {
        Write-Host "`n❌ Failed to retrieve organization configuration" -ForegroundColor Red
        Write-Host "Script terminated.`n" -ForegroundColor Yellow
        #exit 1
    }

    # Modern universal mailbox statistics gathering
    $mailboxData = Get-AllMailboxData

    if (-not $mailboxData) {
        throw "Failed to gather mailbox data"
    }

    $analysis = Get-BitTitanLicenseCalculation -MailboxData $mailboxData

    if (-not $analysis) {
        throw "Failed to calculate license requirements"
    }

    if (-not $OutputPath) {
        $defaultFileName = "BitTitan-LicenseAssessment-$($orgConfig.Name.Replace(' ','_'))-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        $desktop = [Environment]::GetFolderPath("Desktop")
        $OutputPath = Join-Path $desktop $defaultFileName
    }

    $outputDir = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path $outputDir)) {
        Write-Host "`n⚠️  Output directory does not exist: $outputDir" -ForegroundColor Yellow
        Write-Host "Creating directory..." -ForegroundColor Cyan
        try {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
            Write-Host "✅ Directory created" -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Failed to create directory" -ForegroundColor Red
            Write-Host "Defaulting to Desktop..." -ForegroundColor Yellow
            $desktop = [Environment]::GetFolderPath("Desktop")
            $OutputPath = Join-Path $desktop $defaultFileName
        }
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Generating HTML Report" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    $reportCreated = New-HtmlReport -Analysis $analysis -MailboxData $mailboxData -OutputPath $OutputPath -OrgConfig $orgConfig

    if ($reportCreated) {
        $scriptEndTime = Get-Date
        $totalTime = $scriptEndTime - $scriptStartTime

        # Format total time properly
        $totalHours = [math]::Floor($totalTime.TotalHours)
        $totalMinutes = $totalTime.Minutes
        $totalSeconds = $totalTime.Seconds

        $timeString = if ($totalHours -gt 0) {
            "$totalHours hour(s), $totalMinutes minute(s), $totalSeconds second(s)"
        } elseif ($totalMinutes -gt 0) {
            "$totalMinutes minute(s), $totalSeconds second(s)"
        } else {
            "$totalSeconds second(s)"
        }

        Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║              Assessment Complete!                      ║" -ForegroundColor Green
        Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Green
        
        Write-Host "📊 Summary:" -ForegroundColor Cyan
        Write-Host "   Active Mailboxes: $($analysis.ActiveMailboxes.Total)" -ForegroundColor White
        Write-Host "   Inactive Mailboxes: $($analysis.InactiveMailboxes.Total)" -ForegroundColor White
        
        $totalLicenses = $analysis.LicenseRequirements.MigrationWizMailbox + $analysis.LicenseRequirements.UserMigrationBundle
        Write-Host "   Total Licenses Required: $totalLicenses" -ForegroundColor White
        Write-Host "     • MigrationWiz-Mailbox: $($analysis.LicenseRequirements.MigrationWizMailbox)" -ForegroundColor Gray
        Write-Host "       (includes $($analysis.LicenseRequirements.MailboxesOver50GB) mailboxes over 50 GB requiring 2 licenses each)" -ForegroundColor DarkGray
        Write-Host "     • User Migration Bundle: $($analysis.LicenseRequirements.UserMigrationBundle)" -ForegroundColor Gray

        if ($analysis.LicenseRequirements.ArchiveOver100GB -gt 0) {
            Write-Host "   ⚠️  Archives Over 100 GB: $($analysis.LicenseRequirements.ArchiveOver100GB)" -ForegroundColor Yellow
        }
        
        Write-Host "`n📁 Report Location:" -ForegroundColor Cyan
        Write-Host "   $OutputPath" -ForegroundColor White
        Write-Host "`n⏱️  Total Time: $timeString" -ForegroundColor Gray
        Write-Host ""
        $openReport = Read-Host "Would you like to open the report now? (Y/N)"
        if ($openReport -eq 'Y' -or $openReport -eq 'y') {
            try {
                Write-Host "Opening report..." -ForegroundColor Cyan -NoNewline
                Start-Process $OutputPath
                Write-Host "✅ Report opened in your default browser" -ForegroundColor Green
            }
            catch {
                Write-Host "⚠️  Could not open report automatically" -ForegroundColor Yellow
                Write-Host "Please open the file manually: $OutputPath" -ForegroundColor White
            }
        }
        Write-Host "`n🎉 The BitTitan License Assessment is complete!" -ForegroundColor Green
        Write-Host "Thank you for using this script." -ForegroundColor Green
        Write-Host "`Happy Migrating and Have Fun!`n" -ForegroundColor Green

    } else {
        throw "Report generation returned false"
    }
}
catch {
    Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Red
    Write-Host "║                 Assessment Failed                      ║" -ForegroundColor Red
    Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Red
    
    Write-Host "Error: $($_.Exception.Message)`n" -ForegroundColor Red
    
    if ($_.Exception.InnerException) {
        Write-Host "Additional Details:" -ForegroundColor Yellow
        Write-Host "$($_.Exception.InnerException.Message)`n" -ForegroundColor Yellow
    }
    
    Write-Host "Stack Trace:" -ForegroundColor Gray
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    
    Write-Host "`n💡 Troubleshooting:" -ForegroundColor Yellow
    Write-Host "1. Module Issues:" -ForegroundColor White
    Write-Host "   Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force`n" -ForegroundColor Gray
    
    Write-Host "2. Connection Issues:" -ForegroundColor White
    Write-Host "   - Verify you have Exchange Administrator role" -ForegroundColor Gray
    Write-Host "   - Check if admin consent was granted for the app registration" -ForegroundColor Gray
    Write-Host "   - Try running without ClientId/TenantId for interactive auth`n" -ForegroundColor Gray
    
    Write-Host "3. Permission Issues:" -ForegroundColor White
    Write-Host "   - Ensure your account has permissions to read mailbox data" -ForegroundColor Gray
    Write-Host "   - Check for Conditional Access policies`n" -ForegroundColor Gray
    
    Write-Host "For more help, see the README.md troubleshooting section`n" -ForegroundColor Cyan
    #exit 1
}
finally {
    try {
        # (Optionally disconnect)
    }
    catch {}
}

#endregion