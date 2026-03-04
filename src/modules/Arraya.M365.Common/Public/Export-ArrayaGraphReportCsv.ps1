function Export-ArrayaGraphReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,
        [Parameter(Mandatory = $false)]
        [string]$Activity = 'Downloading Microsoft Graph report CSV',
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers
    )

    $tempCsvPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("arraya-graph-report-" + [guid]::NewGuid().ToString('N') + ".csv")
    $resolvedHeaders = @{}
    if ($Headers) {
        $resolvedHeaders = $Headers.Clone()
    }
    elseif ($global:GraphHeaders) {
        $resolvedHeaders = $global:GraphHeaders.Clone()
    }
    elseif ($global:GraphToken) {
        $resolvedHeaders = @{
            'Content-Type'     = 'application/json'
            'Authorization'    = "Bearer $global:GraphToken"
            'ConsistencyLevel' = 'eventual'
        }
    }

    $useSdk = $false
    if (Get-Command -Name Invoke-MgGraphRequest -ErrorAction SilentlyContinue) {
        $mgContext = Get-MgContext -ErrorAction SilentlyContinue
        if ($mgContext) {
            $useSdk = $true
        }
    }

    try {
        if ($useSdk) {
            try {
                Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputFilePath $tempCsvPath -ErrorAction Stop | Out-Null
            }
            catch {
                if ($resolvedHeaders.Count -eq 0) {
                    throw
                }
                $useSdk = $false
            }
        }

        if (-not $useSdk) {
            if ($resolvedHeaders.Count -eq 0) {
                throw "No Graph authentication context is available for '$Activity'. Connect with Microsoft Graph SDK or provide REST headers."
            }

            $response = Invoke-WebRequest -Uri $Uri -Headers $resolvedHeaders -Method GET -MaximumRedirection 5 -ErrorAction Stop
            if ($null -eq $response -or [string]::IsNullOrWhiteSpace($response.Content)) {
                throw "No CSV content returned for '$Activity'."
            }

            Set-Content -Path $tempCsvPath -Value $response.Content -Encoding UTF8 -Force
        }

        if (-not (Test-Path -Path $tempCsvPath)) {
            throw "CSV report file was not created for '$Activity'."
        }

        return @(Import-Csv -Path $tempCsvPath -ErrorAction Stop)
    }
    finally {
        if (Test-Path -Path $tempCsvPath) {
            Remove-Item -Path $tempCsvPath -Force -ErrorAction SilentlyContinue
        }
    }
}
