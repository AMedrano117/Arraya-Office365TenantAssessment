function Test-ArrayaTenantSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [hashtable]$Snapshot,
        [Parameter(Mandatory = $false)]
        [ValidateSet('Export', 'ImprovementPlan', 'Collection')]
        [string]$Purpose = 'Export'
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $schemaVersion = 0

    if (-not $Snapshot) {
        $errors.Add('Snapshot is null.')
    }
    elseif (-not ($Snapshot -is [hashtable])) {
        $errors.Add('Snapshot must be a hashtable.')
    }

    if ($errors.Count -eq 0) {
        if ($Snapshot.Contains('SchemaVersion')) {
            try {
                $schemaVersion = [int]$Snapshot['SchemaVersion']
            }
            catch {
                $warnings.Add("SchemaVersion '$($Snapshot['SchemaVersion'])' is not numeric. Treating as version 0.")
                $schemaVersion = 0
            }
        }
        else {
            $warnings.Add('SchemaVersion was not found. Treating payload as legacy schema.')
        }
    }

    if ($errors.Count -eq 0) {
        if ($schemaVersion -ge 2) {
            foreach ($requiredRoot in @('Metadata', 'CollectionPlan', 'Data', 'Derived', 'Diagnostics')) {
                if (-not $Snapshot.Contains($requiredRoot) -or -not ($Snapshot[$requiredRoot] -is [System.Collections.IDictionary])) {
                    $errors.Add("Schema v2 snapshot is missing required hashtable section '$requiredRoot'.")
                }
            }

            if ($Snapshot.Contains('Data') -and ($Snapshot['Data'] -is [System.Collections.IDictionary])) {
                $requiredDataDomains = @('Exchange', 'Identity', 'Collaboration', 'Security', 'Tenant', 'Governance', 'Other')
                foreach ($domain in $requiredDataDomains) {
                    if (-not $Snapshot['Data'].Contains($domain)) {
                        $warnings.Add("Data domain '$domain' is not present. Exporters may emit partial output.")
                    }
                }
            }

            if ($Purpose -eq 'ImprovementPlan') {
                $derived = $Snapshot['Derived']
                $hasFindings = $derived -and (
                    $derived.Contains('Findings') -or
                    $derived.Contains('BestPracticeFindings')
                )
                if (-not $hasFindings) {
                    $warnings.Add("Derived findings were not found for improvement-plan generation. Rule-engine fallback will be used.")
                }
            }
        }
        else {
            if (-not $Snapshot.Contains('Data') -and $Purpose -eq 'Export') {
                $warnings.Add('Legacy snapshot missing Data wrapper. Import adapter will attempt best-effort conversion.')
            }
        }
    }

    return [PSCustomObject]@{
        Valid         = ($errors.Count -eq 0)
        SchemaVersion = $schemaVersion
        Errors        = @($errors)
        Warnings      = @($warnings)
    }
}
