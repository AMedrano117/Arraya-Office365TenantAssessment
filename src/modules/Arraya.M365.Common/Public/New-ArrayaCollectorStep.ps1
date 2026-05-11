function New-ArrayaCollectorStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Section,
        [Parameter(Mandatory = $false)]
        [string]$Workload,
        [Parameter(Mandatory = $false)]
        [bool]$Enabled = $true,
        [Parameter(Mandatory = $false)]
        [string]$SkipReason = 'Disabled by output profile',
        [Parameter(Mandatory = $false)]
        [string[]]$Produces = @(),
        [Parameter(Mandatory = $false)]
        [string[]]$DependsOn = @(),
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    return [pscustomobject][ordered]@{
        Name        = $Name
        Section     = $Section
        Workload    = $Workload
        Enabled     = [bool]$Enabled
        SkipReason  = $SkipReason
        Produces    = @($Produces)
        DependsOn   = @($DependsOn)
        ScriptBlock = $ScriptBlock
    }
}
