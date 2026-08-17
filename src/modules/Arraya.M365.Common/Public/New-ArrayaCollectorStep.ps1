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
        # Marks a collector whose data the deliverables cannot be trusted without. A required
        # step that FAILS degrades the whole run to Failed; a required step that is skipped
        # because the profile disabled it does not, since that is a deliberate choice.
        [Parameter(Mandatory = $false)]
        [bool]$Required = $false,
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
        Required    = [bool]$Required
        ScriptBlock = $ScriptBlock
    }
}
