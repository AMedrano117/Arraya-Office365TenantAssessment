function Invoke-ArrayaRetry {
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[int]$MaxRetries=3,[int]$InitialDelaySeconds=2)
    $attempt=0
    do {
        try { $attempt++; return & $ScriptBlock }
        catch {
            if ($attempt -ge $MaxRetries) { throw }
            Start-Sleep -Seconds ([math]::Pow($InitialDelaySeconds,$attempt))
        }
    } while ($true)
}
