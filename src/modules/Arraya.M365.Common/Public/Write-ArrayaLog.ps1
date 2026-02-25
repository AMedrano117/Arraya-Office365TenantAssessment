function Write-ArrayaLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('INFO','WARN','ERROR')][string]$Level,
        [Parameter(Mandatory)][string]$Message,
        [string]$Path
    )
    $line = '{0} [{1}] {2}' -f (Get-Date).ToString('o'), $Level, $Message
    if ($Path) { Add-Content -Path $Path -Value $line }
    switch ($Level) {
        'INFO'  { Write-Verbose $Message }
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Error $Message }
    }
}
