$targets = @('./src','./tools') | Where-Object { Test-Path $_ }
$issues = Invoke-ScriptAnalyzer -Path $targets -Recurse
if ($issues) { $issues | Format-Table -AutoSize; throw "PSScriptAnalyzer found $($issues.Count) issue(s)." }
Write-Host 'PSScriptAnalyzer passed.'
