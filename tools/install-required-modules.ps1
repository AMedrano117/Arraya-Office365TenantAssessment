$modules = @(
  'Pester',
  'PSScriptAnalyzer',
  'Microsoft.Graph.Authentication',
  'ExchangeOnlineManagement',
  'ImportExcel'
)
foreach ($m in $modules) {
  if (-not (Get-Module -ListAvailable -Name $m)) {
    Install-Module $m -Scope CurrentUser -Force -AllowClobber
  }
}
