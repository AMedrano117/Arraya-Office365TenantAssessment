$config = New-PesterConfiguration
$config.Run.Path = './tests'
$config.Output.Verbosity = 'Detailed'
$config.TestResult.Enabled = $true
$config.TestResult.OutputPath = './TestResults.xml'
$config.TestResult.OutputFormat = 'NUnitXml'
$config.CodeCoverage.Enabled = $true
$config.CodeCoverage.Path = (Get-ChildItem './src/modules' -Recurse -Filter '*.ps1').FullName
$config.CodeCoverage.OutputPath = './CoverageResults.xml'
$config.Run.PassThru = $true
$result = Invoke-Pester -Configuration $config
if ($result.FailedCount -gt 0) { exit 1 }
