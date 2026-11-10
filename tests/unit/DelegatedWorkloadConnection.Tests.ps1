Describe 'Delegated workload module loading and SharePoint sign-in' {
    BeforeAll {
        $collectorPath = Join-Path $PSScriptRoot '..\..\src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $tokens = $null
        $parseErrors = $null
        $collectorAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $collectorPath, [ref]$tokens, [ref]$parseErrors
        )
        $parseErrors | Should -BeNullOrEmpty
        $wantedFunctions = @(
            'Ensure-AssessmentTeamsModuleAvailable'
            'Ensure-AssessmentSharePointModuleAvailable'
            'Invoke-AssessmentSharePointDelegatedConnect'
            'Connect-AssessmentSharePoint'
            'Connect-AssessmentTeams'
        )
        $definitions = @($collectorAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $wantedFunctions -contains $node.Name
        }, $true))
        $definitions.Count | Should -Be $wantedFunctions.Count
        foreach ($definition in $definitions) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        # No tenant connections are made by these tests.
        function Connect-SPOService {
            [CmdletBinding()]
            param([string]$Url, [bool]$UseSystemBrowser)
        }
        function Test-AssessmentSharePointSessionReady { return $false }
        function Get-AssessmentSharePointAdminUrl {
            param([string]$InitialDomain)
            return 'https://contoso-admin.sharepoint.com'
        }
        function Write-AssessmentInteractiveAuthNotice {
            param([string]$ServiceName, [string]$PromptDescription)
        }
        function Get-AssessmentInteractiveTokenFailureOperatorMessage {
            param([string]$ServiceName, [string]$UnderlyingError)
            return $null
        }
        function Write-Log {
            param([string]$Type, [string]$Message, $ExportFileLocation)
        }
        $ExportDetails = @{}
        $script:sharePointManifest = 'C:\Users\Operator\Documents\PowerShell\Modules\Microsoft.Online.SharePoint.PowerShell\16.0.27709.12000\Microsoft.Online.SharePoint.PowerShell.psd1'
    }

    BeforeEach {
        Mock Get-Module { }
        Mock Import-Module { }
        Mock Write-Warning { }
        Mock Write-Host { }
    }

    It 'loads Teams when it is installed only in the PowerShell 7 user module directory' {
        Mock Get-Module { [pscustomobject]@{ Name = 'MicrosoftTeams' } } -ParameterFilter {
            $Name -eq 'MicrosoftTeams' -and $ListAvailable
        }
        Mock Import-Module {
            if ($UseWindowsPowerShell) {
                throw 'Failure from remote command: no valid module file was found in any module directory.'
            }
        }

        Ensure-AssessmentTeamsModuleAvailable | Should -BeTrue
        Should -Invoke Import-Module -Times 1 -Exactly -ParameterFilter { $Name -eq 'MicrosoftTeams' }
    }

    It 'reuses a loaded Teams module without importing or rediscovering it' {
        Mock Get-Module { [pscustomobject]@{ Name = 'MicrosoftTeams' } } -ParameterFilter {
            $Name -eq 'MicrosoftTeams' -and -not $ListAvailable
        }

        Ensure-AssessmentTeamsModuleAvailable | Should -BeTrue
        Should -Invoke Import-Module -Times 0
        Should -Invoke Get-Module -Times 0 -ParameterFilter { $ListAvailable }
    }

    It 'returns unavailable when Teams is absent without trying an import' {
        Ensure-AssessmentTeamsModuleAvailable | Should -BeFalse
        Should -Invoke Import-Module -Times 0
    }

    It 'continues with Graph-only Teams coverage if an installed module cannot load' {
        Mock Get-Module { [pscustomobject]@{ Name = 'MicrosoftTeams' } } -ParameterFilter { $ListAvailable }
        Mock Import-Module { throw 'Teams assembly load failed' }

        $result = Connect-AssessmentTeams -AuthenticationType Interactive

        $result.Teams | Should -BeFalse
        $result.Status | Should -Be 'GraphOnly'
        $result.Message | Should -Match 'could not be imported'
        Should -Invoke Write-Warning -Times 1 -Exactly -ParameterFilter { $Message -match 'Teams assembly load failed' }
    }

    It 'loads the newest SharePoint module natively when it supports PowerShell 7' {
        Mock Get-Module {
            [pscustomobject]@{ Version = [version]'16.0.27011.12008'; Path = 'C:\older\SharePoint.psd1' }
            [pscustomobject]@{ Version = [version]'16.0.27709.12000'; Path = $script:sharePointManifest }
        } -ParameterFilter { $Name -eq 'Microsoft.Online.SharePoint.PowerShell' -and $ListAvailable }
        Mock Import-Module {
            if ($Name -ne $script:sharePointManifest -or $UseWindowsPowerShell) {
                throw 'The newest SharePoint module is available natively in the PS7 user module directory.'
            }
        }

        Ensure-AssessmentSharePointModuleAvailable | Should -BeTrue
        Should -Invoke Import-Module -Times 1 -Exactly
    }

    It 'falls back to the resolved SharePoint manifest when native loading fails and compatibility search cannot find its name' {
        Mock Get-Module {
            [pscustomobject]@{ Version = [version]'16.0.27011.12008'; Path = 'C:\older\SharePoint.psd1' }
            [pscustomobject]@{ Version = [version]'16.0.27709.12000'; Path = $script:sharePointManifest }
        } -ParameterFilter { $Name -eq 'Microsoft.Online.SharePoint.PowerShell' -and $ListAvailable }
        Mock Import-Module {
            if ($Name -ne $script:sharePointManifest -or -not $UseWindowsPowerShell) {
                throw 'SharePoint must load the resolved PS7 user module path in the compatibility session.'
            }
        }

        Ensure-AssessmentSharePointModuleAvailable | Should -BeTrue
        Should -Invoke Import-Module -Times 2 -Exactly
    }

    It 'reports SharePoint import failures without terminating the assessment' {
        Mock Get-Module {
            [pscustomobject]@{ Version = [version]'16.0.27709.12000'; Path = $script:sharePointManifest }
        } -ParameterFilter { $ListAvailable }
        Mock Import-Module { throw 'SharePoint compatibility session unavailable' }

        Ensure-AssessmentSharePointModuleAvailable | Should -BeFalse
        Should -Invoke Write-Warning -Times 1 -Exactly -ParameterFilter { $Message -match 'SharePoint compatibility session unavailable' }
    }

    It 'requests system-browser authentication when supported by SharePoint' {
        Mock Connect-SPOService { }

        Invoke-AssessmentSharePointDelegatedConnect -Url 'https://contoso-admin.sharepoint.com'

        Should -Invoke Connect-SPOService -Times 1 -Exactly -ParameterFilter {
            $Url -eq 'https://contoso-admin.sharepoint.com' -and $UseSystemBrowser -and $ErrorAction -eq 'Stop'
        }
    }

    It 'remains compatible with SharePoint versions without the system-browser parameter' {
        Mock Get-Command { [pscustomobject]@{ Parameters = @{} } } -ParameterFilter { $Name -eq 'Connect-SPOService' }
        Mock Connect-SPOService { }

        Invoke-AssessmentSharePointDelegatedConnect -Url 'https://contoso-admin.sharepoint.com'

        Should -Invoke Connect-SPOService -Times 1 -Exactly -ParameterFilter {
            $Url -eq 'https://contoso-admin.sharepoint.com' -and -not $PSBoundParameters.ContainsKey('UseSystemBrowser')
        }
    }

    It 'preserves the underlying SharePoint sign-in failure in the fallback result' {
        Mock Ensure-AssessmentSharePointModuleAvailable { return $true }
        Mock Connect-SPOService { throw 'AADSTS53003: Access blocked by Conditional Access' }

        $result = Connect-AssessmentSharePoint -AuthenticationType Interactive -InitialDomain 'contoso.onmicrosoft.com'

        $result.SharePointOnline | Should -BeFalse
        $result.Status | Should -Be 'GraphFallback'
        $result.Message | Should -Match 'AADSTS53003'
    }
}

Describe 'SharePoint admin URL discovery' {
    BeforeAll {
        $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        . (Join-Path $repoRoot 'src\modules\Arraya.M365.Common\Public\Get-ArrayaObjectValue.ps1')
        $tokens = $null
        $parseErrors = $null
        $collectorAst = [System.Management.Automation.Language.Parser]::ParseFile(
            (Join-Path $repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'),
            [ref]$tokens, [ref]$parseErrors
        )
        $parseErrors | Should -BeNullOrEmpty
        $definition = $collectorAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -eq 'Get-AssessmentSharePointAdminUrl'
        }, $true)
        . ([scriptblock]::Create($definition.Extent.Text))

        function Get-ArrayaGraphResource {
            [CmdletBinding()]
            param([string]$Uri, [string]$Activity, [switch]$SuppressProgress, [switch]$SuppressAccessDeniedWarning)
        }
    }

    BeforeEach {
        Mock Write-Warning { }
    }

    It 'uses the actual SharePoint hostname when it differs from the initial tenant domain' {
        Mock Get-ArrayaGraphResource { @{ webUrl = 'https://renamedtenant.sharepoint.com/' } }

        Get-AssessmentSharePointAdminUrl -InitialDomain 'originaltenant.onmicrosoft.com' |
            Should -Be 'https://renamedtenant-admin.sharepoint.com'

        Should -Invoke Get-ArrayaGraphResource -Times 1 -Exactly -ParameterFilter {
            $Uri -eq 'https://graph.microsoft.com/v1.0/sites/root?$select=id,webUrl,displayName'
        }
        Should -Invoke Write-Warning -Times 0
    }

    It 'handles Graph dictionary responses without requiring an initial domain' {
        Mock Get-ArrayaGraphResource {
            $site = [System.Collections.Generic.Dictionary[string, object]]::new()
            $site['webUrl'] = 'https://renamedtenant.sharepoint.com/sites/root'
            return $site
        }

        Get-AssessmentSharePointAdminUrl | Should -Be 'https://renamedtenant-admin.sharepoint.com'
    }

    It 'preserves the discovered SharePoint cloud suffix <Suffix>' -TestCases @(
        @{ Suffix = 'sharepoint.com' }
        @{ Suffix = 'sharepoint.us' }
        @{ Suffix = 'sharepoint-mil.us' }
        @{ Suffix = 'sharepoint.cn' }
    ) {
        param($Suffix)
        Mock Get-ArrayaGraphResource { [pscustomobject]@{ webUrl = "https://tenant.$Suffix/" } }

        Get-AssessmentSharePointAdminUrl -InitialDomain 'tenant.onmicrosoft.com' |
            Should -Be "https://tenant-admin.$Suffix"
    }

    It 'warns when Graph discovery fails and falls back to the initial domain' {
        Mock Get-ArrayaGraphResource { throw 'Access denied to root site' }

        Get-AssessmentSharePointAdminUrl -InitialDomain 'contoso.onmicrosoft.com' |
            Should -Be 'https://contoso-admin.sharepoint.com'

        Should -Invoke Write-Warning -Times 1 -Exactly -ParameterFilter {
            $Message -match 'verify the actual admin URL'
        }
    }

    It 'does not derive an admin URL from an unrecognized root hostname' {
        Mock Get-ArrayaGraphResource { @{ webUrl = 'https://tenant.sharepoint.com.example.org/' } }

        Get-AssessmentSharePointAdminUrl -InitialDomain 'contoso.onmicrosoft.com' |
            Should -Be 'https://contoso-admin.sharepoint.com'
    }

    It 'returns no URL when neither root-site discovery nor the initial domain is available' {
        Mock Get-ArrayaGraphResource { $null }

        Get-AssessmentSharePointAdminUrl | Should -BeNullOrEmpty
    }
}
