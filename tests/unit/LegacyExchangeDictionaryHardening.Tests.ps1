Describe 'Legacy and Exchange dictionary membership hardening' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:legacyCollectorPath = Join-Path $script:repoRoot 'src\scripts\migrated\legacy\Get-FullTenantReportDetails.ps1'
        $script:exchangeHelperPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Private\Invoke-ExchangeMailboxStatHelpers.ps1'
        $script:exchangeMailboxPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Public\Get-AllExchangeMailboxDetails.ps1'

        $tokens = $null
        $parseErrors = $null
        $legacyAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $script:legacyCollectorPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $parseErrors | Should -BeNullOrEmpty

        $legacyFunctions = @(
            'Get-AssessmentGraphScopePlanFlag'
            'Get-ArrayaGraphAdminReportSettings'
            'Get-DictionaryValue'
            'Get-CrossTenantNestedValue'
            'Resolve-CrossTenantTrustValue'
            'Get-ExternalExposureValue'
        )
        $legacyDefinitions = @($legacyAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $legacyFunctions -contains $node.Name
        }, $true))
        $legacyDefinitions.Count | Should -Be $legacyFunctions.Count
        foreach ($definition in $legacyDefinitions) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        $namedPropertyAssignment = @($legacyAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq '$getNamedPropertyValue'
        }, $true) | Select-Object -First 1)
        $namedPropertyAssignment.Count | Should -Be 1
        $script:getNamedPropertyValueUnderTest = & ([scriptblock]::Create($namedPropertyAssignment[0].Right.Extent.Text))

        $tokens = $null
        $parseErrors = $null
        $exchangeHelperAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $script:exchangeHelperPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $parseErrors | Should -BeNullOrEmpty
        $exchangeHelperFunctions = @(
            'Convert-ToMailboxGuidKey'
            'Get-MailboxGuidCandidateKeys'
            'Test-MailboxStatCached'
        )
        $exchangeHelperDefinitions = @($exchangeHelperAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $exchangeHelperFunctions -contains $node.Name
        }, $true))
        $exchangeHelperDefinitions.Count | Should -Be $exchangeHelperFunctions.Count
        foreach ($definition in $exchangeHelperDefinitions) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        $tokens = $null
        $parseErrors = $null
        $exchangeMailboxAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $script:exchangeMailboxPath,
            [ref]$tokens,
            [ref]$parseErrors
        )
        $parseErrors | Should -BeNullOrEmpty
        $exchangeMailboxFunctions = @(
            'Test-DictionaryContainsKey'
            'Convert-MailboxDelegateIdentityToText'
            'Resolve-MailboxFromPermissionIdentity'
            'Get-PermissionIdentityCandidateValues'
            'Resolve-RecipientFromPermissionIdentity'
        )
        $exchangeMailboxDefinitions = @($exchangeMailboxAst.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $exchangeMailboxFunctions -contains $node.Name
        }, $true))
        $exchangeMailboxDefinitions.Count | Should -Be $exchangeMailboxFunctions.Count
        foreach ($definition in $exchangeMailboxDefinitions) {
            . ([scriptblock]::Create($definition.Extent.Text))
        }

        function Get-ArrayaGraphResource {
            return $script:graphAdminReportResponse
        }

        function Get-ArrayaObjectValue {
            param(
                [AllowNull()][object]$Object,
                [string[]]$Names
            )

            if ($null -eq $Object) {
                return $null
            }

            foreach ($name in $Names) {
                if ($Object.PSObject.Properties[$name]) {
                    return $Object.$name
                }
            }
            return $null
        }
    }

    Context 'legacy source dictionary readers' {
        It 'reads graph scope flags from generic and concurrent dictionaries without throwing for missing keys' {
            foreach ($dictionaryType in @('Generic', 'Concurrent')) {
                $dictionary = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $dictionary['NeedsReportsData'] = $false

                (Get-AssessmentGraphScopePlanFlag -GraphScopePlan $dictionary -Name 'NeedsReportsData' -Default $true) | Should -BeFalse
                (Get-AssessmentGraphScopePlanFlag -GraphScopePlan $dictionary -Name 'NeedsSitesData' -Default $true) | Should -BeTrue
            }
        }

        It 'reads live Graph admin report settings from generic and concurrent dictionaries' {
            foreach ($dictionaryType in @('Generic', 'Concurrent')) {
                $script:graphAdminReportResponse = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $script:graphAdminReportResponse['displayConcealedNames'] = $true

                $result = Get-ArrayaGraphAdminReportSettings -SuppressProgress
                $result.Available | Should -BeTrue
                $result.DisplayConcealedNames | Should -BeTrue
                $result.IdentifiableNamesInReports | Should -BeFalse

                $script:graphAdminReportResponse = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $missingResult = Get-ArrayaGraphAdminReportSettings -SuppressProgress
                $missingResult.DisplayConcealedNames | Should -BeNullOrEmpty
            }
        }

        It 'reads generic dictionary values in migration, AD Connect, and external-exposure helpers' {
            foreach ($dictionaryType in @('Generic', 'Concurrent')) {
                $dictionary = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $dictionary['DisplayName'] = 'Example tenant'

                (Get-DictionaryValue -Dictionary $dictionary -Key 'DisplayName') | Should -Be 'Example tenant'
                (Get-DictionaryValue -Dictionary $dictionary -Key 'Missing') | Should -BeNullOrEmpty
                (& $script:getNamedPropertyValueUnderTest -Object $dictionary -Names @('DisplayName')) | Should -Be 'Example tenant'
                (& $script:getNamedPropertyValueUnderTest -Object $dictionary -Names @('Missing')) | Should -BeNullOrEmpty
                (Get-ExternalExposureValue -Object $dictionary -Names @('DisplayName')) | Should -Be 'Example tenant'
                (Get-ExternalExposureValue -Object $dictionary -Names @('Missing')) | Should -BeNullOrEmpty
            }
        }

        It 'traverses and resolves cross-tenant values in generic and concurrent dictionaries' {
            foreach ($dictionaryType in @('Generic', 'Concurrent')) {
                $trust = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $automaticTrust = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $automaticTrust['isMfaAccepted'] = $true
                $trust['automaticTrustSettings'] = $automaticTrust

                (Get-CrossTenantNestedValue -SourceObject $trust -PropertyPaths @('automaticTrustSettings.isMfaAccepted')) | Should -BeTrue
                (Get-CrossTenantNestedValue -SourceObject $trust -PropertyPaths @('automaticTrustSettings.missing')) | Should -BeNullOrEmpty

                $nullTrust = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $nullTrust['isMfaAccepted'] = $null
                (Resolve-CrossTenantTrustValue -TrustObject $nullTrust -PropertyNames @('isMfaAccepted')) | Should -Be 'Not configured'
            }
        }
    }

    Context 'Exchange dictionary readers' {
        It 'checks cached mailbox statistics without falling through to an incompatible Contains overload' {
            $mailboxGuid = '11111111-2222-3333-4444-555555555555'
            $mailbox = [pscustomobject]@{ ExchangeGuid = $mailboxGuid }

            foreach ($dictionaryType in @('Generic', 'Concurrent')) {
                $present = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $present[$mailboxGuid] = [pscustomobject]@{ TotalItemSize = '1 GB' }
                $missing = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }

                (Test-MailboxStatCached -MailboxRecord $mailbox -StatsHash $present) | Should -BeTrue
                (Test-MailboxStatCached -MailboxRecord $mailbox -StatsHash $missing) | Should -BeFalse
            }

            $ordered = [ordered]@{}
            $ordered[$mailboxGuid] = $true
            (Test-MailboxStatCached -MailboxRecord $mailbox -StatsHash $ordered) | Should -BeTrue
            (Test-MailboxStatCached -MailboxRecord $mailbox -StatsHash ([ordered]@{})) | Should -BeFalse
        }

        It 'checks generic dictionary membership through the shared Exchange helper' {
            foreach ($dictionaryType in @('Generic', 'Concurrent')) {
                $dictionary = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $dictionary['Present'] = $true

                (Test-DictionaryContainsKey -Dictionary $dictionary -Key 'Present') | Should -BeTrue
                (Test-DictionaryContainsKey -Dictionary $dictionary -Key 'Missing') | Should -BeFalse
            }

            $ordered = [ordered]@{ Present = $true }
            (Test-DictionaryContainsKey -Dictionary $ordered -Key 'Present') | Should -BeTrue
            (Test-DictionaryContainsKey -Dictionary $ordered -Key 'Missing') | Should -BeFalse
        }

        It 'resolves mailbox and recipient indexes backed by generic dictionaries' {
            foreach ($dictionaryType in @('Generic', 'Concurrent')) {
                $mailboxes = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $recipients = if ($dictionaryType -eq 'Generic') {
                    [System.Collections.Generic.Dictionary[string, object]]::new()
                }
                else {
                    [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
                }
                $mailbox = [pscustomobject]@{
                    DisplayName        = 'Migration User'
                    Identity           = 'migration.user@example.com'
                    PrimarySmtpAddress = 'migration.user@example.com'
                    UserPrincipalName  = 'migration.user@example.com'
                }
                $recipient = [pscustomobject]@{
                    DisplayName        = 'Migration Recipient'
                    Identity           = 'migration.recipient@example.com'
                    PrimarySmtpAddress = 'migration.recipient@example.com'
                    Alias              = 'migration.recipient'
                }
                $mailboxes['migration.user@example.com'] = $mailbox
                $recipients['migration.recipient@example.com'] = $recipient
                $tenantStats = @{
                    'AllMailboxes-MailIdentity' = $mailboxes
                    AllMailboxes                = $mailboxes
                    AllRecipients               = $recipients
                }

                (Resolve-MailboxFromPermissionIdentity -TenantStatsHash $tenantStats -Identity 'migration.user@example.com').PrimarySmtpAddress | Should -Be 'migration.user@example.com'
                (Resolve-MailboxFromPermissionIdentity -TenantStatsHash $tenantStats -Identity 'missing@example.com') | Should -BeNullOrEmpty
                (Resolve-RecipientFromPermissionIdentity -TenantStatsHash $tenantStats -Identity 'migration.recipient@example.com').PrimarySmtpAddress | Should -Be 'migration.recipient@example.com'
                (Resolve-RecipientFromPermissionIdentity -TenantStatsHash $tenantStats -Identity 'missing@example.com') | Should -BeNullOrEmpty
            }
        }
    }
}
