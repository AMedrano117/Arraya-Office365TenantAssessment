BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    Import-Module (Join-Path $repoRoot 'src/modules/Arraya.M365.Exchange/Arraya.M365.Exchange.psd1') -Force
    $script:HybridSourcePath = Join-Path $repoRoot 'src/modules/Arraya.M365.Exchange/Public/Get-ExchangeHybridConfiguration.ps1'
    $script:ProbeSourcePath = Join-Path $repoRoot 'src/modules/Arraya.M365.Exchange/Private/Invoke-ArrayaExchangeHybridProbe.ps1'

    # The probe helpers are module-private, so exercise them by dot-sourcing into this scope.
    . $script:ProbeSourcePath
}

Describe 'Invoke-ArrayaExchangeHybridProbe' {
    It 'reports success and the returned value when the query answers' {
        $probe = Invoke-ArrayaExchangeHybridProbe -Name 'Test' -ScriptBlock { @('a', 'b') }

        $probe.Succeeded | Should -BeTrue
        $probe.FailureKind | Should -BeNullOrEmpty
        @($probe.Value).Count | Should -Be 2
    }

    It 'reports success with an empty value when the query legitimately finds nothing' {
        $probe = Invoke-ArrayaExchangeHybridProbe -Name 'Test' -ScriptBlock { @() }

        $probe.Succeeded | Should -BeTrue
        @($probe.Value).Count | Should -Be 0
    }

    It 'classifies an access-denied failure as PermissionDenied' {
        $probe = Invoke-ArrayaExchangeHybridProbe -Name 'Test' -ScriptBlock { throw 'Access is denied to this cmdlet.' }

        $probe.Succeeded | Should -BeFalse
        $probe.FailureKind | Should -Be 'PermissionDenied'
    }

    It 'classifies an RBAC failure as PermissionDenied' {
        $probe = Invoke-ArrayaExchangeHybridProbe -Name 'Test' -ScriptBlock { throw "The user isn't assigned to any management role." }

        $probe.FailureKind | Should -Be 'PermissionDenied'
    }

    It 'classifies an unrelated failure as Error' {
        $probe = Invoke-ArrayaExchangeHybridProbe -Name 'Test' -ScriptBlock { throw 'The remote server returned a 503.' }

        $probe.Succeeded | Should -BeFalse
        $probe.FailureKind | Should -Be 'Error'
    }

    It 'reports Unavailable when the required cmdlet is absent from the session' {
        $probe = Invoke-ArrayaExchangeHybridProbe -Name 'Test' -RequiresCommand 'Get-DefinitelyNotARealCmdlet' -ScriptBlock { 'never runs' }

        $probe.Succeeded | Should -BeFalse
        $probe.FailureKind | Should -Be 'Unavailable'
        @($probe.Value).Count | Should -Be 0
    }
}

Describe 'Test-ArrayaExchangeProbePermissionError' {
    It 'matches common Exchange permission wording' {
        foreach ($message in @(
            'Access is denied.',
            'You are not authorized to perform this operation.',
            'Insufficient permissions to run this cmdlet.',
            'Forbidden'
        )) {
            $record = New-Object System.Management.Automation.ErrorRecord (
                [System.Exception]::new($message), 'x', [System.Management.Automation.ErrorCategory]::NotSpecified, $null
            )
            Test-ArrayaExchangeProbePermissionError -ErrorRecord $record | Should -BeTrue -Because "'$message' is a permission failure"
        }
    }

    It 'does not match transient service errors' {
        $record = New-Object System.Management.Automation.ErrorRecord (
            [System.Exception]::new('The service is temporarily unavailable.'), 'x', [System.Management.Automation.ErrorCategory]::NotSpecified, $null
        )

        Test-ArrayaExchangeProbePermissionError -ErrorRecord $record | Should -BeFalse
    }
}

Describe 'Hybrid classification does not answer from failed probes' {
    BeforeAll {
        $script:hybridSource = Get-Content -Path $script:HybridSourcePath -Raw
    }

    It 'no longer suppresses probe errors with empty catch blocks' {
        $script:hybridSource | Should -Not -Match 'catch \{\}'
    }

    It 'queries with -ErrorAction Stop so failures are detectable' {
        $script:hybridSource | Should -Not -Match 'Get-InboundConnector -ErrorAction SilentlyContinue \| Where-Object'
        $script:hybridSource | Should -Match 'Get-MigrationEndpoint -ErrorAction Stop'
        $script:hybridSource | Should -Match 'Get-OrganizationConfig -ErrorAction Stop'
    }

    It 'classifies unanswered probes as Unknown or PermissionDenied rather than None' {
        $script:hybridSource | Should -Match "'PermissionDenied'"
        $script:hybridSource | Should -Match "'Unknown'"
        $script:hybridSource | Should -Match '\$failedDecisionProbes'
    }

    It 'returns a null IsHybridConfigured when evidence is incomplete' {
        $script:hybridSource | Should -Match '\$isHybridConfigured = if \(\$isHybrid\)'
        $script:hybridSource | Should -Match 'IsHybridConfigured\s+= \$isHybridConfigured'
    }

    It 'surfaces which probes did not answer' {
        $script:hybridSource | Should -Match 'UnansweredProbes'
        $script:hybridSource | Should -Match 'EvidenceComplete'
    }

    It 'does not report a definitive false from the outer error handler' {
        $script:hybridSource | Should -Not -Match 'IsHybridConfigured = \$false'
        $script:hybridSource | Should -Match 'IsHybridConfigured   = \$null'
    }
}

Describe 'Hybrid reporting discloses undetermined state' {
    BeforeAll {
        $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $script:htmlSource = Get-Content -Path (Join-Path $repoRoot 'src/scripts/assessments/HTML Scripts/Invoke-HTMLHelperFunctions.ps1') -Raw
    }

    It 'raises a finding when hybrid state could not be determined' {
        $script:htmlSource | Should -Match "HybridStatus -in @\('Unknown', 'PermissionDenied'\)"
        $script:htmlSource | Should -Match 'not the same as confirming the tenant is not hybrid'
    }
}
