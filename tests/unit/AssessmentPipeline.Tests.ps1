Describe 'Arraya.M365.AssessmentPipeline' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:manifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.AssessmentPipeline\Arraya.M365.AssessmentPipeline.psd1'
        Import-Module -Name $script:manifestPath -Force -DisableNameChecking
    }

    It 'has a manifest and phase-based public commands' {
        Test-Path $script:manifestPath | Should -BeTrue

        @(
            'Get-ArrayaAssessmentPipelinePhase'
            'New-ArrayaAssessmentPipelineContext'
            'Invoke-ArrayaAssessmentPipeline'
            'Resume-ArrayaAssessmentPipeline'
        ) | ForEach-Object {
            Get-Command -Name $_ -ErrorAction Stop | Should -Not -BeNullOrEmpty
        }
    }

    It 'returns the phase contract in the expected order' {
        $phases = Get-ArrayaAssessmentPipelinePhase
        @($phases.Name) | Should -Be @(
            'Connection'
            'TenantOverview'
            'Identity'
            'Exchange'
            'Collaboration'
            'Endpoint'
            'Governance'
            'Export'
        )
        ($phases | Where-Object Name -eq 'Identity').DependsOn | Should -Be @('TenantOverview')
        ($phases | Where-Object Name -eq 'Export').Checkpointed | Should -BeFalse
    }

    It 'creates a checkpoint-backed pipeline context with pending phase states' {
        $context = New-ArrayaAssessmentPipelineContext -Mode Full -ExportPath $TestDrive -OutputProfile 'SolutionsEngineer' -ReportingMode 'Operator'

        $context.Mode | Should -Be 'Full'
        $context.OutputProfile | Should -Be 'SolutionsEngineer'
        $context.ReportingMode | Should -Be 'Operator'
        $context.CheckpointRoot | Should -Match 'Support[\\/]Pipeline$'
        $context.PhaseStates.Connection.Status | Should -Be 'Pending'
        $context.PhaseStates.Export.CheckpointPath | Should -BeNullOrEmpty
    }

    It 'orchestrates through a named phase and records checkpoints' {
        InModuleScope Arraya.M365.AssessmentPipeline -Parameters @{ Drive = $TestDrive } {
            param($Drive)

            $script:connectionCalls = New-Object System.Collections.Generic.List[object]
            $script:tenantOverviewCalls = New-Object System.Collections.Generic.List[object]
            $script:identityCalls = New-Object System.Collections.Generic.List[object]
            $script:exchangeCalls = New-Object System.Collections.Generic.List[object]
            $script:collaborationCalls = New-Object System.Collections.Generic.List[object]
            $script:endpointCalls = New-Object System.Collections.Generic.List[object]
            $script:bridgeCalls = New-Object System.Collections.Generic.List[object]
            function Test-ArrayaAssessmentPipelineSessionReadiness { return $true }
            function Invoke-ArrayaAssessmentConnectionPhase {
                param($Context)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Connection'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:connectionCalls.Add([pscustomobject]@{
                    Phase = 'Connection'
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Connection'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentTenantOverviewPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'TenantOverview'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:tenantOverviewCalls.Add([pscustomobject]@{
                    Phase             = 'TenantOverview'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'TenantOverview'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentIdentityPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Identity'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:identityCalls.Add([pscustomobject]@{
                    Phase             = 'Identity'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Identity'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentExchangePhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Exchange'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:exchangeCalls.Add([pscustomobject]@{
                    Phase             = 'Exchange'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Exchange'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentCollaborationPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Collaboration'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:collaborationCalls.Add([pscustomobject]@{
                    Phase             = 'Collaboration'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Collaboration'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentEndpointPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Endpoint'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:endpointCalls.Add([pscustomobject]@{
                    Phase             = 'Endpoint'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Endpoint'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaLegacyAssessmentPhaseBridge {
                param($Context, $PhaseName, $InputSnapshotPath, $ReuseCurrentSessions)

                $script:bridgeCalls.Add([pscustomobject]@{
                    Phase               = $PhaseName
                    InputSnapshotPath   = $InputSnapshotPath
                    ReuseCurrentSession = [bool]$ReuseCurrentSessions
                }) | Out-Null

                $checkpointPath = if ($PhaseName -eq 'Export') {
                    $null
                }
                else {
                    Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName $PhaseName
                }

                if (-not [string]::IsNullOrWhiteSpace($checkpointPath)) {
                    '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                }

                return [pscustomobject]@{
                    Phase           = $PhaseName
                    CheckpointPath  = $checkpointPath
                    CompletedAt     = Get-Date
                    DurationSeconds = 1
                }
            }

            $result = Invoke-ArrayaAssessmentPipeline -Mode Full -ThroughPhase Identity -ExportPath $Drive
            @($script:connectionCalls.Phase) | Should -Be @('Connection')
            @($script:tenantOverviewCalls.Phase) | Should -Be @('TenantOverview')
            @($script:identityCalls.Phase) | Should -Be @('Identity')
            $script:bridgeCalls.Count | Should -Be 0
            $script:tenantOverviewCalls[0].InputSnapshotPath | Should -Be $result.PhaseStates.Connection.CheckpointPath
            $script:identityCalls[0].InputSnapshotPath | Should -Be $result.PhaseStates.TenantOverview.CheckpointPath
            $result.PhaseStates.Identity.Status | Should -Be 'Completed'
            Test-Path -Path $result.PhaseStates.Identity.CheckpointPath | Should -BeTrue

            $statePath = Join-Path $result.CheckpointRoot 'pipeline.state.json'
            Test-Path -Path $statePath | Should -BeTrue
        }
    }

    It 'resumes from checkpoint and skips previously completed phases' {
        InModuleScope Arraya.M365.AssessmentPipeline -Parameters @{ Drive = $TestDrive } {
            param($Drive)

            $script:connectionCalls = New-Object System.Collections.Generic.List[object]
            $script:tenantOverviewCalls = New-Object System.Collections.Generic.List[object]
            $script:identityCalls = New-Object System.Collections.Generic.List[object]
            $script:exchangeCalls = New-Object System.Collections.Generic.List[object]
            $script:collaborationCalls = New-Object System.Collections.Generic.List[object]
            $script:endpointCalls = New-Object System.Collections.Generic.List[object]
            $script:bridgeCalls = New-Object System.Collections.Generic.List[object]
            function Test-ArrayaAssessmentPipelineSessionReadiness { return $false }
            function Invoke-ArrayaAssessmentConnectionPhase {
                param($Context)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Connection'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:connectionCalls.Add([pscustomobject]@{
                    Phase = 'Connection'
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Connection'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentTenantOverviewPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'TenantOverview'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:tenantOverviewCalls.Add([pscustomobject]@{
                    Phase             = 'TenantOverview'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'TenantOverview'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentIdentityPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Identity'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:identityCalls.Add([pscustomobject]@{
                    Phase             = 'Identity'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Identity'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentExchangePhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Exchange'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:exchangeCalls.Add([pscustomobject]@{
                    Phase             = 'Exchange'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Exchange'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentCollaborationPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Collaboration'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:collaborationCalls.Add([pscustomobject]@{
                    Phase             = 'Collaboration'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Collaboration'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaAssessmentEndpointPhase {
                param($Context, $InputSnapshotPath)

                $checkpointPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName 'Endpoint'
                '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                $script:endpointCalls.Add([pscustomobject]@{
                    Phase             = 'Endpoint'
                    InputSnapshotPath = $InputSnapshotPath
                }) | Out-Null

                return [pscustomobject]@{
                    Phase          = 'Endpoint'
                    CheckpointPath = $checkpointPath
                }
            }
            function Invoke-ArrayaLegacyAssessmentPhaseBridge {
                param($Context, $PhaseName, $InputSnapshotPath, $ReuseCurrentSessions)

                $script:bridgeCalls.Add([pscustomobject]@{
                    Phase               = $PhaseName
                    ReuseCurrentSession = [bool]$ReuseCurrentSessions
                }) | Out-Null
                $checkpointPath = if ($PhaseName -eq 'Export') { $null } else { Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $Context.CheckpointRoot -PhaseName $PhaseName }
                if (-not [string]::IsNullOrWhiteSpace($checkpointPath)) {
                    '{}' | Set-Content -Path $checkpointPath -Encoding UTF8
                }

                return [pscustomobject]@{
                    Phase           = $PhaseName
                    CheckpointPath  = $checkpointPath
                    CompletedAt     = Get-Date
                    DurationSeconds = 1
                }
            }

            $initial = Invoke-ArrayaAssessmentPipeline -Mode Full -ThroughPhase Exchange -ExportPath $Drive
            @($script:connectionCalls.Phase) | Should -Be @('Connection')
            @($script:tenantOverviewCalls.Phase) | Should -Be @('TenantOverview')
            @($script:identityCalls.Phase) | Should -Be @('Identity')
            @($script:exchangeCalls.Phase) | Should -Be @('Exchange')
            $script:bridgeCalls.Clear()
            $resumed = Resume-ArrayaAssessmentPipeline -CheckpointRoot $initial.CheckpointRoot -ThroughPhase Governance

            @($script:collaborationCalls.Phase) | Should -Be @('Collaboration')
            @($script:endpointCalls.Phase) | Should -Be @('Endpoint')
            @($script:bridgeCalls.Phase) | Should -Be @('Governance')
            $script:bridgeCalls[0].ReuseCurrentSession | Should -BeFalse
            $resumed.PhaseStates.Exchange.Status | Should -Be 'Completed'
            $resumed.PhaseStates.Governance.Status | Should -Be 'Completed'
        }
    }

    It 'runs the native identity phase with delegated legacy steps and a native group step' {
        InModuleScope Arraya.M365.AssessmentPipeline -Parameters @{ Drive = $TestDrive } {
            param($Drive)

            $context = New-ArrayaAssessmentPipelineContext -Mode Full -ExportPath $Drive -OutputProfile 'SolutionsEngineer' -ReportingMode 'Operator'
            $inputSnapshotPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $context.CheckpointRoot -PhaseName 'TenantOverview'
            $inputSnapshot = New-ArrayaTenantSnapshot -Metadata ([ordered]@{
                OutputProfileLabel = 'SolutionsEngineer'
                ReportingMode      = 'Operator'
            })
            Export-ArrayaTenantSnapshot -Snapshot $inputSnapshot -Path $inputSnapshotPath

            $script:legacyIdentityStepCalls = New-Object System.Collections.Generic.List[object]
            $script:groupCalls = New-Object System.Collections.Generic.List[object]
            function Test-ArrayaAssessmentPipelineSessionReadiness { return $true }
            function Invoke-ArrayaAssessmentConnectionPhase { throw 'Connection bootstrap should not be required in this test.' }
            function Invoke-ArrayaAssessmentLegacyPipelineStep {
                param($Context, $PhaseName, $StepName, $InputSnapshotPath, $CheckpointPath, $ProgressTotalSteps, $ProgressStartingStep, $ReuseCurrentSessions)

                $script:legacyIdentityStepCalls.Add([pscustomobject]@{
                    PhaseName            = $PhaseName
                    StepName             = $StepName
                    InputSnapshotPath    = $InputSnapshotPath
                    ProgressTotalSteps   = $ProgressTotalSteps
                    ProgressStartingStep = $ProgressStartingStep
                    ReuseCurrentSessions = [bool]$ReuseCurrentSessions
                }) | Out-Null

                if (-not (Test-Path -Path $CheckpointPath)) {
                    '{}' | Set-Content -Path $CheckpointPath -Encoding UTF8
                }

                return [pscustomobject]@{
                    Phase          = 'Identity'
                    CheckpointPath = $CheckpointPath
                }
            }
            function Get-EntraIDGroups {
                param($detailLevel, $GraphAuthType, $Context)

                $Context.TenantStats['EntraIDGroups'] = @{
                    'group-1' = [pscustomobject]@{
                        ID          = 'group-1'
                        DisplayName = 'Test Group'
                    }
                }
                $script:groupCalls.Add([pscustomobject]@{
                    DetailLevel   = $detailLevel
                    GraphAuthType = @($GraphAuthType) -join ','
                }) | Out-Null
                return $Context.TenantStats['EntraIDGroups']
            }

            $result = Invoke-ArrayaAssessmentIdentityPhase -Context $context -InputSnapshotPath $inputSnapshotPath
            @($script:legacyIdentityStepCalls.PhaseName | Select-Object -Unique) | Should -Be @('Identity')
            @($script:legacyIdentityStepCalls.StepName) | Should -Be @(
                'Users',
                'Admins',
                'Domains',
                'Authentication/SSO configuration',
                'Federation/cross-tenant configuration',
                'Conditional Access policies',
                'MFA registration details'
            )
            $script:legacyIdentityStepCalls[0].ProgressStartingStep | Should -Be 0
            $script:legacyIdentityStepCalls[3].ProgressStartingStep | Should -Be 4
            $script:legacyIdentityStepCalls[0].ReuseCurrentSessions | Should -BeTrue
            @($script:groupCalls.GraphAuthType) | Should -Be @('SDK')

            $snapshot = Import-ArrayaTenantSnapshot -Path $result.CheckpointPath -SkipValidation
            $snapshot.Data.Identity.EntraIDGroups.'group-1'.DisplayName | Should -Be 'Test Group'
        }
    }

    It 'runs the native exchange phase with native collectors and delegated legacy steps' {
        InModuleScope Arraya.M365.AssessmentPipeline -Parameters @{ Drive = $TestDrive } {
            param($Drive)

            $context = New-ArrayaAssessmentPipelineContext -Mode Full -ExportPath $Drive -OutputProfile 'SolutionsEngineer' -ReportingMode 'Operator'
            $inputSnapshotPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $context.CheckpointRoot -PhaseName 'Identity'
            $inputSnapshot = New-ArrayaTenantSnapshot -Metadata ([ordered]@{
                OutputProfileLabel = 'SolutionsEngineer'
                ReportingMode      = 'Operator'
            })
            Export-ArrayaTenantSnapshot -Snapshot $inputSnapshot -Path $inputSnapshotPath

            $script:legacyExchangeStepCalls = New-Object System.Collections.Generic.List[object]
            $script:exchangeCollectorCalls = New-Object System.Collections.Generic.List[string]
            function Test-ArrayaAssessmentPipelineSessionReadiness { return $true }
            function Invoke-ArrayaAssessmentConnectionPhase { throw 'Connection bootstrap should not be required in this test.' }
            function Invoke-ArrayaAssessmentLegacyPipelineStep {
                param($Context, $PhaseName, $StepName, $InputSnapshotPath, $CheckpointPath, $ProgressTotalSteps, $ProgressStartingStep, $ReuseCurrentSessions)

                $script:legacyExchangeStepCalls.Add([pscustomobject]@{
                    PhaseName            = $PhaseName
                    StepName             = $StepName
                    ProgressTotalSteps   = $ProgressTotalSteps
                    ProgressStartingStep = $ProgressStartingStep
                    ReuseCurrentSessions = [bool]$ReuseCurrentSessions
                }) | Out-Null

                $snapshot = if (Test-Path -Path $InputSnapshotPath) {
                    Import-ArrayaTenantSnapshot -Path $InputSnapshotPath -SkipValidation
                }
                else {
                    New-ArrayaTenantSnapshot
                }
                Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $CheckpointPath

                return [pscustomobject]@{
                    Phase          = $PhaseName
                    CheckpointPath = $CheckpointPath
                }
            }
            function Get-AllRecipientDetails {
                param($detailLevel, $Context)
                $Context.TenantStats['AllRecipients'] = @{ 'recipient-1' = [pscustomobject]@{ DisplayName = 'Recipient One' } }
                $script:exchangeCollectorCalls.Add('Recipients') | Out-Null
            }
            function Get-AllExchangeMailboxDetails {
                param($detailLevel, $Context)
                $Context.TenantStats['MailboxFullDetails'] = @{ 'mailbox-1' = [pscustomobject]@{ DisplayName = 'Mailbox One' } }
                $Context.TenantStats['AllMailboxes'] = @{ 'mailbox-1' = [pscustomobject]@{ DisplayName = 'Mailbox One' } }
                $script:exchangeCollectorCalls.Add('Mailboxes') | Out-Null
            }
            function Get-ExchangeGroupDetails {
                param($detailLevel, $Context)
                $Context.TenantStats['ExchangeGroups'] = @{ 'group-1' = [pscustomobject]@{ DisplayName = 'Exchange Group One' } }
                $script:exchangeCollectorCalls.Add('Groups') | Out-Null
            }
            function Get-AllPublicFolderDetails {
                param($detailLevel, $Context)
                $Context.TenantStats['PublicFolders'] = @{ 'folder-1' = [pscustomobject]@{ Name = 'Public Folder One' } }
                $script:exchangeCollectorCalls.Add('PublicFolders') | Out-Null
            }
            function Get-ExchangeHybridConfiguration {
                param($detailLevel, $Context)
                $Context.TenantStats['HybridConfiguration'] = @{ Summary = [pscustomobject]@{ HybridEnabled = $true } }
                $script:exchangeCollectorCalls.Add('Hybrid') | Out-Null
            }
            function Get-MailFlowRulesandConnectors {
                param($detailLevel, $Context)
                $Context.TenantStats['MailFlowConnectors'] = @{ 'connector-1' = [pscustomobject]@{ Id = 'connector-1' } }
                $Context.TenantStats['MailFlowRules'] = @{ 'rule-1' = [pscustomobject]@{ Name = 'Rule One' } }
                $script:exchangeCollectorCalls.Add('MailFlow') | Out-Null
            }
            function Get-ThirdPartySpamFilteringConfig {
                param($Context)
                $Context.TenantStats['SpamFilteringConfig'] = @{ Configuration = [pscustomobject]@{ Uses3rdPartyFiltering = $false } }
                $script:exchangeCollectorCalls.Add('SpamFiltering') | Out-Null
            }
            function Get-SMTPRelayConfiguration {
                param($Context)
                $Context.TenantStats['SMTPRelayConfig'] = @{ Configuration = [pscustomobject]@{ SMTPAuthEnabled = $false } }
                $script:exchangeCollectorCalls.Add('SMTPRelay') | Out-Null
            }

            $result = Invoke-ArrayaAssessmentExchangePhase -Context $context -InputSnapshotPath $inputSnapshotPath

            @($script:exchangeCollectorCalls) | Should -Be @(
                'Recipients',
                'Mailboxes',
                'Groups',
                'PublicFolders',
                'Hybrid',
                'MailFlow',
                'SpamFiltering',
                'SMTPRelay'
            )
            @($script:legacyExchangeStepCalls.PhaseName | Select-Object -Unique) | Should -Be @('Exchange')
            @($script:legacyExchangeStepCalls.StepName) | Should -Be @(
                'Email activity insights',
                'Exchange governance summaries'
            )
            $script:legacyExchangeStepCalls[0].ProgressStartingStep | Should -Be 6
            $script:legacyExchangeStepCalls[1].ProgressStartingStep | Should -Be 9
            $script:legacyExchangeStepCalls[0].ReuseCurrentSessions | Should -BeTrue

            $snapshot = Import-ArrayaTenantSnapshot -Path $result.CheckpointPath -SkipValidation
            $snapshot.Data.Exchange.AllRecipients.'recipient-1'.DisplayName | Should -Be 'Recipient One'
            $snapshot.Data.Exchange.MailboxFullDetails.'mailbox-1'.DisplayName | Should -Be 'Mailbox One'
            $snapshot.Data.Exchange.MailFlowConnectors.'connector-1'.Id | Should -Be 'connector-1'
        }
    }

    It 'runs the native collaboration phase with delegated legacy steps' {
        InModuleScope Arraya.M365.AssessmentPipeline -Parameters @{ Drive = $TestDrive } {
            param($Drive)

            $context = New-ArrayaAssessmentPipelineContext -Mode Full -ExportPath $Drive -OutputProfile 'SolutionsEngineer' -ReportingMode 'Operator'
            $inputSnapshotPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $context.CheckpointRoot -PhaseName 'Exchange'
            $inputSnapshot = New-ArrayaTenantSnapshot -Metadata ([ordered]@{
                OutputProfileLabel = 'SolutionsEngineer'
                ReportingMode      = 'Operator'
            })
            Export-ArrayaTenantSnapshot -Snapshot $inputSnapshot -Path $inputSnapshotPath

            $script:legacyCollaborationStepCalls = New-Object System.Collections.Generic.List[object]
            function Test-ArrayaAssessmentPipelineSessionReadiness { return $true }
            function Invoke-ArrayaAssessmentConnectionPhase { throw 'Connection bootstrap should not be required in this test.' }
            function Get-MgContext { return [pscustomobject]@{ TenantId = 'tenant-id' } }
            function Invoke-ArrayaAssessmentLegacyPipelineStep {
                param($Context, $PhaseName, $StepName, $InputSnapshotPath, $CheckpointPath, $ProgressTotalSteps, $ProgressStartingStep, $ReuseCurrentSessions)

                $script:legacyCollaborationStepCalls.Add([pscustomobject]@{
                    PhaseName            = $PhaseName
                    StepName             = $StepName
                    ProgressTotalSteps   = $ProgressTotalSteps
                    ProgressStartingStep = $ProgressStartingStep
                    ReuseCurrentSessions = [bool]$ReuseCurrentSessions
                }) | Out-Null

                $snapshot = if (Test-Path -Path $InputSnapshotPath) {
                    Import-ArrayaTenantSnapshot -Path $InputSnapshotPath -SkipValidation
                }
                else {
                    New-ArrayaTenantSnapshot
                }
                Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $CheckpointPath

                return [pscustomobject]@{
                    Phase          = $PhaseName
                    CheckpointPath = $CheckpointPath
                }
            }

            $result = Invoke-ArrayaAssessmentCollaborationPhase -Context $context -InputSnapshotPath $inputSnapshotPath

            @($script:legacyCollaborationStepCalls.PhaseName | Select-Object -Unique) | Should -Be @('Collaboration')
            @($script:legacyCollaborationStepCalls.StepName) | Should -Be @(
                'Unified groups',
                'SharePoint/OneDrive sites (API)',
                'Teams voice details',
                'Teams inventory (MGGraph)'
            )
            $script:legacyCollaborationStepCalls[1].ProgressStartingStep | Should -Be 1
            $script:legacyCollaborationStepCalls[3].ProgressStartingStep | Should -Be 3
            $script:legacyCollaborationStepCalls[0].ReuseCurrentSessions | Should -BeTrue
            Test-Path -Path $result.CheckpointPath | Should -BeTrue
        }
    }

    It 'runs the native endpoint phase with delegated legacy steps' {
        InModuleScope Arraya.M365.AssessmentPipeline -Parameters @{ Drive = $TestDrive } {
            param($Drive)

            $context = New-ArrayaAssessmentPipelineContext -Mode Full -ExportPath $Drive -OutputProfile 'SolutionsEngineer' -ReportingMode 'Operator'
            $inputSnapshotPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $context.CheckpointRoot -PhaseName 'Collaboration'
            $inputSnapshot = New-ArrayaTenantSnapshot -Metadata ([ordered]@{
                OutputProfileLabel = 'SolutionsEngineer'
                ReportingMode      = 'Operator'
            })
            Export-ArrayaTenantSnapshot -Snapshot $inputSnapshot -Path $inputSnapshotPath

            $script:legacyEndpointStepCalls = New-Object System.Collections.Generic.List[object]
            function Test-ArrayaAssessmentPipelineSessionReadiness { return $true }
            function Invoke-ArrayaAssessmentConnectionPhase { throw 'Connection bootstrap should not be required in this test.' }
            function Invoke-ArrayaAssessmentLegacyPipelineStep {
                param($Context, $PhaseName, $StepName, $InputSnapshotPath, $CheckpointPath, $ProgressTotalSteps, $ProgressStartingStep, $ReuseCurrentSessions)

                $script:legacyEndpointStepCalls.Add([pscustomobject]@{
                    PhaseName            = $PhaseName
                    StepName             = $StepName
                    ProgressTotalSteps   = $ProgressTotalSteps
                    ProgressStartingStep = $ProgressStartingStep
                    ReuseCurrentSessions = [bool]$ReuseCurrentSessions
                }) | Out-Null

                $snapshot = if (Test-Path -Path $InputSnapshotPath) {
                    Import-ArrayaTenantSnapshot -Path $InputSnapshotPath -SkipValidation
                }
                else {
                    New-ArrayaTenantSnapshot
                }
                Export-ArrayaTenantSnapshot -Snapshot $snapshot -Path $CheckpointPath

                return [pscustomobject]@{
                    Phase          = $PhaseName
                    CheckpointPath = $CheckpointPath
                }
            }

            $result = Invoke-ArrayaAssessmentEndpointPhase -Context $context -InputSnapshotPath $inputSnapshotPath

            @($script:legacyEndpointStepCalls.PhaseName | Select-Object -Unique) | Should -Be @('Endpoint')
            @($script:legacyEndpointStepCalls.StepName) | Should -Be @(
                'Devices',
                'Endpoint operational summaries'
            )
            $script:legacyEndpointStepCalls[1].ProgressStartingStep | Should -Be 1
            $script:legacyEndpointStepCalls[0].ReuseCurrentSessions | Should -BeTrue
            Test-Path -Path $result.CheckpointPath | Should -BeTrue
        }
    }

    It 'builds a tenant overview checkpoint from the prior connection snapshot' {
        InModuleScope Arraya.M365.AssessmentPipeline -Parameters @{ Drive = $TestDrive } {
            param($Drive)

            $context = New-ArrayaAssessmentPipelineContext -Mode Full -ExportPath $Drive -OutputProfile 'SolutionsEngineer' -ReportingMode 'Operator'
            $inputSnapshotPath = Get-ArrayaAssessmentPipelineCheckpointPath -CheckpointRoot $context.CheckpointRoot -PhaseName 'Connection'
            $inputSnapshot = New-ArrayaTenantSnapshot -Metadata ([ordered]@{
                OutputProfileLabel = 'SolutionsEngineer'
                ReportingMode      = 'Operator'
                Tenant             = [ordered]@{
                    DisplayName       = 'Contoso'
                    DefaultDomainName = 'contoso.onmicrosoft.com'
                }
            })
            Export-ArrayaTenantSnapshot -Snapshot $inputSnapshot -Path $inputSnapshotPath

            function Get-ArrayaAssessmentPipelineTenantOverviewInfo {
                return [pscustomobject]@{
                    DisplayName   = 'Contoso'
                    TenantId      = 'tenant-id'
                    InitialDomain = 'contoso.onmicrosoft.com'
                    DefaultDomain = 'contoso.com'
                }
            }
            function Get-ArrayaAssessmentPipelineLicenseSkus {
                return @{
                    sku1 = [pscustomobject]@{
                        SkuPartNumber   = 'SPE_E3'
                        SkuFriendlyName = 'Microsoft 365 E3'
                        PurchasedUnits  = 25
                        ConsumedUnits   = 20
                    }
                }
            }
            function Get-ArrayaAssessmentPipelineAdConnectSyncDetails {
                return [pscustomobject]@{
                    AdConnectConfiguration = [ordered]@{
                        Summary = [pscustomobject]@{
                            OnPremisesSyncEnabled = $true
                        }
                        SyncServices = @()
                        RecentErrors = @()
                        ErrorCount   = 0
                    }
                    PasswordLifecycleSummary = [pscustomobject]@{
                        OnPremisesSyncEnabled = $true
                    }
                }
            }

            $result = Invoke-ArrayaAssessmentTenantOverviewPhase -Context $context -InputSnapshotPath $inputSnapshotPath
            $snapshot = Import-ArrayaTenantSnapshot -Path $result.CheckpointPath -SkipValidation

            $snapshot.Data.Tenant.TenantInfo.DisplayName | Should -Be 'Contoso'
            $snapshot.Data.Identity.LicenseSKUs.sku1.SkuPartNumber | Should -Be 'SPE_E3'
            $snapshot.Data.Tenant.AdConnectConfiguration.Summary.OnPremisesSyncEnabled | Should -BeTrue
            $snapshot.Data.Governance.PasswordLifecycleSummary.OnPremisesSyncEnabled | Should -BeTrue
        }
    }
}
