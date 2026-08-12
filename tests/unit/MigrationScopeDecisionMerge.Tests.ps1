Describe 'Migration scope decision snapshot merge' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:mergeScriptPath = Join-Path $script:repoRoot 'src\scripts\operations\Merge-M365MigrationScopeDecisions.ps1'
        $script:commonManifestPath = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        Import-Module -Name $script:commonManifestPath -Force -DisableNameChecking -WarningAction SilentlyContinue

        function New-TestMigrationDecisionSnapshot {
            return @{
                SchemaVersion  = 2
                Metadata       = @{ Tenant = @{ DisplayName = 'Contoso'; TenantId = 'tenant-id' } }
                CollectionPlan = @{ Profile = 'Presales' }
                Data           = @{
                    Exchange      = @{ AllMailboxes = @([pscustomobject]@{ DisplayName = 'User One'; PrimarySmtpAddress = 'user1@contoso.com' }) }
                    Identity      = @{}
                    Collaboration = @{}
                    Security      = @{}
                    Tenant        = @{}
                    Governance    = @{}
                    Other         = @{}
                }
                Derived        = @{ ExistingDerivedTable = @([pscustomobject]@{ Name = 'Preserved'; Value = 7 }) }
                Diagnostics    = @{}
            }
        }

        function Write-TestMigrationDecisionCsv {
            param([string]$Path, [object[]]$Rows)
            $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
        }
    }

    BeforeEach {
        $script:sourcePath = Join-Path $TestDrive 'source-snapshot.json'
        Export-ArrayaTenantSnapshot -Snapshot (New-TestMigrationDecisionSnapshot) -Path $script:sourcePath
        $script:sourceHashBefore = (Get-FileHash -LiteralPath $script:sourcePath -Algorithm SHA256).Hash
    }

    It 'merges a completed CSV into Derived without changing the source snapshot' {
        $csvPath = Join-Path $TestDrive 'MigrationScopeDecisions.csv'
        $outputPath = Join-Path $TestDrive 'merged-snapshot.json'
        Write-TestMigrationDecisionCsv -Path $csvPath -Rows @(
            [pscustomobject]@{
                DecisionKey = 'MAILBOX:user1@contoso.com'; DecisionPrompt = 'Confirm mailbox disposition'; CustomerConfirmedValue = 'Migrate'
                Status = 'approved'; InScope = 'Yes'; TargetMapping = 'user1@target.example'; MigrationTool = 'BitTitan MigrationWiz'
                DecisionOwner = 'Messaging Lead'; DueDate = '2026-08-20'; Notes = 'Pilot wave'
            },
            [pscustomobject]@{
                DecisionKey = 'BT-06'; DecisionPrompt = 'Evaluate TMB'; CustomerConfirmedValue = 'No'
                Status = 'Open'; InScope = ''; TargetMapping = ''; MigrationTool = ''; DecisionOwner = 'SE'; DueDate = ''; Notes = ''
            }
        )

        $result = & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $csvPath -OutputPath $outputPath

        $result.Written | Should -BeTrue
        $result.DecisionCount | Should -Be 2
        $result.ConfirmedCount | Should -Be 1
        Test-Path -LiteralPath $outputPath | Should -BeTrue
        (Get-FileHash -LiteralPath $script:sourcePath -Algorithm SHA256).Hash | Should -Be $script:sourceHashBefore

        $source = Import-ArrayaTenantSnapshot -Path $script:sourcePath
        $source['Derived'].Contains('MigrationScopeDecisions') | Should -BeFalse
        $merged = Import-ArrayaTenantSnapshot -Path $outputPath
        @($merged['Derived']['MigrationScopeDecisions']).Count | Should -Be 2
        $mailboxDecision = @($merged['Derived']['MigrationScopeDecisions'] | Where-Object { $_['DecisionKey'] -eq 'MAILBOX:user1@contoso.com' } | Select-Object -First 1)[0]
        $mailboxDecision['Status'] | Should -Be 'Approved'
        $mailboxDecision['InScope'] | Should -BeTrue
        $mailboxDecision['TargetMapping'] | Should -Be 'user1@target.example'
        @($merged['Derived']['ExistingDerivedTable'])[0]['Name'] | Should -Be 'Preserved'
        @($merged['Data']['Exchange']['AllMailboxes'])[0]['PrimarySmtpAddress'] | Should -Be 'user1@contoso.com'

        $replayContext = Import-ArrayaTenantSnapshotContext -Path $outputPath -Purpose Export
        @($replayContext.LegacyData['MigrationScopeDecisions']).Count | Should -Be 2
        @($replayContext.LegacyData['MigrationScopeDecisions'] | Where-Object { $_['DecisionKey'] -eq 'MAILBOX:user1@contoso.com' }).Count | Should -Be 1
    }

    It 'supports WhatIf without writing an output snapshot' {
        $csvPath = Join-Path $TestDrive 'whatif.csv'
        $outputPath = Join-Path $TestDrive 'whatif-output.json'
        Write-TestMigrationDecisionCsv -Path $csvPath -Rows @(
            [pscustomobject]@{ DecisionKey = 'BT-09'; Status = 'Confirmed'; InScope = ''; CustomerConfirmedValue = 'Validated' }
        )

        $result = & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $csvPath -OutputPath $outputPath -WhatIf

        $result.Written | Should -BeFalse
        $result.DecisionCount | Should -Be 1
        Test-Path -LiteralPath $outputPath | Should -BeFalse
        (Get-FileHash -LiteralPath $script:sourcePath -Algorithm SHA256).Hash | Should -Be $script:sourceHashBefore
    }

    It 'rejects duplicate DecisionKey values case-insensitively' {
        $csvPath = Join-Path $TestDrive 'duplicate.csv'
        $outputPath = Join-Path $TestDrive 'duplicate-output.json'
        Write-TestMigrationDecisionCsv -Path $csvPath -Rows @(
            [pscustomobject]@{ DecisionKey = 'BT-09'; Status = 'Confirmed' },
            [pscustomobject]@{ DecisionKey = 'bt-09'; Status = 'Approved' }
        )

        { & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $csvPath -OutputPath $outputPath } |
            Should -Throw '*duplicate DecisionKey*BT-09*'
        Test-Path -LiteralPath $outputPath | Should -BeFalse
    }

    It 'rejects rows with a blank DecisionKey' {
        $csvPath = Join-Path $TestDrive 'blank-key.csv'
        $outputPath = Join-Path $TestDrive 'blank-key-output.json'
        Write-TestMigrationDecisionCsv -Path $csvPath -Rows @(
            [pscustomobject]@{ DecisionKey = ''; Status = 'Needs Input'; Notes = 'Accidental row' }
        )

        { & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $csvPath -OutputPath $outputPath } |
            Should -Throw '*blank DecisionKey*'
    }

    It 'rejects blank or unsupported Status values' -ForEach @(
        @{ Case = 'blank'; StatusValue = ''; Pattern = '*blank Status*' },
        @{ Case = 'unsupported'; StatusValue = 'Done-ish'; Pattern = '*invalid Status*Done-ish*' }
    ) {
        $csvPath = Join-Path $TestDrive ("status-{0}.csv" -f $Case)
        $outputPath = Join-Path $TestDrive ("status-{0}.json" -f $Case)
        Write-TestMigrationDecisionCsv -Path $csvPath -Rows @(
            [pscustomobject]@{ DecisionKey = 'BT-09'; Status = $StatusValue }
        )

        { & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $csvPath -OutputPath $outputPath } |
            Should -Throw $Pattern
    }

    It 'refuses to overwrite the source snapshot even with Force' {
        $csvPath = Join-Path $TestDrive 'same-path.csv'
        Write-TestMigrationDecisionCsv -Path $csvPath -Rows @(
            [pscustomobject]@{ DecisionKey = 'BT-09'; Status = 'Confirmed' }
        )

        { & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $csvPath -OutputPath $script:sourcePath -Force } |
            Should -Throw '*must differ from AssessmentJsonPath*'
        (Get-FileHash -LiteralPath $script:sourcePath -Algorithm SHA256).Hash | Should -Be $script:sourceHashBefore
    }

    It 'requires a distinct JSON output path' {
        $csvPath = Join-Path $TestDrive 'unsafe-output.csv'
        Write-TestMigrationDecisionCsv -Path $csvPath -Rows @(
            [pscustomobject]@{ DecisionKey = 'BT-09'; Status = 'Confirmed' }
        )

        { & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $csvPath -OutputPath (Join-Path $TestDrive 'not-json.txt') } |
            Should -Throw '*must use the .json extension*'
    }

    It 'imports the MigrationScopeDecisions worksheet from XLSX when ImportExcel is available' {
        if (-not (Get-Command -Name Export-Excel -ErrorAction SilentlyContinue) -or -not (Get-Command -Name Import-Excel -ErrorAction SilentlyContinue)) {
            Set-ItResult -Skipped -Because 'ImportExcel is not available in this environment.'
            return
        }

        $xlsxPath = Join-Path $TestDrive 'completed-presales.xlsx'
        $outputPath = Join-Path $TestDrive 'xlsx-merged.json'
        @(
            [pscustomobject]@{
                DecisionKey = 'SG-04'; DecisionPrompt = 'Validate endpoints'; CustomerConfirmedValue = 'Test passed'
                Status = 'Resolved'; InScope = $true; TargetMapping = 'Source to destination'; MigrationTool = 'ShareGate'
                DecisionOwner = 'Collaboration Lead'; DueDate = [datetime]'2026-08-19'; Notes = 'Representative site tested'
            }
        ) | Export-Excel -Path $xlsxPath -WorksheetName 'MigrationScopeDecisions' -AutoSize

        $result = & $script:mergeScriptPath -AssessmentJsonPath $script:sourcePath -DecisionInputPath $xlsxPath -OutputPath $outputPath

        $result.Written | Should -BeTrue
        $result.WorksheetName | Should -Be 'MigrationScopeDecisions'
        $merged = Import-ArrayaTenantSnapshot -Path $outputPath
        $decision = @($merged['Derived']['MigrationScopeDecisions'])[0]
        $decision['DecisionKey'] | Should -Be 'SG-04'
        $decision['Status'] | Should -Be 'Resolved'
        $decision['InScope'] | Should -BeTrue
    }
}
