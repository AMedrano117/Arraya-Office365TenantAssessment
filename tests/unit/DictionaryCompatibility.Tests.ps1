Describe 'IDictionary implementation compatibility' {
    BeforeAll {
        $script:previousImportModuleWarningPreference = $PSDefaultParameterValues['Import-Module:WarningAction']
        $PSDefaultParameterValues['Import-Module:WarningAction'] = 'SilentlyContinue'
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path

        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1') -Force -ErrorAction Stop
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Graph\Arraya.M365.Graph.psd1') -Force -ErrorAction Stop
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Exchange\Arraya.M365.Exchange.psd1') -Force -ErrorAction Stop
        Import-Module (Join-Path $script:repoRoot 'src\modules\Arraya.M365.Reporting\Arraya.M365.Reporting.psd1') -Force -ErrorAction Stop

        $script:dictionaryKinds = @('Dictionary', 'ConcurrentDictionary')
        $script:collectorModules = @(
            [pscustomobject]@{ Name = 'Arraya.M365.Graph'; Resolver = 'Graph' }
            [pscustomobject]@{ Name = 'Arraya.M365.Exchange'; Resolver = 'Exchange' }
            [pscustomobject]@{ Name = 'Arraya.M365.Reporting'; Resolver = 'Reporting' }
        )

        function script:New-CompatibilityDictionary {
            param(
                [Parameter(Mandatory = $true)]
                [ValidateSet('Dictionary', 'ConcurrentDictionary')]
                [string]$Kind
            )

            if ($Kind -eq 'ConcurrentDictionary') {
                return [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
            }

            return [System.Collections.Generic.Dictionary[string, object]]::new()
        }

        function script:New-CompatibilityContext {
            param(
                [Parameter(Mandatory = $true)]
                [ValidateSet('Dictionary', 'ConcurrentDictionary')]
                [string]$Kind
            )

            [pscustomobject]@{
                TenantStats        = New-CompatibilityDictionary -Kind $Kind
                Policies           = New-CompatibilityDictionary -Kind $Kind
                Runtime            = New-CompatibilityDictionary -Kind $Kind
                Metadata           = New-CompatibilityDictionary -Kind $Kind
                ExportFileLocation = $null
            }
        }
    }

    AfterAll {
        if ($null -ne $script:previousImportModuleWarningPreference) {
            $PSDefaultParameterValues['Import-Module:WarningAction'] = $script:previousImportModuleWarningPreference
        }
        else {
            $null = $PSDefaultParameterValues.Remove('Import-Module:WarningAction')
        }
    }

    It 'uses generic and concurrent dictionaries for collector cache hits and misses' {
        foreach ($kind in $script:dictionaryKinds) {
            $runtime = New-CompatibilityDictionary -Kind $kind
            $cache = New-CompatibilityDictionary -Kind $kind
            $cacheStats = New-CompatibilityDictionary -Kind $kind
            $cacheStats['Hits'] = 0
            $cacheStats['Misses'] = 0
            $cacheStats['Writes'] = 0
            $runtime['CollectorCache'] = $cache
            $runtime['CollectorCacheStats'] = $cacheStats
            $context = [pscustomobject]@{ Runtime = $runtime }

            Set-ArrayaCollectorCacheValue -Context $context -Key 'present' -Value $kind | Out-Null

            (Get-ArrayaCollectorCacheValue -Context $context -Key 'present') | Should -Be $kind -Because $kind
            (Get-ArrayaCollectorCacheValue -Context $context -Key 'missing') | Should -BeNullOrEmpty -Because $kind
            $context.Runtime['CollectorCacheStats']['Hits'] | Should -Be 1 -Because $kind
            $context.Runtime['CollectorCacheStats']['Misses'] | Should -Be 1 -Because $kind
            $context.Runtime['CollectorCacheStats']['Writes'] | Should -Be 1 -Because $kind
        }
    }

    It 'uses generic and concurrent runtime dictionaries for cached Graph requests' {
        foreach ($kind in $script:dictionaryKinds) {
            $context = New-CompatibilityContext -Kind $kind
            $uri = 'https://example.invalid/dictionary-compatibility/{0}' -f $kind.ToLowerInvariant()

            $first = Invoke-ArrayaGraphCollectionRequest -Uri $uri -Context $context -ScriptBlock { 'collected' }
            $second = Invoke-ArrayaGraphCollectionRequest -Uri $uri -Context $context -ScriptBlock { throw 'The cached result should be used.' }

            $first | Should -Be 'collected' -Because $kind
            $second | Should -Be 'collected' -Because $kind
            $context.Runtime['GraphRequestStats']['Requests'] | Should -Be 2 -Because $kind
            $context.Runtime['GraphRequestStats']['CacheHits'] | Should -Be 1 -Because $kind
            $context.Runtime['GraphRequestStats']['CacheWrites'] | Should -Be 1 -Because $kind
        }
    }

    It 'normalizes generic and concurrent collector contexts in each workload module' {
        foreach ($kind in $script:dictionaryKinds) {
            foreach ($collectorModule in $script:collectorModules) {
                $context = New-CompatibilityContext -Kind $kind
                $module = Get-Module -Name $collectorModule.Name -ErrorAction Stop
                $resolved = & $module {
                    param($resolverName, $candidateContext)

                    switch ($resolverName) {
                        'Graph' { Resolve-ArrayaGraphCollectorContext -Context $candidateContext -DetailLevel minimum }
                        'Exchange' { Resolve-ArrayaExchangeCollectorContext -Context $candidateContext -DetailLevel minimum }
                        'Reporting' { Resolve-ArrayaReportingCollectorContext -Context $candidateContext }
                    }
                } $collectorModule.Resolver $context

                ([System.Collections.IDictionary]$resolved.Metadata).Contains('StartedAt') |
                    Should -BeTrue -Because "$kind in $($collectorModule.Name)"
                $resolved.Runtime.GetType().Name | Should -Be $context.Runtime.GetType().Name
            }
        }
    }

    It 'uses generic and concurrent Exchange policy and runtime dictionaries' {
        foreach ($kind in $script:dictionaryKinds) {
            $context = New-CompatibilityContext -Kind $kind
            $context.Policies['CollectionDepth'] = [pscustomobject]@{
                CollectUnifiedGroupMailboxStats = $true
                IsMinimum                       = $false
            }
            $cachedLookup = [pscustomobject]@{
                Rows                 = 1
                DownloadSucceeded    = $true
                ByGroupId            = @{}
                ByPrimarySmtpAddress = @{}
            }
            $context.Runtime['Office365GroupsActivityMailboxLookup'] = $cachedLookup
            $module = Get-Module -Name 'Arraya.M365.Exchange' -ErrorAction Stop

            $result = & $module {
                param($candidateContext)

                [pscustomobject]@{
                    ShouldCollect = Test-ShouldCollectUnifiedGroupMailboxStats -Context $candidateContext -DetailLevel 'minimum'
                    Lookup        = Get-Office365GroupsActivityMailboxLookup -Context $candidateContext
                }
            } $context

            $result.ShouldCollect | Should -BeTrue -Because $kind
            $result.Lookup | Should -Be $cachedLookup -Because $kind
        }
    }
}
