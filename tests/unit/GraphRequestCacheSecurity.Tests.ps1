Describe 'Graph request cache security' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:commonManifest = Join-Path $script:repoRoot 'src\modules\Arraya.M365.Common\Arraya.M365.Common.psd1'
        Import-Module -Name $script:commonManifest -Force -WarningAction SilentlyContinue -DisableNameChecking -ErrorAction Stop
    }

    It 'does not retain bearer tokens in cache keys or vary the cache by refreshed tokens' {
        $context = New-ArrayaAssessmentContext
        $script:requestCount = 0
        $uri = 'https://graph.microsoft.com/v1.0/users?$top=1'

        $first = Invoke-ArrayaGraphCollectionRequest `
            -Context $context `
            -Uri $uri `
            -GraphMode 'REST' `
            -Headers @{ Authorization = 'Bearer first-sensitive-token'; ConsistencyLevel = 'eventual' } `
            -ScriptBlock {
                $script:requestCount++
                'first response'
            }
        $second = Invoke-ArrayaGraphCollectionRequest `
            -Context $context `
            -Uri $uri `
            -GraphMode 'REST' `
            -Headers @{ authorization = 'Bearer refreshed-sensitive-token'; ConsistencyLevel = 'eventual' } `
            -ScriptBlock {
                $script:requestCount++
                'second response'
            }

        $first | Should -Be 'first response'
        $second | Should -Be 'first response'
        $script:requestCount | Should -Be 1
        $cacheKeys = @($context.Runtime['CollectorCache'].Keys)
        $cacheKeys.Count | Should -Be 1
        ($cacheKeys -join ';') | Should -Not -Match 'first-sensitive-token|refreshed-sensitive-token|Bearer|Authorization'
    }

    It 'still varies the cache by non-sensitive response headers' {
        $context = New-ArrayaAssessmentContext
        $script:requestCount = 0
        $uri = 'https://graph.microsoft.com/v1.0/users?$top=1'

        $html = Invoke-ArrayaGraphCollectionRequest `
            -Context $context `
            -Uri $uri `
            -GraphMode 'REST' `
            -Headers @{ Authorization = 'Bearer sensitive-token'; Prefer = 'outlook.body-content-type="html"' } `
            -ScriptBlock {
                $script:requestCount++
                'html response'
            }
        $text = Invoke-ArrayaGraphCollectionRequest `
            -Context $context `
            -Uri $uri `
            -GraphMode 'REST' `
            -Headers @{ Authorization = 'Bearer sensitive-token'; Prefer = 'outlook.body-content-type="text"' } `
            -ScriptBlock {
                $script:requestCount++
                'text response'
            }

        $html | Should -Be 'html response'
        $text | Should -Be 'text response'
        $script:requestCount | Should -Be 2
        @($context.Runtime['CollectorCache'].Keys).Count | Should -Be 2
    }

    AfterAll {
        Remove-Module -Name 'Arraya.M365.Common' -Force -ErrorAction SilentlyContinue
    }
}
