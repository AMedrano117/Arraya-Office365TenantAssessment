Describe 'Repository documentation links' {
    BeforeAll {
        $script:repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $script:markdownFiles = @(& git -C $script:repoRoot ls-files -- '*.md')
        if ($LASTEXITCODE -ne 0 -or $script:markdownFiles.Count -eq 0) {
            throw 'Unable to enumerate tracked Markdown files with git.'
        }
    }

    It 'resolves every local Markdown target to an existing file' {
        $missingTargets = @(
            foreach ($relativeMarkdownPath in $script:markdownFiles) {
                $markdownPath = Join-Path $script:repoRoot $relativeMarkdownPath
                if (-not (Test-Path -LiteralPath $markdownPath -PathType Leaf)) {
                    continue
                }
                $content = Get-Content -LiteralPath $markdownPath -Raw
                foreach ($match in [regex]::Matches($content, '!?(?:\[[^\]]*\])\((?<Target>[^)]+)\)')) {
                    $target = [string]$match.Groups['Target'].Value
                    $target = $target.Trim().Trim('<', '>')
                    if (
                        [string]::IsNullOrWhiteSpace($target) -or
                        $target.StartsWith('#') -or
                        $target -match '^[a-z][a-z0-9+.-]*:'
                    ) {
                        continue
                    }

                    $targetPath = ($target -split '[#?]', 2)[0]
                    if ([string]::IsNullOrWhiteSpace($targetPath)) {
                        continue
                    }

                    $targetPath = [System.Uri]::UnescapeDataString($targetPath)
                    $resolvedTarget = [System.IO.Path]::GetFullPath((Join-Path (Split-Path $markdownPath -Parent) $targetPath))
                    if (-not (Test-Path -LiteralPath $resolvedTarget)) {
                        '{0} -> {1}' -f $relativeMarkdownPath, $target
                    }
                }
            }
        )

        $missingTargets | Should -BeNullOrEmpty -Because "local documentation links should not be broken: $($missingTargets -join '; ')"
    }
}
