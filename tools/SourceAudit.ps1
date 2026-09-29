param(
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceRoot = Join-Path $repoRoot 'src'

if (-not (Test-Path $sourceRoot)) {
    Write-Host '[FAIL] src directory was not found.'
    exit 1
}

$files = @(Get-ChildItem -Path $sourceRoot -Recurse -Filter '*.cs' -File)
if ($files.Count -eq 0) {
    Write-Host '[FAIL] No C# source files were found.'
    exit 1
}

$rules = @(
    @{ Name = 'Unqualified Timer construction'; Pattern = '(?<![\w\.])new\s+Timer\s*\('; Reason = 'Use System.Windows.Forms.Timer explicitly to avoid System.Threading.Timer ambiguity.' },
    @{ Name = 'Unqualified Timer declaration'; Pattern = '(?m)^\s*(?:(?:private|public|protected|internal|static|readonly)\s+)*Timer\s+\w+'; Reason = 'Use System.Windows.Forms.Timer explicitly.' },
    @{ Name = 'C# 8 using declaration'; Pattern = '\busing\s+var\s+'; Reason = 'Framework csc.exe compatibility requires classic using blocks.' },
    @{ Name = 'C# 9 record'; Pattern = '(?m)^\s*(?:public|internal|private|protected)?\s*record\s+'; Reason = 'record syntax is not allowed in the current Framework compiler target.' },
    @{ Name = 'C# 9 init accessor'; Pattern = '\binit\s*;'; Reason = 'init accessors are not allowed in the current Framework compiler target.' },
    @{ Name = 'C# 10 global using'; Pattern = '\bglobal\s+using\s+'; Reason = 'global using is not allowed in the current Framework compiler target.' },
    @{ Name = 'File-scoped namespace'; Pattern = '(?m)^\s*namespace\s+[A-Za-z_][A-Za-z0-9_\.]*\s*;'; Reason = 'Use block namespaces for the current compiler target.' },
    @{ Name = 'Null-coalescing assignment'; Pattern = '\?\?='; Reason = 'Avoid C# 8-only syntax in Framework csc.exe sources.' },
    @{ Name = 'Target-typed new'; Pattern = '=\s*new\s*\('; Reason = 'Use an explicit constructed type for the current compiler target.' },
    @{ Name = 'System.Text.Json'; Pattern = 'System\.Text\.Json'; Reason = 'System.Text.Json is not part of the no-NuGet .NET Framework build references.' },
    @{ Name = 'Span API'; Pattern = '\b(?:ReadOnlySpan|Span)\s*<'; Reason = 'Span-based APIs are outside the current conservative Framework compatibility baseline.' },
    @{ Name = 'DateOnly/TimeOnly API'; Pattern = '\b(?:DateOnly|TimeOnly)\b'; Reason = 'DateOnly/TimeOnly are not available in .NET Framework.' }
)

$failures = New-Object System.Collections.Generic.List[object]

foreach ($file in $files) {
    $text = [System.IO.File]::ReadAllText($file.FullName)

    foreach ($rule in $rules) {
        $matches = [regex]::Matches($text, $rule.Pattern)
        foreach ($match in $matches) {
            $prefix = $text.Substring(0, $match.Index)
            $line = 1 + ([regex]::Matches($prefix, "`n")).Count
            $failures.Add([pscustomobject]@{
                File = $file.FullName.Substring($repoRoot.Length + 1)
                Line = $line
                Rule = $rule.Name
                Reason = $rule.Reason
            })
        }
    }
}

if ($failures.Count -gt 0) {
    Write-Host '[FAIL] Conservative .NET Framework source audit found unsupported/ambiguous patterns:'
    foreach ($failure in $failures) {
        Write-Host ('  {0}:{1}  {2}' -f $failure.File, $failure.Line, $failure.Rule)
        Write-Host ('    ' + $failure.Reason)
    }
    exit 1
}

if (-not $Quiet) {
    Write-Host ('[OK] Source audit passed for {0} C# file(s).' -f $files.Count)
    Write-Host '[OK] No known ambiguous Timer or selected modern-only syntax patterns were found.'
}

exit 0
