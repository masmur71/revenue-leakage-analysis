param([string]$Container='bintel-pg',[string]$User='bintel',[string]$Source='bintel',[string]$Copy='leakage_lab')
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
function Invoke-Docker {
    & docker @args
    if ($LASTEXITCODE -ne 0) { throw "Docker command failed: $($args[0])" }
}
Invoke-Docker exec $Container mkdir -p /tmp/leakage-phase2
Invoke-Docker cp "$root/sql/." "${Container}:/tmp/leakage-phase2"
$cases=@(
 @{Name='2025';Start='2025-01-01';End='2025-12-31';AsOf='2025-12-31'},
 @{Name='2024';Start='2024-01-01';End='2024-12-31';AsOf='2024-12-31'},
 @{Name='mid2025';Start='2025-01-01';End='2025-06-30';AsOf='2025-06-30'},
 @{Name='partial2024';Start='2024-01-15';End='2024-01-31';AsOf='2024-01-31'},
 @{Name='full';Start='2024-01-01';End='2025-12-31';AsOf='2025-12-31'}
)
foreach ($case in $cases) {
    $params=@('-v',"period_start=$($case.Start)",'-v',"period_end=$($case.End)",'-v',"as_of_date=$($case.AsOf)")
    foreach ($db in @($Source,$Copy)) {
        $suffix=if ($db -eq $Source) {'source'} else {'copy'}
        $output=Join-Path $root "docs/analysis-validation-$($case.Name)-$suffix.txt"
        Invoke-Docker exec $Container psql -U $User -d $db -X -q -v ON_ERROR_STOP=1 @params -f /tmp/leakage-phase2/03_analysis/02_validate_analysis.sql *> $output
    }
    $sourceLines=Get-Content (Join-Path $root "docs/analysis-validation-$($case.Name)-source.txt")
    $copyLines=Get-Content (Join-Path $root "docs/analysis-validation-$($case.Name)-copy.txt")
    if (Compare-Object $sourceLines $copyLines) { throw "Source/copy diagnostics differ: $($case.Name)" }
    Write-Output "PASS: $($case.Name), diagnostic checks and eight full-row fingerprints match"
}
Invoke-Docker exec $Container psql -U $User -d $Source -X -q -v ON_ERROR_STOP=1 -f /tmp/leakage-phase2/03_analysis/01_business_report.sql *> (Join-Path $root 'docs/analysis-report-2025.txt')
Write-Output 'PASS: default business report generated in read-only transaction'
# Normalize table-padding whitespace in generated text evidence for clean diffs.
Get-ChildItem (Join-Path $root 'docs/analysis-*.txt') | ForEach-Object {
    $evidencePath=$_.FullName
    $lines=Get-Content -LiteralPath $evidencePath | ForEach-Object { $_.TrimEnd() }
    [IO.File]::WriteAllText($evidencePath,(($lines -join [Environment]::NewLine).TrimEnd()+[Environment]::NewLine))
}
