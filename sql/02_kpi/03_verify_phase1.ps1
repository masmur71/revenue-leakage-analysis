param([string]$Container='bintel-pg',[string]$User='bintel',[string]$Source='bintel',[string]$Copy='leakage_lab')
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
function Invoke-Docker {
    & docker @args
    if ($LASTEXITCODE -ne 0) { throw "Docker command failed: $($args[0])" }
}
Invoke-Docker exec $Container mkdir -p /tmp/leakage-kpi
Invoke-Docker cp "$PSScriptRoot/." "${Container}:/tmp/leakage-kpi"
$cases=@(
    @{Name='2025';Start='2025-01-01';End='2025-12-31';AsOf='2025-12-31'},
    @{Name='2024';Start='2024-01-01';End='2024-12-31';AsOf='2024-12-31'},
    @{Name='mid2025';Start='2025-01-01';End='2025-06-30';AsOf='2025-06-30'},
    @{Name='partial2024';Start='2024-01-15';End='2024-01-31';AsOf='2024-01-31'},
    @{Name='full';Start='2024-01-01';End='2025-12-31';AsOf='2025-12-31'}
)
foreach ($case in $cases) {
    $output=Join-Path $root "docs/kpi-validation-$($case.Name).txt"
    Invoke-Docker exec $Container psql -U $User -d $Source -X -v ON_ERROR_STOP=1 `
        -v "period_start=$($case.Start)" -v "period_end=$($case.End)" -v "as_of_date=$($case.AsOf)" `
        -f /tmp/leakage-kpi/02_validate_metrics.sql *> $output
    Write-Output "PASS: $($case.Name) contracts"
}
$sourceOutput=Join-Path $root 'docs/kpi-reference-2025.txt'
$copyOutput=Join-Path $root 'docs/kpi-reference-restored-2025.txt'
Invoke-Docker exec $Container psql -U $User -d $Source -X -v ON_ERROR_STOP=1 -f /tmp/leakage-kpi/01_reference_metrics.sql > $sourceOutput
Invoke-Docker exec $Container psql -U $User -d $Copy -X -v ON_ERROR_STOP=1 -f /tmp/leakage-kpi/01_reference_metrics.sql > $copyOutput
if (Compare-Object (Get-Content $sourceOutput) (Get-Content $copyOutput)) { throw 'Source/copy metric outputs differ' }
Write-Output 'PASS: source/copy reference metrics, RFM and cohorts match'
$invalidCases=@(
 @{Name='future cutoff';Start='2025-01-01';End='2025-12-31';AsOf='2026-01-01'},
 @{Name='period after snapshot';Start='2025-01-01';End='2025-12-31';AsOf='2025-06-30'},
 @{Name='reversed period';Start='2025-12-31';End='2025-01-01';AsOf='2025-12-31'},
 @{Name='before source coverage';Start='2023-12-31';End='2025-12-31';AsOf='2025-12-31'}
)
foreach ($case in $invalidCases) {
    $ErrorActionPreference='Continue'
    $result=& docker exec $Container psql -U $User -d $Source -X -v ON_ERROR_STOP=1 `
        -v "period_start=$($case.Start)" -v "period_end=$($case.End)" -v "as_of_date=$($case.AsOf)" `
        -f /tmp/leakage-kpi/01_reference_metrics.sql 2>&1
    $exitCode=$LASTEXITCODE
    $ErrorActionPreference='Stop'
    if ($exitCode -eq 0 -or ($result -join "`n") -notmatch 'Require 2024-01-01') { throw "Expected date validation rejection: $($case.Name)" }
    Write-Output "PASS: rejected $($case.Name)"
}
