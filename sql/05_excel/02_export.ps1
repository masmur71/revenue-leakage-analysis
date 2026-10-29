param(
    [string]$Container = 'bintel-pg',
    [string]$User = 'bintel',
    [ValidateSet('bintel','leakage_lab')][string]$Database = 'bintel',
    [string]$PeriodStart = '2025-01-01',
    [string]$PeriodEnd = '2025-12-31',
    [string]$AsOfDate = '2025-12-31'
)
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
foreach ($value in @($PeriodStart,$PeriodEnd,$AsOfDate)) {
    if ($value -cnotmatch '^\d{4}-\d{2}-\d{2}$') { throw 'Dates must be ISO YYYY-MM-DD.' }
    [void][datetime]::ParseExact($value,'yyyy-MM-dd',[cultureinfo]::InvariantCulture)
}
if ($PeriodStart -lt '2024-01-01' -or $PeriodStart -gt $PeriodEnd -or $PeriodEnd -gt $AsOfDate -or $AsOfDate -gt '2025-12-31') {
    throw 'Require 2024-01-01 <= period_start <= period_end <= as_of_date <= 2025-12-31.'
}
function Invoke-Docker {
    & docker @args
    if ($LASTEXITCODE -ne 0) { throw "Docker command failed: $($args[0])" }
}
# Files in /tmp only; no permanent database DDL or source writes.
Invoke-Docker exec $Container mkdir -p /tmp/leakage-phase4/export
Invoke-Docker cp "$root/sql/." "${Container}:/tmp/leakage-phase4/sql"
$params = @('-v',"period_start=$PeriodStart",'-v',"period_end=$PeriodEnd",'-v',"as_of_date=$AsOfDate")
$evidence = Join-Path $root 'docs/excel'
New-Item -ItemType Directory -Force -Path $evidence | Out-Null
Invoke-Docker exec $Container psql -U $User -d $Database -X -q -v ON_ERROR_STOP=1 @params -f /tmp/leakage-phase4/sql/02_kpi/02_validate_metrics.sql *> (Join-Path $evidence 'source-kpi-validation.txt')
Invoke-Docker exec $Container psql -U $User -d $Database -X -q -v ON_ERROR_STOP=1 @params -f /tmp/leakage-phase4/sql/03_analysis/02_validate_analysis.sql *> (Join-Path $evidence 'source-analysis-validation.txt')
Invoke-Docker exec $Container psql -U $User -d $Database -X -q -v ON_ERROR_STOP=1 @params -f /tmp/leakage-phase4/sql/05_excel/00_export.sql *> (Join-Path $evidence 'export-log.txt')
$destination = Join-Path $root 'data/excel'
New-Item -ItemType Directory -Force -Path $destination | Out-Null
Invoke-Docker cp "${Container}:/tmp/leakage-phase4/export/." $destination
$files = @(Get-ChildItem -LiteralPath $destination -Filter '*.csv' | Sort-Object Name | ForEach-Object {
    [ordered]@{file=$_.Name; bytes=$_.Length; sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()}
})
[ordered]@{database=$Database; period_start=$PeriodStart; period_end=$PeriodEnd; as_of_date=$AsOfDate; source_transaction='REPEATABLE READ READ ONLY'; files=$files} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $destination 'manifest.json') -Encoding UTF8
Write-Output "PASS: exported $($files.Count) CSV datasets; native Excel verification remains pending."
