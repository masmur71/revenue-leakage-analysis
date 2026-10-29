param(
    [string]$Container = 'bintel-pg',
    [string]$User = 'bintel',
    [ValidatePattern('^leakage_qa_[a-z0-9_]+$')][string]$Database = 'leakage_qa_phase6',
    [string]$Source = 'bintel'
)
# Restore the bundled snapshot, validate unchanged reference SQL, and reproduce Excel inputs.
# Source is SELECT-only; destination must be new. No database is ever dropped.
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$out = Join-Path $root 'docs/qa'
$work = Join-Path $root 'data/phase6-qa'
if (-not (Test-Path (Join-Path $root 'data/excel/manifest.json'))) { throw 'Generate and validate default Excel exports first.' }
New-Item -ItemType Directory -Force $out,$work | Out-Null
function Invoke-Docker {
    $result = & docker @args 2>&1
    if ($LASTEXITCODE -ne 0) { throw ($result -join "`n") }
    return $result
}
if ($Source -eq $Database) { throw 'QA destination must differ from the source.' }
$exists = Invoke-Docker exec $Container psql -X -At -U $User -d postgres -c "SELECT count(*) FROM pg_database WHERE datname='$Database'"
if ("$exists".Trim() -ne '0') { throw 'QA database already exists; choose a new leakage_qa_* name. Never overwrite.' }
Invoke-Docker cp (Join-Path $root 'data/analytics.dump') "${Container}:/tmp/leakage-phase6.dump" | Out-Null
Invoke-Docker exec $Container createdb -U $User -T template0 $Database | Out-Null
Invoke-Docker exec $Container pg_restore -U $User -d $Database --no-owner --no-privileges --exit-on-error /tmp/leakage-phase6.dump | Out-Null
Invoke-Docker exec $Container mkdir -p /tmp/leakage-phase6/export | Out-Null
Invoke-Docker cp (Join-Path $root 'sql/.') "${Container}:/tmp/leakage-phase6/sql" | Out-Null
# Export SQL has a fixed output path. Change ONLY its scratch copy, preserving query logic.
foreach ($file in Get-ChildItem (Join-Path $root 'sql/05_excel') -Filter '*.sql') {
    $scratch = Join-Path $work $file.Name
    (Get-Content $file.FullName -Raw).Replace('/tmp/leakage-phase4/export/','/tmp/leakage-phase6/export/') | Set-Content $scratch -Encoding UTF8
    Invoke-Docker cp $scratch "${Container}:/tmp/leakage-phase6/sql/05_excel/$($file.Name)" | Out-Null
}
foreach ($db in @($Source,$Database)) {
    $label = if ($db -eq $Source) {'source'} else {'restored'}
    Invoke-Docker exec $Container psql -X -q -U $User -d $db -v ON_ERROR_STOP=1 -f /tmp/leakage-phase6/sql/01_audit/01_audit.sql | Set-Content (Join-Path $out "audit-$label.txt")
}
if (Compare-Object (Get-Content "$out/audit-source.txt") (Get-Content "$out/audit-restored.txt")) { throw 'Source/restored audit differs.' }
$cases = @(
    @('2025','2025-01-01','2025-12-31','2025-12-31'),
    @('2024','2024-01-01','2024-12-31','2024-12-31'),
    @('mid2025','2025-01-01','2025-06-30','2025-06-30'),
    @('partial2024','2024-01-15','2024-01-31','2024-01-31'),
    @('full','2024-01-01','2025-12-31','2025-12-31')
)
foreach ($case in $cases) {
    foreach ($kind in @('02_kpi/02_validate_metrics.sql','03_analysis/02_validate_analysis.sql')) {
        $stem = if ($kind.StartsWith('02')) {'kpi'} else {'analysis'}
        foreach ($db in @($Source,$Database)) {
            $label = if ($db -eq $Source) {'source'} else {'restored'}
            Invoke-Docker exec $Container psql -X -q -U $User -d $db -v ON_ERROR_STOP=1 -v "period_start=$($case[1])" -v "period_end=$($case[2])" -v "as_of_date=$($case[3])" -f "/tmp/leakage-phase6/sql/$kind" | Set-Content (Join-Path $out "$stem-$($case[0])-$label.txt")
        }
        if (Compare-Object (Get-Content "$out/$stem-$($case[0])-source.txt") (Get-Content "$out/$stem-$($case[0])-restored.txt")) { throw "$stem differs in $($case[0])" }
    }
}
Invoke-Docker exec $Container psql -X -q -U $User -d $Database -v ON_ERROR_STOP=1 -f /tmp/leakage-phase6/sql/05_excel/00_export.sql | Set-Content "$out/export.txt"
Invoke-Docker cp "${Container}:/tmp/leakage-phase6/export/." $work | Out-Null
$manifest = Get-Content (Join-Path $root 'data/excel/manifest.json') -Raw | ConvertFrom-Json
$checks = foreach ($file in $manifest.files) {
    $fresh = Join-Path $work $file.file
    $hash = (Get-FileHash -LiteralPath $fresh -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($hash -ne $file.sha256) { throw "Reproduced export differs: $($file.file)" }
    [pscustomobject]@{file=$file.file;sha256=$hash;matches_accepted_source=$true}
}
$checks | Export-Csv (Join-Path $out 'export-hashes.csv') -NoTypeInformation
Invoke-Docker exec $Container psql -X -q -U $User -d $Database -v ON_ERROR_STOP=1 -f /tmp/leakage-phase6/sql/03_analysis/04_recommendation_evidence.sql | Set-Content "$out/recommendations.txt"
@{
    status='PASS'; database=$Database; source=$Source; date=[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow,'SE Asia Standard Time').ToString('yyyy-MM-dd'); timezone='Asia/Jakarta';
    audit='13 table counts and 40 controls; source/restored output identical';
    scenarios=$cases.Count; suites='KPI and analysis; source/restored output identical';
    export_files=$checks.Count; exports='SHA-256 equal to accepted Excel source manifest';
    dump_sha256=(Get-FileHash (Join-Path $root 'data/analytics.dump') -Algorithm SHA256).Hash.ToLowerInvariant();
    workbook_sha256=(Get-FileHash (Join-Path $root 'excel/RevenueLeakage.xlsx') -Algorithm SHA256).Hash.ToLowerInvariant();
    limits='No new native refresh or UI check; use separately rechecked Phase 4 captures. QA database retained; no cleanup/drop.'
} | ConvertTo-Json | Set-Content "$out/reproduction.json"
# Keep psql evidence readable and whitespace-clean for Git without changing values.
Get-ChildItem $out -Filter '*.txt' | ForEach-Object {
    $lines = Get-Content $_.FullName | ForEach-Object { $_.TrimEnd() }
    [IO.File]::WriteAllText($_.FullName,($lines -join [Environment]::NewLine).TrimEnd()+[Environment]::NewLine)
}
Write-Output "PASS: new $Database restored, five source/copy scenarios match, $($checks.Count) exports match accepted manifest."
