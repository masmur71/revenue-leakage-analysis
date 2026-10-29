param([string]$Container='bintel-pg',[string]$User='bintel')
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$evidence=Join-Path $root 'docs/optimization'
function Invoke-Docker {
    & docker @args
    if ($LASTEXITCODE -ne 0) { throw "Docker operation failed ($($args[0])); inspect evidence output" }
}
function ConvertTo-InvariantRows {
    process {
        $values=[ordered]@{}
        foreach ($property in $_.PSObject.Properties) {
            $value=$property.Value
            if ($value -is [IFormattable]) { $value=$value.ToString($null,[Globalization.CultureInfo]::InvariantCulture) }
            $values[$property.Name]=$value
        }
        [pscustomobject]$values
    }
}
function Invoke-LabSql([string]$Sql) {
    Invoke-Docker exec $Container psql -U $User -d leakage_lab -X -q -At -v ON_ERROR_STOP=1 -c $Sql
}
Invoke-Docker exec $Container mkdir -p /tmp/leakage-phase3
Invoke-Docker cp "$root/sql/." "${Container}:/tmp/leakage-phase3"
$existing=Invoke-LabSql "SELECT count(*) FROM pg_indexes WHERE schemaname='analytics' AND indexname IN ('leakage_opt_call_match','leakage_opt_promo_flag');"
if ([int]$existing -ne 0) { throw 'Experiment index names already exist; preserve existing objects and stop' }
Invoke-LabSql "SELECT json_build_object('database',current_database(),'version',version(),'captured_at',clock_timestamp(),'data','original synthetic snapshot; no generated rows','settings',json_build_object('jit','off','max_parallel_workers_per_gather',0,'work_mem','16MB','repetitions',3),'rows',(SELECT json_object_agg(relname,n_live_tup) FROM pg_stat_user_tables WHERE schemaname='analytics'));" > (Join-Path $evidence 'environment.json')
# Update planner statistics on the copy only, once before all stages.
Invoke-LabSql 'ANALYZE analytics.fact_sales; ANALYZE analytics.fact_returns; ANALYZE analytics.fact_visits; ANALYZE analytics.dim_outlet;'
$helpfulCreated=$false
$unhelpfulCreated=$false
try {
    foreach ($stage in @('baseline','call_index','unhelpful_index')) {
        if ($stage -eq 'call_index') {
            Invoke-LabSql 'CREATE INDEX leakage_opt_call_match ON analytics.fact_sales (salesman_id,sales_date,outlet_id);'
            $helpfulCreated=$true
        }
        if ($stage -eq 'unhelpful_index') {
            Invoke-LabSql 'CREATE INDEX leakage_opt_promo_flag ON analytics.fact_sales (is_promo_valid);'
            $unhelpfulCreated=$true
        }
        Invoke-Docker exec $Container psql -U $User -d leakage_lab -X -q -At -v ON_ERROR_STOP=1 -f /tmp/leakage-phase3/04_optimization/01_benchmark.sql *> (Join-Path $evidence "$stage.jsonl")
        Write-Output "PASS: $stage, exact equality and repeated EXPLAIN captured"
    }
    Invoke-LabSql "SELECT indexrelname,pg_relation_size(indexrelid) bytes,pg_size_pretty(pg_relation_size(indexrelid)) size FROM pg_stat_user_indexes WHERE indexrelname IN ('leakage_opt_call_match','leakage_opt_promo_flag') ORDER BY indexrelname;" > (Join-Path $evidence 'index-sizes.txt')
    $cases=@(
        @{Name='2025';Start='2025-01-01';End='2025-12-31';AsOf='2025-12-31'},
        @{Name='2024';Start='2024-01-01';End='2024-12-31';AsOf='2024-12-31'},
        @{Name='mid2025';Start='2025-01-01';End='2025-06-30';AsOf='2025-06-30'},
        @{Name='partial2024';Start='2024-01-15';End='2024-01-31';AsOf='2024-01-31'},
        @{Name='full';Start='2024-01-01';End='2025-12-31';AsOf='2025-12-31'}
    )
    foreach ($case in $cases) {
        Invoke-Docker exec $Container psql -U $User -d leakage_lab -X -q -v ON_ERROR_STOP=1 -v "period_start=$($case.Start)" -v "period_end=$($case.End)" -v "as_of_date=$($case.AsOf)" -f /tmp/leakage-phase3/04_optimization/02_validate.sql *> (Join-Path $evidence "equality-$($case.Name).txt")
        Write-Output "PASS: $($case.Name) exact multiset equality for three candidates"
    }
}
finally {
    # Remove only indexes created by this invocation; never touch existing indexes.
    if ($unhelpfulCreated) { Invoke-LabSql 'DROP INDEX analytics.leakage_opt_promo_flag;' }
    if ($helpfulCreated) { Invoke-LabSql 'DROP INDEX analytics.leakage_opt_call_match;' }
}
$rows=foreach ($stage in @('baseline','call_index','unhelpful_index')) {
    foreach ($line in Get-Content (Join-Path $evidence "$stage.jsonl")) {
        $item=$line | ConvertFrom-Json
        if ($item.kind -eq 'plan') {
            $plan=$item.explain[0]
            [pscustomobject]@{
                stage=$stage;query=$item.query;variant=$item.variant;repetition=$item.repetition
                execution_ms=$plan.'Execution Time';planning_ms=$plan.'Planning Time'
                shared_hit_blocks=$plan.Plan.'Shared Hit Blocks';shared_read_blocks=$plan.Plan.'Shared Read Blocks'
                temp_read_blocks=$plan.Plan.'Temp Read Blocks';temp_written_blocks=$plan.Plan.'Temp Written Blocks'
                actual_rows=$plan.Plan.'Actual Rows'
            }
        }
    }
}
$rows | ConvertTo-InvariantRows | Export-Csv (Join-Path $evidence 'timings.csv') -NoTypeInformation
$summary=$rows | Group-Object stage,query,variant | ForEach-Object {
    $sorted=@($_.Group.execution_ms | Sort-Object)
    [pscustomobject]@{stage=$_.Group[0].stage;query=$_.Group[0].query;variant=$_.Group[0].variant
        median_ms=$sorted[1];min_ms=$sorted[0];max_ms=$sorted[2]
        median_shared_hits=(@($_.Group.shared_hit_blocks | Sort-Object))[1]
        median_shared_reads=(@($_.Group.shared_read_blocks | Sort-Object))[1]
        median_temp_written=(@($_.Group.temp_written_blocks | Sort-Object))[1]}
}
$summary | ConvertTo-InvariantRows | Export-Csv (Join-Path $evidence 'summary.csv') -NoTypeInformation
$remaining=Invoke-LabSql "SELECT count(*) FROM pg_indexes WHERE indexname IN ('leakage_opt_call_match','leakage_opt_promo_flag');"
if ([int]$remaining -ne 0) { throw 'Experiment indexes were not cleaned up' }
Write-Output 'PASS: summary generated and experiment indexes removed; only copy statistics refreshed'

$remaining | Set-Content (Join-Path $evidence 'cleanup.txt')
Invoke-Docker exec $Container psql -U $User -d leakage_lab -X -q -v ON_ERROR_STOP=1 -f /tmp/leakage-phase3/01_audit/01_audit.sql *> (Join-Path $evidence 'post-audit.txt')
Invoke-Docker exec $Container psql -U $User -d leakage_lab -X -q -v ON_ERROR_STOP=1 -f /tmp/leakage-phase3/04_optimization/02_validate.sql *> (Join-Path $evidence 'final-equality.txt')
# Test guard on postgres, never on the protected source database.
& docker exec $Container psql -U $User -d postgres -X -q -v ON_ERROR_STOP=1 -f /tmp/leakage-phase3/04_optimization/00_candidates.sql *> (Join-Path $evidence 'database-guard.txt')
if ($LASTEXITCODE -ne 3) { throw 'Wrong-database guard must reject execution with psql exit code 3' }
Get-ChildItem (Join-Path $evidence '*.txt') | ForEach-Object {
    $evidencePath=$_.FullName
    $lines=Get-Content -LiteralPath $evidencePath | ForEach-Object { $_.TrimEnd() }
    [IO.File]::WriteAllText($evidencePath,(($lines -join [Environment]::NewLine).TrimEnd()+[Environment]::NewLine))
}
Write-Output 'PASS: post-audit, no-index equality and wrong-database guard'
