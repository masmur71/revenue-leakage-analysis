param(
    [string]$Container = 'bintel-pg',
    [string]$SourceDatabase = 'bintel',
    [string]$User = 'bintel',
    [string]$CopyDatabase = 'leakage_lab'
)
$ErrorActionPreference = 'Stop'
if ($CopyDatabase -eq $SourceDatabase -or $CopyDatabase -notmatch '^[a-z][a-z0-9_]*$') { throw 'Use a distinct, simple copy database name.' }
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
function Invoke-Docker {
    & docker @args
    if ($LASTEXITCODE -ne 0) { throw "Docker command failed: $($args[0])" }
}
$exists = Invoke-Docker exec $Container psql -U $User -d $SourceDatabase -X -Atc "SELECT count(*) FROM pg_database WHERE datname='$CopyDatabase';"
if ($exists -ne '0') { throw "Database $CopyDatabase already exists. Choose a new name; this script never drops databases." }
Invoke-Docker exec $Container pg_dump -U $User -d $SourceDatabase -n analytics --no-owner --no-privileges -Fc -f /tmp/leakage-analytics.dump
Invoke-Docker exec $Container pg_dump -U $User -d $SourceDatabase -n analytics --no-owner --no-privileges --schema-only -f /tmp/leakage-schema.sql
Invoke-Docker cp "${Container}:/tmp/leakage-analytics.dump" (Join-Path $root 'data/analytics.dump')
Invoke-Docker cp "${Container}:/tmp/leakage-schema.sql" (Join-Path $PSScriptRoot 'analytics_schema.sql')
Invoke-Docker exec $Container createdb -U $User -T template0 $CopyDatabase
Invoke-Docker exec $Container pg_restore -U $User -d $CopyDatabase --no-owner --no-privileges --exit-on-error /tmp/leakage-analytics.dump
Write-Output "Restored analytics into $CopyDatabase. Source database was not modified."
