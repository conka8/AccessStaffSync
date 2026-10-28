#requires -Version 5.1
<#
.SYNOPSIS
  Copies simulated WhatsApp updates, runs two imports via Access COM, checks example outcomes.
.DESCRIPTION
  Create StaffSyncDemo.accdb first with Install-Demo.ps1.
  The demo DB must not be open in another Access instance during this run.
#>
[CmdletBinding()]
param([string]$Destination = (Join-Path $PSScriptRoot '..\demo'))
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Destination = [IO.Path]::GetFullPath($Destination)
$dbPath = Join-Path $Destination 'StaffSyncDemo.accdb'
if (!(Test-Path $dbPath)) { throw 'First run setup\Install-Demo.ps1' }
$inbox = Join-Path $Destination 'inbox'
$app = $null
try {
    $app = New-Object -ComObject Access.Application
    $app.OpenCurrentDatabase($dbPath)
    $expected = @{
        'updates_001.xlsx' = @{ TotalRows=6; AppliedRows=4; StaleRows=0; DuplicateRows=1; RejectedRows=1 }
        'updates_002.xlsx' = @{ TotalRows=7; AppliedRows=3; StaleRows=1; DuplicateRows=1; RejectedRows=2 }
    }
    foreach ($name in @('updates_001.xlsx', 'updates_002.xlsx')) {
        Copy-Item (Join-Path $root ('samples\inbox\' + $name)) (Join-Path $inbox $name) -ErrorAction Stop
        Write-Host "`nImporting: $name" -ForegroundColor Cyan
        [void]$app.Run('SyncTick')
        $db = $app.CurrentDb()
        $rs = $db.OpenRecordset('SELECT [FileKey],[Result],[TotalRows],[AppliedRows],[StaleRows],[DuplicateRows],[RejectedRows] FROM [tblImportFiles] WHERE [FileKey] = ''' + $name + '''')
        if ($rs.EOF) { throw "No completed receipt for $name (inspect staffsync.log)." }
        Write-Host ("{0} | result={1} | total={2} applied={3} stale={4} duplicate={5} rejected={6}" -f `
            $rs.Fields('FileKey').Value, $rs.Fields('Result').Value, $rs.Fields('TotalRows').Value, `
            $rs.Fields('AppliedRows').Value, $rs.Fields('StaleRows').Value, `
            $rs.Fields('DuplicateRows').Value, $rs.Fields('RejectedRows').Value)
        foreach ($field in $expected[$name].Keys) {
            if ([int]$rs.Fields($field).Value -ne $expected[$name][$field]) {
                throw "Assertion failed for $name $field : expected $($expected[$name][$field]), got $($rs.Fields($field).Value)"
            }
        }
        $rs.Close()
        $receipt = Join-Path (Join-Path $Destination 'receipts') ($name + '.receipt.txt')
        if (!(Test-Path $receipt)) { throw "Receipt missing: $receipt" }
    }
    $rs = $app.CurrentDb().OpenRecordset("SELECT [StaffID],[JobOfferStatus],[Availability],[Address],[MessagingConsent] FROM [tblStaff] WHERE [StaffID] IN ('000042','000314','001234') ORDER BY [StaffID]")
    Write-Host "`nFinal staff records:" -ForegroundColor Cyan
    while (!$rs.EOF) {
        Write-Host ("{0} status={1} avail={2} address={3} consent={4}" -f $rs.Fields('StaffID').Value, `
            $rs.Fields('JobOfferStatus').Value, $rs.Fields('Availability').Value, `
            $rs.Fields('Address').Value, $rs.Fields('MessagingConsent').Value)
        $rs.MoveNext()
    }
    $rs.Close()
    $numEvents = [int]$app.Run('CountImportedEvents')
    if ($numEvents -ne 8) { throw "Expected 8 unique committed events, found $numEvents" }
    Write-Host "`nPASS: counts, receipts, and 8 unique committed events." -ForegroundColor Green
    Write-Host "Receipts: $(Join-Path $Destination 'receipts')" -ForegroundColor Green
    Write-Host 'Re-open the Access demo to see the populated audit tables.'
} finally {
    if ($null -ne $app) {
        try { $app.CloseCurrentDatabase() } catch {}
        try { $app.Quit(2) } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}
