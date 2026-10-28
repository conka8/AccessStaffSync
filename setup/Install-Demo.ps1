#requires -Version 5.1
<#
.SYNOPSIS
  Creates a disposable, isolated Access .accdb and loads the VBA automation.
.DESCRIPTION
  Requires desktop Microsoft Access and Excel on Windows.
  Importing a .bas through VBIDE requires temporary enablement of Access Trust Center:
    Trust access to the VBA project object model.
  Do NOT run on the client's operational database.
#>
[CmdletBinding()]
param(
    [string]$Destination = (Join-Path $PSScriptRoot '..\demo'),
    [switch]$StartHidden
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$Destination = [IO.Path]::GetFullPath($Destination)
$database = Join-Path $Destination 'StaffSyncDemo.accdb'
$modulePath = Join-Path $root 'vba\modStaffSync.bas'
if (Test-Path $database) { throw "Demo DB exists: $database. Refusing to overwrite it." }
foreach ($sub in @('', 'inbox', 'archive', 'receipts')) {
    $target = if ($sub -eq '') { $Destination } else { Join-Path $Destination $sub }
    New-Item -ItemType Directory -Path $target -Force | Out-Null
}
$app = $null
try {
    $app = New-Object -ComObject Access.Application
    $app.NewCurrentDatabase($database)
    try {
        $module = $app.VBE.ActiveVBProject.VBComponents.Import($modulePath)
    } catch {
        throw "VBA import blocked. In Access, enable File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model' temporarily; then remove the unfinished demo DB and rerun. Original error: $_"
    }
    $app.Run('InitializeDemo') | Out-Null
    $app.Run('SeedDemoStaff', 4000) | Out-Null
    $app.Run('PutSetting', 'InboxPath', (Join-Path $Destination 'inbox')) | Out-Null
    $app.Run('PutSetting', 'ArchivePath', (Join-Path $Destination 'archive')) | Out-Null
    $app.Run('PutSetting', 'ReceiptPath', (Join-Path $Destination 'receipts')) | Out-Null
    $app.Run('PutSetting', 'ReadyDelaySeconds', '0') | Out-Null # demo only

    # A form's expression-based Timer event does not require a form-code module.
    $form = $app.CreateForm()
    $temporaryName = $form.Name
    $form.Caption = 'AccessStaffSync - 60-second background sync'
    $form.RecordSelectors = $false
    $form.NavigationButtons = $false
    $form.DividingLines = $false
    $form.ScrollBars = 0
    $form.TimerInterval = 60000
    $form.OnTimer = '=SyncTick()'
    if ($StartHidden) { $form.OnLoad = '=HideSyncForm()' }
    # acLabel = 100; acDetail = 0; positioned in twips.
    $label = $app.CreateControl($temporaryName, 100, 0, '', '', 420, 480, 9000, 620)
    $label.Caption = 'StaffSync | Automated Access imports'
    $label.FontSize = 18
    $label.FontBold = $true
    $help = $app.CreateControl($temporaryName, 100, 0, '', '', 420, 1320, 10500, 850)
    $help.Caption = 'Checks the inbox every 60 seconds. Open tblImportEvents / tblImportFiles and the receipts folder to review results.'
    $help.FontSize = 11
    $form.Section(0).Height = 2600
    $app.DoCmd.Save(2, $temporaryName)  # acForm=2
    $app.DoCmd.Close(2, $temporaryName, 1) # acSaveYes=1
    $app.DoCmd.Rename('frmSyncTimer', 2, $temporaryName)

    # StartupForm (not an AutoExec macro) keeps the Timer form loaded.
    $db = $app.CurrentDb()
    try {
        $db.Properties.Item('StartupForm').Value = 'frmSyncTimer'
    } catch {
        $prop = $db.CreateProperty('StartupForm', 10, 'frmSyncTimer') # dbText=10
        $db.Properties.Append($prop)
    }
    $app.CloseCurrentDatabase()
    Write-Host "Created demonstration database: $database" -ForegroundColor Green
    Write-Host 'Demo-only ReadyDelaySeconds=0. For production use 20 or higher.'
    Write-Host 'Open the .accdb in desktop Access to run its 60-second timer form.'
    Write-Host 'For a scripted walkthrough run: .\setup\Run-Demo.ps1'
} finally {
    if ($null -ne $app) {
        try { $app.Quit(2) } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}
