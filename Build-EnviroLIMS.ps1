<#
  Builds EnviroLIMS.accdb next to this script using Microsoft Access.

  Requirements: Windows + Microsoft Access (Microsoft 365, 2016, 2019, 2021 or 2024).
  Usage:        Double-click Build-EnviroLIMS.cmd, or run:
                powershell -ExecutionPolicy Bypass -File Build-EnviroLIMS.ps1 [-NoDemoData] [-Output path]
#>
param(
    [switch]$NoDemoData,
    [string]$Output
)

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Output) { $Output = Join-Path $here 'EnviroLIMS.accdb' }
$srcDir = Join-Path $here 'src'

if (Test-Path $Output) {
    $answer = Read-Host "$Output already exists. Overwrite it? ALL DATA IN IT WILL BE LOST (y/N)"
    if ($answer -notmatch '^[Yy]') { Write-Host 'Cancelled.'; exit 1 }
    Remove-Item $Output -Force
}

Write-Host 'Starting Microsoft Access...'
$access = New-Object -ComObject Access.Application
try {
    $access.AutomationSecurity = 1      # enable macros/VBA in the database we create
    $access.Visible = $true
    $access.NewCurrentDatabase($Output)

    # Import the VBA modules (Access does not need "Trust access to the VBA project")
    $modules = Get-ChildItem $srcDir -Filter '*.bas' | Sort-Object Name
    $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('EnviroLIMS_' + [guid]::NewGuid())
    New-Item -ItemType Directory -Path $tmpDir | Out-Null
    foreach ($m in $modules) {
        Write-Host "Importing $($m.Name)"
        # The VBA editor needs CRLF line endings; normalise in case the files were downloaded with LF.
        $text = [IO.File]::ReadAllText($m.FullName) -replace "`r?`n", "`r`n"
        $tmpFile = Join-Path $tmpDir $m.Name
        [IO.File]::WriteAllText($tmpFile, $text, [Text.Encoding]::ASCII)
        $null = $access.VBE.ActiveVBProject.VBComponents.Import($tmpFile)
        $access.DoCmd.Save(5, $m.BaseName)  # 5 = acModule
    }
    Remove-Item $tmpDir -Recurse -Force

    Write-Host 'Building tables, queries, forms and reports (this takes a minute)...'
    $access.Run('BuildLIMS', (-not $NoDemoData.IsPresent))

    Write-Host "Done. Database created at $Output"
    $access.UserControl = $true         # leave Access open for the user
}
catch {
    Write-Host "Build failed: $($_.Exception.Message)" -ForegroundColor Red
    try { $access.Quit() } catch {}
    exit 1
}
