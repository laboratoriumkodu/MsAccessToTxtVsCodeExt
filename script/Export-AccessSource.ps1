<#
.SYNOPSIS
    Exports VBA modules, forms, reports, macros, queries and table schema from an
    Access .accdb/.mdb file into plain-text files so they can be tracked in Git.

.DESCRIPTION
    Access COM automation only exposes 'SaveAsText' for modules/forms/reports/macros.
    Table data and compiled layout internals stay binary - this captures everything
    that is meaningfully diffable: VBA source, form/report definitions (incl. code-behind),
    macro XML, query SQL, and table/field/relationship DDL-style metadata.

.PARAMETER AccdbPath
    Full path to the .accdb/.mdb file to export from.

.PARAMETER OutputPath
    Folder to write the exported text files into. Created if it doesn't exist.
    Existing contents are left in place (files are overwritten on re-export).

.EXAMPLE
    .\Export-AccessSource.ps1 -AccdbPath C:\source.ACCDB -OutputPath C:\destination\AccessSource
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AccdbPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$AccdbPath = (Resolve-Path -Path $AccdbPath).Path

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
$OutputPath = (Resolve-Path -Path $OutputPath).Path

$acModule = 5
$acForm = 2
$acReport = 3
$acMacro = 4

function New-SubFolder {
    param([string]$Base, [string]$Name)
    $p = Join-Path $Base $Name
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
    return $p
}

function Remove-ComReference {
    param([object]$Obj)
    if ($null -ne $Obj -and [Runtime.InteropServices.Marshal]::IsComObject($Obj)) {
        while ([Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) -gt 0) {}
    }
}

function Get-SafeFileName {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars() -join ''
    $pattern = "[{0}]" -f [Regex]::Escape($invalid)
    return [Regex]::Replace($Name, $pattern, '_')
}

$modulesDir = New-SubFolder $OutputPath 'Modules'
$formsDir = New-SubFolder $OutputPath 'Forms'
$reportsDir = New-SubFolder $OutputPath 'Reports'
$macrosDir = New-SubFolder $OutputPath 'Macros'
$queriesDir = New-SubFolder $OutputPath 'Queries'
$schemaDir = New-SubFolder $OutputPath 'Schema'

Write-Host "Opening Access database: $AccdbPath"
$access = New-Object -ComObject Access.Application
$access.Visible = $false
Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition '[DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(int hWnd, out int processId);'

try {
    $access.OpenCurrentDatabase($AccdbPath, $false)
    Start-Sleep -Seconds 5 # give Access time to finish loading/compiling before automating it
    $db = $access.CurrentDb()

    Write-Host "Exporting VBA modules..."
    foreach ($m in $access.CurrentProject.AllModules) {
        $name = Get-SafeFileName $m.Name
        $file = Join-Path $modulesDir "$name.bas"
        $access.SaveAsText($acModule, $m.Name, $file)
    }

    Write-Host "Exporting forms..."
    foreach ($f in $access.CurrentProject.AllForms) {
        $name = Get-SafeFileName $f.Name
        $file = Join-Path $formsDir "$name.txt"
        $access.SaveAsText($acForm, $f.Name, $file)
    }

    Write-Host "Exporting reports..."
    foreach ($r in $access.CurrentProject.AllReports) {
        $name = Get-SafeFileName $r.Name
        $file = Join-Path $reportsDir "$name.txt"
        $access.SaveAsText($acReport, $r.Name, $file)
    }

    Write-Host "Exporting macros..."
    foreach ($mc in $access.CurrentProject.AllMacros) {
        $name = Get-SafeFileName $mc.Name
        $file = Join-Path $macrosDir "$name.txt"
        $access.SaveAsText($acMacro, $mc.Name, $file)
    }

    Write-Host "Exporting query SQL..."
    foreach ($q in $db.QueryDefs) {
        if ($q.Name.StartsWith('~')) { continue } # skip hidden/temp query defs
        $name = Get-SafeFileName $q.Name
        $file = Join-Path $queriesDir "$name.sql"
        Set-Content -Path $file -Value $q.SQL -Encoding UTF8
    }

    Write-Host "Exporting table schema..."
    foreach ($t in $db.TableDefs) {
        if ($t.Name.StartsWith('MSys') -or $t.Name.StartsWith('~')) { continue }
        $name = Get-SafeFileName $t.Name
        $file = Join-Path $schemaDir "$name.txt"
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("Table: $($t.Name)")
        $lines.Add('')
        $lines.Add('Fields:')
        foreach ($fld in $t.Fields) {
            $lines.Add("  $($fld.Name)`t$($fld.Type)`tSize=$($fld.Size)`tRequired=$($fld.Required)")
        }
        $lines.Add('')
        $lines.Add('Indexes:')
        foreach ($idx in $t.Indexes) {
            $cols = ($idx.Fields | ForEach-Object { $_.Name }) -join ', '
            $lines.Add("  $($idx.Name)`tPrimary=$($idx.Primary)`tUnique=$($idx.Unique)`tFields=$cols")
        }
        Set-Content -Path $file -Value $lines -Encoding UTF8
    }

    Write-Host "Exporting relationships..."
    $relFile = Join-Path $schemaDir '_Relationships.txt'
    $relLines = New-Object System.Collections.Generic.List[string]
    foreach ($rel in $db.Relations) {
        $fields = ($rel.Fields | ForEach-Object { "$($_.Name)->$($_.ForeignName)" }) -join ', '
        $relLines.Add("$($rel.Name): $($rel.Table) -> $($rel.ForeignTable) [$fields] Attributes=$($rel.Attributes)")
    }
    Set-Content -Path $relFile -Value $relLines -Encoding UTF8

    Write-Host "Export complete: $OutputPath"
}
finally {
    # Resolve the real MSACCESS process id up front - Quit() can return before the
    # process actually exits (e.g. if it's stuck on a hidden dialog it can't show).
    $procId = 0
    [Win32.NativeMethods]::GetWindowThreadProcessId($access.hWndAccessApp(), [ref]$procId) | Out-Null

    # Every DAO/Access object still referenced from PowerShell keeps MSACCESS alive after Quit().
    foreach ($v in 'm', 'f', 'r', 'mc', 'q', 't', 'fld', 'idx', 'rel') {
        Remove-ComReference (Get-Variable -Name $v -ValueOnly -ErrorAction SilentlyContinue)
        Remove-Variable -Name $v -ErrorAction SilentlyContinue
    }
    if ($db) {
        try { $db.Close() } catch {}
        Remove-ComReference $db
        $db = $null
    }

    try { $access.CloseCurrentDatabase() } catch {}
    try { $access.Quit(2) } catch {} # acQuitSaveNone - never prompt/save on exit
    Remove-ComReference $access
    $access = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($procId -gt 0) {
        Wait-Process -Id $procId -Timeout 20 -ErrorAction SilentlyContinue
        if (Get-Process -Id $procId -ErrorAction SilentlyContinue) {
            Write-Warning "Access process (PID $procId) did not exit gracefully; forcing it to close."
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
            Wait-Process -Id $procId -Timeout 10 -ErrorAction SilentlyContinue
        }
    }

    $ext = [IO.Path]::GetExtension($AccdbPath).TrimStart('.').ToLower()
    $lockExt = if ($ext -eq 'mdb') { 'ldb' } else { "l$ext" }
    $lockFile = [IO.Path]::ChangeExtension($AccdbPath, $lockExt)
    # The lock file handle can linger briefly after the process exits.
    for ($i = 0; ($i -lt 10) -and (Test-Path $lockFile); $i++) {
        Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
        if (Test-Path $lockFile) { Start-Sleep -Milliseconds 500 }
    }
    if (Test-Path $lockFile) {
        Write-Warning "Lock file '$lockFile' is still in use (is the database open in another Access instance?)."
    }
}
