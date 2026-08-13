#requires -Version 5.1

<#
.SYNOPSIS
Converts legacy .xls workbooks to .xlsx or .xlsm by using desktop Excel.

.DESCRIPTION
Inspects legacy .xls workbooks without running their macros. Workbooks that
contain a VBA project are converted to .xlsm; all others are converted to
.xlsx. The source files are never deleted or overwritten, and existing output
files are skipped.

The default mode is a dry run. Add -Write to create the converted workbooks.
Progress is written to stderr and one summary-only JSON result is written to
stdout. The JSON does not repeat the per-workbook results.

.PARAMETER SourcePath
Source folder containing the legacy .xls workbooks to convert. Defaults to the
current folder. This is not the output destination.

.PARAMETER OutputDirectory
Optional destination folder. When omitted, each converted workbook is saved
next to its source .xls workbook.

.PARAMETER Recurse
Includes .xls workbooks in subfolders. With the default output behavior, each
converted workbook is saved next to its source. When OutputDirectory is set,
the output is flat, so duplicate output names are reported as collisions.

.PARAMETER Write
Creates the converted workbooks. Without this switch, only the plan is shown.

.EXAMPLE
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath 'C:\path\to\legacy-workbooks'

.EXAMPLE
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath 'C:\path\to\legacy-workbooks' -Write
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$SourcePath = ".",

    [string]$OutputDirectory,

    [switch]$Recurse,

    [switch]$Write
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-ProgressLine {
    param(
        [Parameter(Mandatory)]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message
    )

    [Console]::Error.WriteLine("[{0}] {1}", $Level, $Message)
}

function Add-Result {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Results,

        [Parameter(Mandatory)]
        [string]$Status,

        [Parameter(Mandatory)]
        [System.IO.FileInfo]$Source,

        [string]$Destination,

        [AllowNull()]
        [Nullable[bool]]$HasVba,

        [string]$Message = ""
    )

    $format = if ($null -eq $HasVba) {
        $null
    }
    elseif ($HasVba) {
        "xlsm"
    }
    else {
        "xlsx"
    }

    [void]$Results.Add([pscustomobject][ordered]@{
        status      = $Status
        source      = $Source.FullName
        destination = $Destination
        format      = $format
        hasVba      = $HasVba
        message     = $Message
    })
}

$sourceItem = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
if (-not $sourceItem.PSIsContainer) {
    throw "SourcePath must be a folder: $($sourceItem.FullName)"
}

$sourceRoot = $sourceItem.FullName
$hasExplicitOutput = -not [string]::IsNullOrWhiteSpace($OutputDirectory)
$outputRoot = if ($hasExplicitOutput) {
    [System.IO.Path]::GetFullPath($OutputDirectory)
}
else {
    $null
}
$outputLabel = if ($hasExplicitOutput) { $outputRoot } else { "alongside each input workbook" }

$searchArguments = @{
    LiteralPath = $sourceRoot
    Filter      = "*.xls"
    File        = $true
}
if ($Recurse) {
    $searchArguments.Recurse = $true
}

# Windows wildcard matching may return names such as book.xlsx for "*.xls".
# Require the actual extension to be exactly .xls before opening anything.
$files = @(
    Get-ChildItem @searchArguments |
        Where-Object { $_.Extension -ieq ".xls" } |
        Sort-Object FullName
)
$results = [System.Collections.Generic.List[object]]::new()
$plannedDestinations = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)
$excel = $null

Write-ProgressLine "INFO" ("Mode={0} | Source={1} | Output={2} | Files={3}" -f (
        $(if ($Write) { "write" } else { "dry-run" })
    ), $sourceRoot, $outputLabel, $files.Count)

if ($files.Count -gt 0) {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false
        $excel.AutomationSecurity = 3
    }
    catch {
        $message = "Unable to start desktop Excel: $($_.Exception.Message)"
        foreach ($file in $files) {
            Add-Result -Results $results -Status "error" -Source $file -HasVba $null -Message $message
        }
        Write-ProgressLine "ERROR" $message
    }

    if ($null -ne $excel) {
        if ($Write -and $hasExplicitOutput) {
            [void](New-Item -ItemType Directory -Path $outputRoot -Force)
        }

        try {
            foreach ($file in $files) {
                $workbook = $null

                try {
                    # UpdateLinks=0 and ReadOnly=$true keep inspection non-mutating.
                    $workbook = $excel.Workbooks.Open($file.FullName, 0, $true)
                    $hasVba = [bool]$workbook.HasVBProject
                    $extension = if ($hasVba) { ".xlsm" } else { ".xlsx" }
                    $fileFormat = if ($hasVba) { 52 } else { 51 }
                    $destinationDirectory = if ($hasExplicitOutput) {
                        $outputRoot
                    }
                    else {
                        $file.DirectoryName
                    }
                    $destination = Join-Path $destinationDirectory ($file.BaseName + $extension)

                    if (-not $plannedDestinations.Add($destination)) {
                        $message = "Another source workbook maps to the same output name."
                        Add-Result -Results $results -Status "collision" -Source $file `
                            -Destination $destination -HasVba $hasVba -Message $message
                        Write-ProgressLine "ERROR" ("[collision] {0} -> {1}" -f $file.FullName, $destination)
                        continue
                    }

                    if (Test-Path -LiteralPath $destination) {
                        $message = "The destination already exists; it was not overwritten."
                        Add-Result -Results $results -Status "skipped-existing" -Source $file `
                            -Destination $destination -HasVba $hasVba -Message $message
                        Write-ProgressLine "WARNING" ("[skipped-existing] {0} -> {1}" -f $file.Name, $destination)
                        continue
                    }

                    if (-not $Write) {
                        Add-Result -Results $results -Status "would-convert" -Source $file `
                            -Destination $destination -HasVba $hasVba
                        Write-ProgressLine "INFO" ("[would-convert] {0} -> {1}" -f $file.Name, $destination)
                        continue
                    }

                    [void]$workbook.SaveAs($destination, $fileFormat)
                    Add-Result -Results $results -Status "converted" -Source $file `
                        -Destination $destination -HasVba $hasVba
                    Write-ProgressLine "SUCCESS" ("[converted] {0} -> {1}" -f $file.Name, $destination)
                }
                catch {
                    $message = $_.Exception.Message
                    Add-Result -Results $results -Status "error" -Source $file -HasVba $null -Message $message
                    Write-ProgressLine "ERROR" ("[error] {0} | {1}" -f $file.FullName, $message)
                }
                finally {
                    if ($null -ne $workbook) {
                        $workbook.Close($false)
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
                    }
                }
            }
        }
        finally {
            $excel.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}
else {
    Write-ProgressLine "INFO" "No .xls workbooks were found."
}

$counts = [ordered]@{}
foreach ($status in @("would-convert", "converted", "skipped-existing", "collision", "error")) {
    $counts[$status] = @($results | Where-Object status -EQ $status).Count
}

$summary = [ordered]@{
    mode            = if ($Write) { "write" } else { "dry-run" }
    sourcePath      = $sourceRoot
    outputMode      = if ($hasExplicitOutput) { "specified-directory" } else { "alongside-input" }
    outputDirectory = $outputRoot
    recursive       = [bool]$Recurse
    filesFound      = $files.Count
    counts          = $counts
}

$summary | ConvertTo-Json -Depth 5 -Compress

if (($counts.error + $counts.collision) -gt 0) {
    exit 1
}
