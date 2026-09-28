<#
.SYNOPSIS
    Batch-convert PowerPoint files (.ppt, .pptx, .pptm, .pps, .ppsx) to PDF using Microsoft PowerPoint.

.EXAMPLE
    .\Convert-PptToPdf.ps1                         # converts everything in .\input -> .\output
    .\Convert-PptToPdf.ps1 -InputPath "C:\Decks"   # converts a folder
    .\Convert-PptToPdf.ps1 a.pptx b.ppt c.pptx     # converts specific files
    .\Convert-PptToPdf.ps1 -InputPath "C:\Decks" -Recurse -OutputDir "C:\PDFs"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$InputPath,

    [string]$OutputDir,

    [switch]$Recurse,

    [switch]$Overwrite
)

# $PSScriptRoot isn't reliable inside param() defaults on Windows PowerShell 5.1
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $InputPath) { $InputPath = @(Join-Path $scriptDir 'input') }

$extensions = @('.ppt', '.pptx', '.pptm', '.pps', '.ppsx')

# Collect files from the given paths (files and/or folders)
$files = foreach ($p in $InputPath) {
    if (-not (Test-Path -LiteralPath $p)) { Write-Warning "Not found: $p"; continue }
    $item = Get-Item -LiteralPath $p
    if ($item.PSIsContainer) {
        Get-ChildItem -LiteralPath $item.FullName -File -Recurse:$Recurse |
            Where-Object { $extensions -contains $_.Extension.ToLower() -and $_.Name -notlike '~$*' }
    } elseif ($extensions -contains $item.Extension.ToLower()) {
        $item
    } else {
        Write-Warning "Skipping (not a PowerPoint file): $p"
    }
}
$files = @($files)

if ($files.Count -eq 0) {
    Write-Host "No PowerPoint files found." -ForegroundColor Yellow
    return
}

if (-not $OutputDir) { $OutputDir = Join-Path $scriptDir 'output' }
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$OutputDir = (Resolve-Path -LiteralPath $OutputDir).Path

Write-Host "Converting $($files.Count) file(s) -> $OutputDir`n"

$ppSaveAsPDF = 32
$msoFalse = 0
$msoTrue = -1

$powerpoint = New-Object -ComObject PowerPoint.Application
$ok = 0; $failed = 0; $skipped = 0
$i = 0

try {
    foreach ($file in $files) {
        $i++
        $pdfPath = Join-Path $OutputDir ($file.BaseName + '.pdf')
        $label = "[$i/$($files.Count)] $($file.Name)"

        if ((Test-Path -LiteralPath $pdfPath) -and -not $Overwrite) {
            Write-Host "$label  -> skipped (PDF exists, use -Overwrite)" -ForegroundColor DarkYellow
            $skipped++
            continue
        }

        $pres = $null
        try {
            # Open(FileName, ReadOnly, Untitled, WithWindow)
            $pres = $powerpoint.Presentations.Open($file.FullName, $msoTrue, $msoFalse, $msoFalse)
            $pres.SaveAs($pdfPath, $ppSaveAsPDF)
            Write-Host "$label  -> OK" -ForegroundColor Green
            $ok++
        } catch {
            Write-Host "$label  -> FAILED: $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        } finally {
            if ($pres) {
                $pres.Close()
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
            }
        }
    }
} finally {
    $powerpoint.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "`nDone. Converted: $ok  Skipped: $skipped  Failed: $failed"
Write-Host "PDFs are in: $OutputDir"
