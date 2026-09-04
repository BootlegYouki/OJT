<#
.SYNOPSIS
    Dynamic & Interactive Philippine Star Front Page Compiler (January - December)
    Allows user to select the Drive and Folder as the source directory, scans for
    Philippine Star (PS) frontpages (pages 000 through 004), and compiles them into target directories.

.DESCRIPTION
    - Interactive drive and folder selection (or fully automated via parameters).
    - Can scan any drive (C:, D:, E:, etc.) and any folder name (e.g. "copied folders").
    - Accurately identifies frontpage logos via embedded masthead markers:
        * "TRUTH SHALL PREVAIL" (motto under logo)
        * "philstar.com" / "ThePhilippineStar" + "VOL." (issue volume header)
    - Distinguishes main editorial frontpage from false covers / wrap covers:
        * Primary Front Page: PS_YYYY_MMDD.pdf
        * False Page / Wrap Cover: PS_YYYY_MMDD_FALSE.pdf
    - High Performance: In-memory .NET C# stream decompilation (processes an entire month in ~5 seconds).

.EXAMPLE
    # Interactive mode (prompts for drive, folder, and month)
    .\Compile-Frontpages.ps1 -Interactive

    # Specify drive and folder via parameters
    .\Compile-Frontpages.ps1 -Drive D -FolderName "copied folders" -Month 5

    # Direct source and target paths
    .\Compile-Frontpages.ps1 -SourceDir "D:\copied folders\05 MAY" -TargetDir "C:\Users\OJT_KIRK_DARREN\Desktop\FRONT_PAGE_2015\05 MAY"
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Month = "",

    [Parameter()]
    [string]$Drive = "",

    [Parameter()]
    [string]$FolderName = "",

    [Parameter()]
    [string]$SourceDir = "",

    [Parameter()]
    [string]$TargetDir = "",

    [Parameter()]
    [string]$BaseSourceDir = "",

    [Parameter()]
    [string]$BaseTargetDir = "C:\Users\OJT_KIRK_DARREN\Desktop\FRONT_PAGE_2015",

    [Parameter()]
    [int]$StartDay = 1,

    [Parameter()]
    [int]$EndDay = 31,

    [Parameter()]
    [switch]$Overwrite = $false,

    [Parameter()]
    [switch]$Interactive = $false
)

$ErrorActionPreference = "Stop"

# High-Performance In-Memory PDF Stream Text Extractor (.NET C#)
$csharpCode = @'
using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

public class PSFrontPageDetector {
    public static string ExtractAllText(string path) {
        if (!File.Exists(path)) return "";
        byte[] bytes;
        try {
            bytes = File.ReadAllBytes(path);
        } catch {
            return "";
        }
        string pdf = Encoding.ASCII.GetString(bytes);
        StringBuilder sb = new StringBuilder();

        if (pdf.IndexOf("philstar", StringComparison.OrdinalIgnoreCase) >= 0) sb.Append("philstar ");
        if (pdf.IndexOf("PHILIPPINE STAR", StringComparison.OrdinalIgnoreCase) >= 0) sb.Append("PHILIPPINE STAR ");

        int pos = 0;
        while (true) {
            int sStart = pdf.IndexOf("stream", pos, StringComparison.Ordinal);
            if (sStart < 0) break;
            int dataStart = sStart + 6;
            if (dataStart < bytes.Length && bytes[dataStart] == '\r') dataStart++;
            if (dataStart < bytes.Length && bytes[dataStart] == '\n') dataStart++;

            int sEnd = pdf.IndexOf("endstream", dataStart);
            if (sEnd < 0) break;

            int len = sEnd - dataStart;
            while (len > 0 && (bytes[dataStart + len - 1] == '\r' || bytes[dataStart + len - 1] == '\n')) {
                len--;
            }

            if (len > 6 && (bytes[dataStart] * 256 + bytes[dataStart + 1]) % 31 == 0 && (bytes[dataStart] & 0x0F) == 8) {
                try {
                    using (var ms = new MemoryStream(bytes, dataStart + 2, len - 6))
                    using (var ds = new DeflateStream(ms, CompressionMode.Decompress))
                    using (var outMs = new MemoryStream()) {
                        ds.CopyTo(outMs);
                        string streamStr = Encoding.ASCII.GetString(outMs.ToArray());
                        var matches = Regex.Matches(streamStr, @"\(((\\.|[^\\)])*)\)");
                        foreach (Match m in matches) {
                            sb.Append(m.Groups[1].Value);
                        }
                        sb.Append(" ");
                    }
                } catch {}
            }
            pos = sEnd + 9;
        }
        return sb.ToString();
    }

    public static bool HasFrontPageLogo(string path) {
        string text = ExtractAllText(path);
        if (string.IsNullOrEmpty(text)) return false;

        bool hasPhilstarWeb = text.IndexOf("philstar.com", StringComparison.OrdinalIgnoreCase) >= 0;
        bool hasPhilstarHandle = text.IndexOf("ThePhilippineStar", StringComparison.OrdinalIgnoreCase) >= 0;
        bool hasVol = text.IndexOf("VOL.", StringComparison.OrdinalIgnoreCase) >= 0;
        bool hasMotto = text.IndexOf("TRUTH SHALL PREVAIL", StringComparison.OrdinalIgnoreCase) >= 0;

        // Frontpage masthead banner detection:
        // 1. Has both web and social handle banner: "philstar.com" + "ThePhilippineStar" (matches Sundays & false covers)
        // 2. Has volume line + PhilStar marker: "VOL." + ("philstar.com" or "ThePhilippineStar")
        // 3. Has frontpage motto: "TRUTH SHALL PREVAIL"
        return (hasPhilstarWeb && hasPhilstarHandle) ||
               (hasPhilstarWeb && hasVol) ||
               (hasPhilstarHandle && hasVol) ||
               hasMotto;
    }
}
'@

if (-not ([System.Management.Automation.PSTypeName]'PSFrontPageDetector').Type) {
    Add-Type -TypeDefinition $csharpCode
}

# Standard month names mapping
$monthMap = @{
    "1" = "01 JAN"; "01" = "01 JAN"; "JAN" = "01 JAN"; "JANUARY" = "01 JAN"; "01 JAN" = "01 JAN"
    "2" = "02 FEB"; "02" = "02 FEB"; "FEB" = "02 FEB"; "FEBRUARY" = "02 FEB"; "02 FEB" = "02 FEB"
    "3" = "03 MAR"; "03" = "03 MAR"; "MAR" = "03 MAR"; "MARCH" = "03 MAR"; "03 MAR" = "03 MAR"
    "4" = "04 APR"; "04" = "04 APR"; "APR" = "04 APR"; "APRIL" = "04 APR"; "04 APR" = "04 APR"
    "5" = "05 MAY"; "05" = "05 MAY"; "MAY" = "05 MAY"; "05 MAY" = "05 MAY"
    "6" = "06 JUN"; "06" = "06 JUN"; "JUN" = "06 JUN"; "JUNE" = "06 JUN"; "06 JUN" = "06 JUN"
    "7" = "07 JUL"; "07" = "07 JUL"; "JUL" = "07 JUL"; "JULY" = "07 JUL"; "07 JUL" = "07 JUL"
    "8" = "08 AUG"; "08" = "08 AUG"; "AUG" = "08 AUG"; "AUGUST" = "08 AUG"; "08 AUG" = "08 AUG"
    "9" = "09 SEP"; "09" = "09 SEP"; "SEP" = "09 SEP"; "SEPTEMBER" = "09 SEP"; "09 SEP" = "09 SEP"
    "10" = "10 OCT"; "OCT" = "10 OCT"; "OCTOBER" = "10 OCT"; "10 OCT" = "10 OCT"
    "11" = "11 NOV"; "NOV" = "11 NOV"; "NOVEMBER" = "11 NOV"; "11 NOV" = "11 NOV"
    "12" = "12 DEC"; "DEC" = "12 DEC"; "DECEMBER" = "12 DEC"; "12 DEC" = "12 DEC"
}

# --- Interactive Drive and Folder Selection ---
if ($Interactive -or ($Drive -eq "" -and $BaseSourceDir -eq "" -and $SourceDir -eq "")) {
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "   Philippine Star Front Page Compiler - Setup Source       " -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan

    # 1. Detect and List Available Drives
    $availableDrives = Get-PSDrive -PSProvider FileSystem | Sort-Object Name
    Write-Host "`nAvailable Drives:" -ForegroundColor Yellow
    $dIndex = 1
    $driveChoices = @{}
    $defaultDrive = "D"
    foreach ($drv in $availableDrives) {
        $freeGB = [math]::Round($drv.Free / 1GB, 1)
        Write-Host "  [$dIndex] Drive $($drv.Name): (Free: $freeGB GB)" -ForegroundColor White
        $driveChoices["$dIndex"] = $drv.Name
        $driveChoices[$drv.Name] = $drv.Name
        if ($drv.Name -eq "D") { $defaultDriveChoice = "$dIndex" }
        $dIndex++
    }

    $driveInput = Read-Host "Select Drive (Default: D)"
    if ([string]::IsNullOrWhiteSpace($driveInput)) {
        $selectedDrive = "D"
    } elseif ($driveChoices.ContainsKey($driveInput.ToUpper())) {
        $selectedDrive = $driveChoices[$driveInput.ToUpper()]
    } else {
        $selectedDrive = $driveInput.Trim().TrimEnd(':')
    }

    Write-Host "Selected Drive: ${selectedDrive}:" -ForegroundColor Green

    # 2. Select Source Folder on Drive
    $defaultFolder = "copied folders"
    $folderInput = Read-Host "`nEnter source folder on ${selectedDrive}:\ (Default: $defaultFolder)"
    if ([string]::IsNullOrWhiteSpace($folderInput)) {
        $selectedFolder = $defaultFolder
    } else {
        $selectedFolder = $folderInput.Trim()
    }

    $BaseSourceDir = "${selectedDrive}:\$selectedFolder"
    Write-Host "Source Base Path: $BaseSourceDir" -ForegroundColor Green

    # 3. Select Month (or ALL)
    $monthInput = Read-Host "`nEnter Month to process (e.g. 05 MAY, MAY, 5, or press Enter for ALL)"
    if (-not [string]::IsNullOrWhiteSpace($monthInput)) {
        $Month = $monthInput.Trim()
    }
}

# Resolve BaseSourceDir if Drive/FolderName specified
if ($BaseSourceDir -eq "") {
    if ($Drive -ne "") {
        $cleanDrive = $Drive.Trim().TrimEnd(':')
        $cleanFolder = if ($FolderName -ne "") { $FolderName.Trim() } else { "copied folders" }
        $BaseSourceDir = "${cleanDrive}:\$cleanFolder"
    } else {
        $BaseSourceDir = "D:\copied folders"
    }
}

# Verify Base Source Directory
if (-not (Test-Path $BaseSourceDir)) {
    Write-Host "[ERROR] Source folder does not exist: $BaseSourceDir" -ForegroundColor Red
    Write-Host "Please check the drive and folder name, or create '$BaseSourceDir' with month folders inside." -ForegroundColor Yellow
    return
}

# Determine which month folders to process
$tasks = @()

if ($SourceDir -ne "" -and $TargetDir -ne "") {
    $tasks += [PSCustomObject]@{
        MonthName = (Split-Path $SourceDir -Leaf)
        Source    = $SourceDir
        Target    = $TargetDir
    }
} elseif ($Month -ne "" -and $Month -ne "ALL") {
    $normKey = $Month.Trim().ToUpper()
    if ($monthMap.ContainsKey($normKey)) {
        $folderName = $monthMap[$normKey]
    } else {
        $folderName = $Month.Trim()
    }
    $sDir = Join-Path -Path $BaseSourceDir -ChildPath $folderName
    $tDir = Join-Path -Path $BaseTargetDir -ChildPath $folderName
    $tasks += [PSCustomObject]@{
        MonthName = $folderName
        Source    = $sDir
        Target    = $tDir
    }
} else {
    # Auto-Discovery: Scan BaseSourceDir for month folders matching "NN MMM"
    $foundDirs = Get-ChildItem -Path $BaseSourceDir -Directory | 
                 Where-Object { $_.Name -match "^\d{2}\s+[A-Za-z]{3}$" } | 
                 Sort-Object Name

    foreach ($d in $foundDirs) {
        $tasks += [PSCustomObject]@{
            MonthName = $d.Name
            Source    = $d.FullName
            Target    = Join-Path -Path $BaseTargetDir -ChildPath $d.Name
        }
    }
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "         Philippine Star Front Page Compiler                " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Source Base       : $BaseSourceDir" -ForegroundColor Gray
Write-Host " Target Base       : $BaseTargetDir" -ForegroundColor Gray
Write-Host " Months Detected   : $($tasks.Count) ($(($tasks.MonthName) -join ', '))" -ForegroundColor Gray
Write-Host " Day Filter        : Day $StartDay to $EndDay" -ForegroundColor Gray
Write-Host " Overwrite         : $Overwrite" -ForegroundColor Gray
Write-Host " Naming Convention : Primary -> PS_YYYY_MMDD.pdf" -ForegroundColor Gray
Write-Host "                     False   -> PS_YYYY_MMDD_FALSE.pdf" -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor Green

if ($tasks.Count -eq 0) {
    Write-Host "`n[INFO] No month folders (e.g. '05 MAY', '06 JUN') found in '$BaseSourceDir'." -ForegroundColor Yellow
    return
}

$totalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$grandProcessed = 0
$grandFalse = 0
$grandSkipped = 0

foreach ($task in $tasks) {
    Write-Host "`n>>> Processing Month: $($task.MonthName)" -ForegroundColor Yellow
    Write-Host "    Source: $($task.Source)" -ForegroundColor DarkGray
    Write-Host "    Target: $($task.Target)" -ForegroundColor DarkGray

    if (-not (Test-Path $task.Source)) {
        Write-Host "    [SKIP] Source folder does not exist: $($task.Source)" -ForegroundColor DarkYellow
        continue
    }

    if (-not (Test-Path $task.Target)) {
        New-Item -ItemType Directory -Path $task.Target -Force | Out-Null
        Write-Host "    [INFO] Created target directory: $($task.Target)" -ForegroundColor Cyan
    }

    $dayFolders = Get-ChildItem -Path $task.Source -Directory | Sort-Object Name
    $monthProcessed = 0
    $monthFalse = 0
    $monthSkipped = 0

    foreach ($folder in $dayFolders) {
        if ($folder.Name -match "^(\d{4})(\d{2})(\d{2})") {
            $year = $Matches[1]
            $monthStr = $Matches[2]
            $dayStr = $Matches[3]
            $dayNum = [int]$dayStr
        } else {
            continue
        }

        if ($dayNum -lt $StartDay -or $dayNum -gt $EndDay) {
            continue
        }

        $psFolder = Get-ChildItem -Path $folder.FullName -Directory -Filter "PS_*" | Select-Object -First 1
        if (-not $psFolder) {
            continue
        }

        $candidates = Get-ChildItem -Path $psFolder.FullName -Filter "*.pdf" | 
                      Where-Object { $_.Name -match "^00[0-4]\.pdf$" } | 
                      Sort-Object Name

        $frontPagesFound = @()
        foreach ($cand in $candidates) {
            if ([PSFrontPageDetector]::HasFrontPageLogo($cand.FullName)) {
                $frontPagesFound += $cand
            }
        }

        if ($frontPagesFound.Count -eq 0) {
            Write-Host "    [WARN] $($folder.Name): No front page detected among 000-004." -ForegroundColor Yellow
            continue
        }

        $copyIdx = 0
        foreach ($fp in $frontPagesFound) {
            $isFalse = $false
            if ($frontPagesFound.Count -gt 1) {
                if ($fp.Name -eq "000.pdf" -or $copyIdx -gt 0) {
                    $isFalse = $true
                }
            }

            if (-not $isFalse) {
                $destFileName = "PS_${year}_${monthStr}${dayStr}.pdf"
            } else {
                $destFileName = "PS_${year}_${monthStr}${dayStr}_FALSE.pdf"
            }

            $destPath = Join-Path -Path $task.Target -ChildPath $destFileName

            if ((Test-Path $destPath) -and -not $Overwrite) {
                $monthSkipped++
                $grandSkipped++
            } else {
                Copy-Item -Path $fp.FullName -Destination $destPath -Force
                if (-not $isFalse) {
                    Write-Host "    [COPIED] $($folder.Name) ($($fp.Name)) -> $destFileName" -ForegroundColor Green
                    $monthProcessed++
                    $grandProcessed++
                } else {
                    Write-Host "    [COPIED FALSE] $($folder.Name) ($($fp.Name)) -> $destFileName" -ForegroundColor Magenta
                    $monthFalse++
                    $grandFalse++
                }
            }
            $copyIdx++
        }
    }

    Write-Host "    Month Summary: $monthProcessed frontpage(s), $monthFalse false page(s), $monthSkipped skipped." -ForegroundColor Cyan
}

$totalStopwatch.Stop()
Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "                     FINAL SUMMARY                          " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Total Elapsed Time: $($totalStopwatch.Elapsed.TotalSeconds.ToString("F2")) seconds" -ForegroundColor Cyan
Write-Host " Total Copied      : $grandProcessed file(s)" -ForegroundColor Green
Write-Host " Total False Pages : $grandFalse file(s)" -ForegroundColor Magenta
Write-Host " Total Skipped     : $grandSkipped file(s)" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Green
