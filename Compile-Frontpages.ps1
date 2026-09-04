<#
.SYNOPSIS
    Ongoing Multi-Year Philippine Star Front Page Compiler
    Source: <Selected_Drive>:\COPIED FOLDERS\<Year>\<Month> (or <Drive>:\COPIED FOLDERS\<Month>)
    Target: Desktop\FRONT_PAGE\<Year>\<Month>

.DESCRIPTION
    - Ongoing Watch Mode: Stays running in real-time, continuously monitoring COPIED FOLDERS
      as files stream from the server. Compiles each new day as soon as its copy finishes.
    - Safe Copy Detection: Verifies file length, checks exclusive read access (FileShare.None),
      and validates standard PDF EOF marker (%%EOF) to prevent reading partially copied files.
    - Multi-Year Support: Automatically detects any year (2014, 2015, 2016, etc.)
    - Destination: ALWAYS Desktop\FRONT_PAGE\<Year>\<Month> (100% AUTOMATIC).
    - Renaming:
        * Primary Front Page: PS_YYYY_MMDD.pdf
        * False Page / Wrap:  PS_YYYY_MMDD_FALSE.pdf (strictly _FALSE, never _copy)
    - Cleanup Prompt: Asks user if they want to delete scanned source months from Drive D:
      to reclaim disk space once compiled.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Month = "",

    [Parameter()]
    [string]$Drive = "",

    [Parameter()]
    [switch]$Watch = $false,

    [Parameter()]
    [int]$IntervalSeconds = 3,

    [Parameter()]
    [int]$StartDay = 1,

    [Parameter()]
    [int]$EndDay = 31,

    [Parameter()]
    [switch]$Overwrite = $false,

    [Parameter()]
    [switch]$DeleteSource = $false,

    [Parameter()]
    [switch]$NonInteractive = $false
)

$ErrorActionPreference = "Stop"

# ============================================================
# High-Performance In-Memory PDF Stream Text Extractor (.NET C#)
# ============================================================
$csharpCode = @'
using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

public class PSFrontPageDetector {
    public static bool IsFileReady(string path) {
        if (!File.Exists(path)) return false;
        try {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None)) {
                if (stream.Length < 1000) return false;
                // Check if PDF ends with %%EOF
                long scanLen = Math.Min(1024, stream.Length);
                stream.Seek(-scanLen, SeekOrigin.End);
                byte[] tail = new byte[scanLen];
                int read = stream.Read(tail, 0, (int)scanLen);
                string tailStr = Encoding.ASCII.GetString(tail, 0, read);
                return tailStr.IndexOf("%%EOF", StringComparison.Ordinal) >= 0;
            }
        } catch {
            return false; // File is locked, being written, or incomplete
        }
    }

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
        bool hasMotto = text.IndexOf("TRUTH SHALL PREVAIL", StringComparison.OrdinalIgnoreCase) >= 0;

        return hasPhilstarWeb || hasMotto;
    }
}
'@

if (-not ([System.Management.Automation.PSTypeName]'PSFrontPageDetector').Type) {
    Add-Type -TypeDefinition $csharpCode
}

# Base Target is ALWAYS Desktop\FRONT_PAGE
$DesktopPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop"
$MasterTargetBase = Join-Path -Path $DesktopPath -ChildPath "FRONT_PAGE"

# --- STEP 1: Select Drive ---
if ($Drive -eq "") {
    if (-not $NonInteractive) {
        Write-Host "============================================================" -ForegroundColor Green
        Write-Host "         Philippine Star Front Page Compiler                " -ForegroundColor Green
        Write-Host "============================================================" -ForegroundColor Green

        $availableDrives = Get-PSDrive -PSProvider FileSystem | Sort-Object Name
        Write-Host "`nSelect Source Drive:" -ForegroundColor Yellow
        foreach ($d in $availableDrives) {
            $freeGB = [math]::Round($d.Free / 1GB, 1)
            $isDef = if ($d.Name -eq "D") { " (Default)" } else { "" }
            Write-Host "  [$($d.Name)] Drive $($d.Name): (Free: $freeGB GB)$isDef" -ForegroundColor White
        }

        $dInput = Read-Host "`nEnter Drive letter [D/C] (Default: D)"
        $Drive = if ([string]::IsNullOrWhiteSpace($dInput)) { "D" } else { $dInput.Trim().TrimEnd(':').ToUpper() }
    } else {
        $Drive = "D"
    }
} else {
    $Drive = $Drive.Trim().TrimEnd(':').ToUpper()
}

# Folder is ALWAYS COPIED FOLDERS
$BaseSourceDir = "${Drive}:\COPIED FOLDERS"

if (-not (Test-Path $BaseSourceDir)) {
    Write-Host "`n[ERROR] '$BaseSourceDir' was not found on Drive ${Drive}:" -ForegroundColor Red
    return
}

# Helper: Detect Year from directory hierarchy or day folder
function Get-FolderYear {
    param([string]$FolderPath)
    $parentName = Split-Path (Split-Path $FolderPath -Parent) -Leaf
    if ($parentName -match "^(19\d\d|20\d\d)$") { return $Matches[1] }
    $leafName = Split-Path $FolderPath -Leaf
    if ($leafName -match "\b(19\d\d|20\d\d)\b") { return $Matches[1] }
    $dayFolders = Get-ChildItem -Path $FolderPath -Directory -ErrorAction SilentlyContinue
    foreach ($df in $dayFolders) {
        if ($df.Name -match "^(\d{4})") { return $Matches[1] }
    }
    return "2015"
}

# Helper: Scan COPIED FOLDERS for all available month directories (handles both <Year>\<Month> and <Month>)
function Get-AvailableMonthFolders {
    param([string]$BaseDir)
    $items = @()
    if (-not (Test-Path $BaseDir)) { return $items }
    $topDirs = Get-ChildItem -Path $BaseDir -Directory -ErrorAction SilentlyContinue | Sort-Object Name
    foreach ($d in $topDirs) {
        if ($d.Name -match "^(19\d\d|20\d\d)$") {
            # Folder is a Year directory (e.g. 2015)
            $yearStr = $d.Name
            $subMonthDirs = Get-ChildItem -Path $d.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name
            foreach ($smd in $subMonthDirs) {
                $mNum = if ($smd.Name -match "^(\d{1,2})") { [int]$Matches[1] } else { 0 }
                $items += [PSCustomObject]@{
                    Year          = $yearStr
                    MonthName     = $smd.Name
                    DisplayName   = "$yearStr\$($smd.Name)"
                    FullName      = $smd.FullName
                    MonthNum      = $mNum
                    ParentYearDir = $d.FullName
                }
            }
        } else {
            # Folder is a Month directory directly under COPIED FOLDERS (e.g. 08 AUG)
            $y = Get-FolderYear -FolderPath $d.FullName
            $mNum = if ($d.Name -match "^(\d{1,2})") { [int]$Matches[1] } else { 0 }
            $items += [PSCustomObject]@{
                Year          = $y
                MonthName     = $d.Name
                DisplayName   = "$y\$($d.Name)"
                FullName      = $d.FullName
                MonthNum      = $mNum
                ParentYearDir = $null
            }
        }
    }
    return $items
}

# --- STEP 2: Mode Selection ---
if (-not $PSBoundParameters.ContainsKey('Watch') -and -not $NonInteractive) {
    Write-Host "`nSelect Execution Mode:" -ForegroundColor Yellow
    Write-Host "  [1] Ongoing Watch Mode (Continuously monitors & compiles days as they copy) (Default)" -ForegroundColor White
    Write-Host "  [2] Run Once (Compile existing files now and exit)" -ForegroundColor White
    $modeChoice = Read-Host "`nChoose mode [1/2] (Default: 1)"
    if ([string]::IsNullOrWhiteSpace($modeChoice) -or $modeChoice.Trim() -eq "1") {
        $Watch = $true
    } else {
        $Watch = $false
    }
}

# --- STEP 3: Month Selection ---
$detectedMonths = Get-AvailableMonthFolders -BaseDir $BaseSourceDir
if ($Month -eq "" -and -not $NonInteractive) {
    Write-Host "`nAvailable folder(s) inside '$BaseSourceDir':" -ForegroundColor Yellow
    if ($detectedMonths.Count -gt 0) {
        foreach ($item in $detectedMonths) {
            Write-Host "  [$($item.MonthNum)] $($item.DisplayName)" -ForegroundColor White
        }
    } else {
        Write-Host "  (No month folders present yet - will monitor for incoming folders)" -ForegroundColor DarkGray
    }
    Write-Host "  [A] Process ALL available / incoming folders (Default)" -ForegroundColor White

    $selection = Read-Host "`nType month number (e.g. 8) or 'A' for ALL (Default: A)"
    if ([string]::IsNullOrWhiteSpace($selection) -or $selection.Trim().ToUpper() -eq "A") {
        $Month = "ALL"
    } else {
        $Month = $selection.Trim()
    }
}

# --- Core Processor Function for a Single Day Folder ---
function Process-DayFolder {
    param(
        [System.IO.DirectoryInfo]$DayFolder,
        [string]$TargetMonthDir,
        [string]$TargetYear,
        [ref]$ProcessedCount,
        [ref]$FalseCount
    )

    if ($DayFolder.Name -notmatch "^(\d{4})(\d{2})(\d{2})") { return $false }
    $fileYear = $Matches[1]
    $monthStr = $Matches[2]
    $dayStr   = $Matches[3]
    $dayNum   = [int]$dayStr

    if ($dayNum -lt $StartDay -or $dayNum -gt $EndDay) { return $false }

    $psFolder = Get-ChildItem -Path $DayFolder.FullName -Directory -Filter "PS_*" | Select-Object -First 1
    if (-not $psFolder) { return $false }

    $candidates = Get-ChildItem -Path $psFolder.FullName -Filter "*.pdf" -ErrorAction SilentlyContinue | 
                  Where-Object { $_.Name -match "^00[0-4]\.pdf$" } | 
                  Sort-Object Name

    if ($candidates.Count -eq 0) { return $false }

    # Every published day must have 001.pdf present before we process
    $has001 = $candidates | Where-Object { $_.Name -eq "001.pdf" }
    if (-not $has001) { return $false }

    # Ensure all candidates are fully written and closed by the copy process
    foreach ($cand in $candidates) {
        if (-not [PSFrontPageDetector]::IsFileReady($cand.FullName)) {
            return $false # File is still being copied; retry on next cycle
        }
    }

    # Inspect candidate files for Front Page masthead / banner (000.pdf is always a false wrap)
    $frontPagesFound = @()
    foreach ($cand in $candidates) {
        if ($cand.Name -eq "000.pdf" -or [PSFrontPageDetector]::HasFrontPageLogo($cand.FullName)) {
            $frontPagesFound += $cand
        }
    }

    if ($frontPagesFound.Count -eq 0) { return $false }

    # Save front pages with exact naming standards (strictly _FALSE for false copies, never _copy)
    if ($frontPagesFound.Count -eq 1) {
        $isFalseSingle = ($frontPagesFound[0].Name -eq "000.pdf")
        $destFileName = if (-not $isFalseSingle) { "PS_${fileYear}_${monthStr}${dayStr}.pdf" } else { "PS_${fileYear}_${monthStr}${dayStr}_FALSE.pdf" }
        $destPath = Join-Path -Path $TargetMonthDir -ChildPath $destFileName

        if (-not (Test-Path $destPath) -or $Overwrite) {
            Copy-Item -Path $frontPagesFound[0].FullName -Destination $destPath -Force
            if (-not $isFalseSingle) {
                Write-Host "    [COPIED] $($DayFolder.Name) ($($frontPagesFound[0].Name)) -> $destFileName" -ForegroundColor Green
                $ProcessedCount.Value++
            } else {
                Write-Host "    [COPIED FALSE] $($DayFolder.Name) ($($frontPagesFound[0].Name)) -> $destFileName" -ForegroundColor Magenta
                $FalseCount.Value++
            }
        }
    } else {
        # Multiple front pages (wraparound advertising cover + primary editorial front page)
        $falseIndex = 0
        foreach ($fp in $frontPagesFound) {
            if ($fp.Name -eq "001.pdf") {
                # 001.pdf is the primary editorial front page
                $destFileName = "PS_${fileYear}_${monthStr}${dayStr}.pdf"
                $destPath = Join-Path -Path $TargetMonthDir -ChildPath $destFileName
                if (-not (Test-Path $destPath) -or $Overwrite) {
                    Copy-Item -Path $fp.FullName -Destination $destPath -Force
                    Write-Host "    [COPIED] $($DayFolder.Name) ($($fp.Name)) -> $destFileName" -ForegroundColor Green
                    $ProcessedCount.Value++
                }
            } else {
                # Any other candidate (000.pdf, 002.pdf, etc.) is strictly a FALSE copy
                $falseSuffix = if ($falseIndex -eq 0) { "_FALSE.pdf" } else { "_FALSE$($falseIndex + 1).pdf" }
                $destFileName = "PS_${fileYear}_${monthStr}${dayStr}${falseSuffix}"
                $destPath = Join-Path -Path $TargetMonthDir -ChildPath $destFileName
                if (-not (Test-Path $destPath) -or $Overwrite) {
                    Copy-Item -Path $fp.FullName -Destination $destPath -Force
                    Write-Host "    [COPIED FALSE] $($DayFolder.Name) ($($fp.Name)) -> $destFileName" -ForegroundColor Magenta
                    $FalseCount.Value++
                }
                $falseIndex++
            }
        }
    }

    return $true
}

# --- Function: Cleanup Scanned Source Folders ---
function Prompt-FolderCleanup {
    param(
        [array]$CompletedFolders,
        [string]$DriveLetter,
        [bool]$AutoDelete,
        [bool]$IsNonInteractive
    )

    if ($CompletedFolders.Count -eq 0) { return }

    $shouldDelete = $false
    if ($AutoDelete) {
        $shouldDelete = $true
    } elseif (-not $IsNonInteractive) {
        Write-Host "`n============================================================" -ForegroundColor Yellow
        Write-Host "             SOURCE FOLDER CLEANUP (FREE DISK SPACE)        " -ForegroundColor Yellow
        Write-Host "============================================================" -ForegroundColor Yellow
        Write-Host "The following source month folder(s) have been processed:" -ForegroundColor White
        foreach ($sf in $CompletedFolders) {
            Write-Host "  - $($sf.FullName)" -ForegroundColor Gray
        }
        Write-Host ""
        $delInput = Read-Host "Do you want to delete these scanned month folder(s) from Drive ${DriveLetter}: to free up space? [Y/N] (Default: N)"
        if ($delInput -match "^[Yy]") {
            $shouldDelete = $true
        }
    }

    if ($shouldDelete) {
        Write-Host "`nDeleting source month folder(s) from Drive ${DriveLetter}:..." -ForegroundColor Yellow
        foreach ($sf in $CompletedFolders) {
            try {
                if (Test-Path $sf.FullName) {
                    Remove-Item -Path $sf.FullName -Recurse -Force
                    Write-Host "  [DELETED] $($sf.FullName) (Disk space freed)" -ForegroundColor Green
                }
                if ($sf.ParentYearDir -and (Test-Path $sf.ParentYearDir)) {
                    $remaining = Get-ChildItem -Path $sf.ParentYearDir -ErrorAction SilentlyContinue
                    if ($remaining.Count -eq 0) {
                        Remove-Item -Path $sf.ParentYearDir -Force -ErrorAction SilentlyContinue
                        Write-Host "  [CLEANED EMPTY YEAR DIR] $($sf.ParentYearDir)" -ForegroundColor DarkGray
                    }
                }
            } catch {
                Write-Host "  [ERROR] Could not delete $($sf.FullName): $_" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "`n[KEPT] Source folder(s) kept on Drive ${DriveLetter}:" -ForegroundColor Cyan
    }
}

# ============================================================
#                      EXECUTION MODES
# ============================================================
if ($Watch) {
    # ------------------------------------------------------------
    #              ONGOING WATCH MODE (REAL-TIME)
    # ------------------------------------------------------------
    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host "    ONGOING WATCH MODE: WAITING & MONITORING LIVE COPY      " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " Source Base : $BaseSourceDir" -ForegroundColor Gray
    Write-Host " Target Base : $MasterTargetBase (Desktop\FRONT_PAGE\<Year>\<Month>)" -ForegroundColor Green
    Write-Host " Month Filter: $(if ($Month -eq '' -or $Month -eq 'ALL') { 'ALL incoming folders' } else { $Month })" -ForegroundColor Gray
    Write-Host " Check Delay : Every $IntervalSeconds seconds" -ForegroundColor Gray
    Write-Host " Controls    : Press 'Q' at any time to finish & cleanup." -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Green

    $processedDaysMap = @{}
    $liveProcessed = 0
    $liveFalse = 0
    $monitoredSourceFolders = @{}
    $completedMonthAnnounced = @{}

    # Pre-populate already compiled target files to avoid unnecessary re-work
    if (Test-Path $MasterTargetBase) {
        $existingFrontPages = Get-ChildItem -Path $MasterTargetBase -Recurse -Filter "PS_*.pdf" -ErrorAction SilentlyContinue
        foreach ($efp in $existingFrontPages) {
            if ($efp.Name -match "^PS_(\d{4})_(\d{2})(\d{2})") {
                $pYear = $Matches[1]
                $pMonthNum = [int]$Matches[2]
                $pDay = $Matches[3]
                # Store date key
                $dateKey = "${pYear}${Matches[2]}${pDay}"
                $processedDaysMap[$dateKey] = $true
            }
        }
    }

    try {
        while ($true) {
            # Check for user pressing 'Q' to cleanly exit watch mode
            try {
                if ([Console]::KeyAvailable) {
                    $key = [Console]::ReadKey($true)
                    if ($key.Key -eq [System.ConsoleKey]::Q) {
                        Write-Host "`n[STOPPED] Watch mode ended by user." -ForegroundColor Yellow
                        break
                    }
                }
            } catch {}

            $availableMonths = Get-AvailableMonthFolders -BaseDir $BaseSourceDir
            if ($Month -ne "" -and $Month -ne "ALL") {
                $cleanMonth = $Month.Trim()
                $availableMonths = $availableMonths | Where-Object { 
                    $_.MonthNum -eq $cleanMonth -or $_.MonthName -like "*$cleanMonth*" -or $_.DisplayName -like "*$cleanMonth*" 
                }
            }

            foreach ($mf in $availableMonths) {
                $monitoredSourceFolders[$mf.FullName] = $mf
                $yearTargetDir = Join-Path -Path $MasterTargetBase -ChildPath $mf.Year
                $monthTargetDir = Join-Path -Path $yearTargetDir -ChildPath $mf.MonthName
                if (-not (Test-Path $monthTargetDir)) {
                    New-Item -ItemType Directory -Path $monthTargetDir -Force | Out-Null
                    Write-Host "`n[TARGET READY] $monthTargetDir" -ForegroundColor Cyan
                }

                $dayFolders = Get-ChildItem -Path $mf.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name
                foreach ($df in $dayFolders) {
                    $uniqueKey = "$($mf.DisplayName)\$($df.Name)"
                    $dateKey = $df.Name

                    if (-not $processedDaysMap.ContainsKey($uniqueKey) -and -not $processedDaysMap.ContainsKey($dateKey)) {
                        $pRef = [ref]$liveProcessed
                        $fRef = [ref]$liveFalse
                        $didProcess = Process-DayFolder -DayFolder $df -TargetMonthDir $monthTargetDir -TargetYear $mf.Year -ProcessedCount $pRef -FalseCount $fRef
                        
                        if ($didProcess) {
                            $processedDaysMap[$uniqueKey] = $true
                            $processedDaysMap[$dateKey] = $true
                        }
                    }
                }

                # Check if this month folder has completed copying all days
                $compiledInTarget = Get-ChildItem -Path $monthTargetDir -Filter "PS_*.pdf" -ErrorAction SilentlyContinue | 
                                   Where-Object { $_.Name -notmatch "_FALSE\.pdf$" }
                $expectedDays = try { [DateTime]::DaysInMonth([int]$mf.Year, [int]$mf.MonthNum) } catch { 31 }
                
                if ($compiledInTarget.Count -ge $expectedDays -and -not $completedMonthAnnounced.ContainsKey($mf.FullName)) {
                    $completedMonthAnnounced[$mf.FullName] = $true
                    Write-Host "`n============================================================" -ForegroundColor Green
                    Write-Host " [MONTH COMPLETE] $($mf.DisplayName): All $($compiledInTarget.Count)/$expectedDays days compiled!" -ForegroundColor Green
                    Write-Host "============================================================" -ForegroundColor Green
                }
            }

            # Heartbeat status line
            $timestamp = (Get-Date).ToString("hh:mm:ss tt")
            Write-Host -NoNewline "`r[$timestamp] Monitoring $BaseSourceDir | Compiled: $liveProcessed new frontpages | Press 'Q' to finish...   "
            Start-Sleep -Seconds $IntervalSeconds
        }
    } finally {
        Write-Host "`n`n============================================================" -ForegroundColor Green
        Write-Host "                     WATCH MODE SUMMARY                     " -ForegroundColor Green
        Write-Host "============================================================" -ForegroundColor Green
        Write-Host " Newly Compiled Frontpages: $liveProcessed file(s)" -ForegroundColor Green
        Write-Host " Newly Compiled False Pages: $liveFalse file(s)" -ForegroundColor Magenta
        Write-Host " Destination Base         : $MasterTargetBase" -ForegroundColor Cyan
        Write-Host "============================================================" -ForegroundColor Green

        $completedList = @($monitoredSourceFolders.Values)
        Prompt-FolderCleanup -CompletedFolders $completedList -DriveLetter $Drive -AutoDelete $DeleteSource -IsNonInteractive $NonInteractive
    }

} else {
    # ------------------------------------------------------------
    #                      RUN ONCE MODE
    # ------------------------------------------------------------
    $detectedItems = Get-AvailableMonthFolders -BaseDir $BaseSourceDir
    if ($detectedItems.Count -eq 0) {
        Write-Host "`n[INFO] No month folders found in '$BaseSourceDir'." -ForegroundColor Yellow
        Write-Host "Please ensure month folders (e.g. '08 AUG' or '2015\08 AUG') are in '$BaseSourceDir'." -ForegroundColor Cyan
        return
    }

    $selectedFolders = @()
    if ($Month -ne "" -and $Month -ne "ALL") {
        $cleanMonth = $Month.Trim()
        $selectedFolders = @($detectedItems | Where-Object { 
            $_.MonthNum -eq $cleanMonth -or $_.MonthName -like "*$cleanMonth*" -or $_.DisplayName -like "*$cleanMonth*" 
        })
    } else {
        $selectedFolders = @($detectedItems)
    }

    if ($selectedFolders.Count -eq 0) {
        Write-Host "`n[WARNING] No folders matched month selection '$Month'." -ForegroundColor Yellow
        return
    }

    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host "         Philippine Star Front Page Compiler                " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " Source Base   : $BaseSourceDir" -ForegroundColor Gray
    Write-Host " Target Base   : $MasterTargetBase (Desktop\FRONT_PAGE\<Year>\<Month>)" -ForegroundColor Green
    Write-Host " Folders Count : $($selectedFolders.Count) ($(($selectedFolders.DisplayName) -join ', '))" -ForegroundColor Gray
    Write-Host " Naming Format : Front Page -> PS_YYYY_MMDD.pdf" -ForegroundColor Gray
    Write-Host "                 False Cover-> PS_YYYY_MMDD_FALSE.pdf" -ForegroundColor Gray
    Write-Host "============================================================" -ForegroundColor Green

    $totalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $grandProcessed = 0
    $grandFalse = 0
    $completedSourceFolders = @()

    foreach ($mf in $selectedFolders) {
        $yearTargetDir = Join-Path -Path $MasterTargetBase -ChildPath $mf.Year
        $monthTargetDir = Join-Path -Path $yearTargetDir -ChildPath $mf.MonthName
        if (-not (Test-Path $monthTargetDir)) {
            New-Item -ItemType Directory -Path $monthTargetDir -Force | Out-Null
            Write-Host "`n[AUTO-CREATED] $monthTargetDir" -ForegroundColor Cyan
        }

        Write-Host "`n>>> Processing: $($mf.DisplayName)" -ForegroundColor Yellow
        $dayFolders = Get-ChildItem -Path $mf.FullName -Directory | Sort-Object Name
        $mProcessed = 0
        $mFalse = 0

        foreach ($df in $dayFolders) {
            $pRef = [ref]$mProcessed
            $fRef = [ref]$mFalse
            $null = Process-DayFolder -DayFolder $df -TargetMonthDir $monthTargetDir -TargetYear $mf.Year -ProcessedCount $pRef -FalseCount $fRef
        }

        $grandProcessed += $mProcessed
        $grandFalse += $mFalse
        $completedSourceFolders += $mf
        Write-Host "    Month Summary: $mProcessed frontpage(s), $mFalse false page(s)." -ForegroundColor Cyan
    }

    $totalStopwatch.Stop()
    Write-Host "`n============================================================" -ForegroundColor Green
    Write-Host "                     FINAL SUMMARY                          " -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host " Total Elapsed Time: $($totalStopwatch.Elapsed.TotalSeconds.ToString("F2")) seconds" -ForegroundColor Cyan
    Write-Host " Total Copied      : $grandProcessed file(s)" -ForegroundColor Green
    Write-Host " Total False Pages : $grandFalse file(s)" -ForegroundColor Magenta
    Write-Host " Destination Base  : $MasterTargetBase" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Green

    Prompt-FolderCleanup -CompletedFolders $completedSourceFolders -DriveLetter $Drive -AutoDelete $DeleteSource -IsNonInteractive $NonInteractive
}
