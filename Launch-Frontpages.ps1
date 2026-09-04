<#
.SYNOPSIS
    Philippine Star Front Page Viewer & Launcher
    Scans Desktop\FRONT_PAGE for compiled frontpages by Year and Month,
    then launches all pages in the month with the default system PDF viewer.

.DESCRIPTION
    - Interactive Year & Month selector.
    - Scans and counts all compiled PDF frontpages (including _FALSE copies).
    - Launches all pages into the default viewer (e.g. Microsoft Edge tabs or Adobe Reader)
      with a smooth pacing interval so all pages load reliably.
    - Supports opening the entire month folder in File Explorer.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$Year = "",

    [Parameter()]
    [string]$Month = "",

    [Parameter()]
    [switch]$OpenFolder = $false,

    [Parameter()]
    [switch]$NonInteractive = $false
)

$ErrorActionPreference = "Stop"

$DesktopPath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop"
$FrontPageBase = Join-Path -Path $DesktopPath -ChildPath "FRONT_PAGE"

Write-Host "============================================================" -ForegroundColor Green
Write-Host "         Philippine Star Front Page Viewer & Launcher        " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Location: $FrontPageBase" -ForegroundColor Gray

if (-not (Test-Path $FrontPageBase)) {
    Write-Host "`n[ERROR] FRONT_PAGE directory not found at: $FrontPageBase" -ForegroundColor Red
    return
}

# --- STEP 1: Scan & Select Year ---
$yearDirs = Get-ChildItem -Path $FrontPageBase -Directory | Sort-Object Name

if ($yearDirs.Count -eq 0) {
    Write-Host "`n[WARNING] No year folders found inside '$FrontPageBase'." -ForegroundColor Yellow
    Write-Host "Please compile front pages first using Run-Compile-Frontpages.bat." -ForegroundColor Cyan
    return
}

$selectedYearDir = $null

if ($Year -ne "") {
    $cleanYear = $Year.Trim()
    $selectedYearDir = $yearDirs | Where-Object { $_.Name -eq $cleanYear } | Select-Object -First 1
}

if (-not $selectedYearDir) {
    Write-Host "`nAvailable Years:" -ForegroundColor Yellow
    $yearMap = @{}
    $yIdx = 1
    $defaultYearIdx = 1

    foreach ($yd in $yearDirs) {
        $mCount = (Get-ChildItem -Path $yd.FullName -Directory -ErrorAction SilentlyContinue).Count
        $pCount = (Get-ChildItem -Path $yd.FullName -Recurse -Filter "*.pdf" -ErrorAction SilentlyContinue).Count
        $isLast = if ($yIdx -eq $yearDirs.Count) { " (Default)" } else { "" }
        Write-Host "  [$yIdx] $($yd.Name) ($mCount months, $pCount pages)$isLast" -ForegroundColor White
        $yearMap["$yIdx"] = $yd
        $yearMap[$yd.Name] = $yd
        $defaultYearIdx = $yIdx
        $yIdx++
    }

    if (-not $NonInteractive) {
        $yInput = Read-Host "`nSelect Year [1-$($yearDirs.Count) or year number] (Default: $($yearDirs[$defaultYearIdx - 1].Name))"
        if ([string]::IsNullOrWhiteSpace($yInput)) {
            $selectedYearDir = $yearDirs[$defaultYearIdx - 1]
        } elseif ($yearMap.ContainsKey($yInput.Trim())) {
            $selectedYearDir = $yearMap[$yInput.Trim()]
        } else {
            Write-Host "`n[ERROR] Invalid year selection: '$yInput'" -ForegroundColor Red
            return
        }
    } else {
        $selectedYearDir = $yearDirs[$defaultYearIdx - 1]
    }
}

Write-Host " Selected Year: $($selectedYearDir.Name)" -ForegroundColor Cyan

# --- STEP 2: Scan & Select Month ---
$monthDirs = Get-ChildItem -Path $selectedYearDir.FullName -Directory | Sort-Object Name

if ($monthDirs.Count -eq 0) {
    Write-Host "`n[WARNING] No month folders found inside '$($selectedYearDir.FullName)'." -ForegroundColor Yellow
    return
}

$selectedMonthDir = $null

if ($Month -ne "") {
    $cleanMonth = $Month.Trim()
    $selectedMonthDir = $monthDirs | Where-Object { 
        $_.Name -eq $cleanMonth -or $_.Name -like "*$cleanMonth*" -or ($_.Name -match "^0?(\d+)" -and $Matches[1] -eq $cleanMonth)
    } | Select-Object -First 1
}

if (-not $selectedMonthDir) {
    Write-Host "`nAvailable Months in $($selectedYearDir.Name):" -ForegroundColor Yellow
    $monthMap = @{}
    $mIdx = 1
    $defaultMonthIdx = $monthDirs.Count

    foreach ($md in $monthDirs) {
        $pdfCount = (Get-ChildItem -Path $md.FullName -Filter "*.pdf" -ErrorAction SilentlyContinue).Count
        $falseCount = (Get-ChildItem -Path $md.FullName -Filter "*_FALSE*.pdf" -ErrorAction SilentlyContinue).Count
        $falseInfo = if ($falseCount -gt 0) { " (includes $falseCount FALSE wrap)" } else { "" }
        $isLast = if ($mIdx -eq $monthDirs.Count) { " (Default)" } else { "" }

        # Extract month number if starts with digits
        $mNum = if ($md.Name -match "^(\d{1,2})") { [int]$Matches[1] } else { $mIdx }
        Write-Host "  [$mNum] $($md.Name) ($pdfCount pages$falseInfo)$isLast" -ForegroundColor White

        $monthMap["$mIdx"] = $md
        $monthMap["$mNum"] = $md
        $monthMap[$md.Name] = $md
        $mIdx++
    }

    if (-not $NonInteractive) {
        $defChoice = if ($monthDirs[$defaultMonthIdx - 1].Name -match "^(\d{1,2})") { [int]$Matches[1] } else { $defaultMonthIdx }
        $mInput = Read-Host "`nSelect Month [month number or name] (Default: $defChoice)"
        if ([string]::IsNullOrWhiteSpace($mInput)) {
            $selectedMonthDir = $monthDirs[$defaultMonthIdx - 1]
        } elseif ($monthMap.ContainsKey($mInput.Trim())) {
            $selectedMonthDir = $monthMap[$mInput.Trim()]
        } else {
            Write-Host "`n[ERROR] Invalid month selection: '$mInput'" -ForegroundColor Red
            return
        }
    } else {
        $selectedMonthDir = $monthDirs[$defaultMonthIdx - 1]
    }
}

Write-Host " Selected Month: $($selectedMonthDir.Name)" -ForegroundColor Cyan

# --- STEP 3: Scan Pages in Month ---
$pages = Get-ChildItem -Path $selectedMonthDir.FullName -Filter "*.pdf" | Sort-Object Name

if ($pages.Count -eq 0) {
    Write-Host "`n[WARNING] No PDF pages found in '$($selectedMonthDir.FullName)'." -ForegroundColor Yellow
    return
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " Found $($pages.Count) page(s) in $($selectedYearDir.Name)\$($selectedMonthDir.Name):" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Green

foreach ($p in $pages) {
    $tag = if ($p.Name -match "_FALSE") { " [FALSE COVER]" } else { "" }
    $color = if ($p.Name -match "_FALSE") { "Magenta" } else { "Gray" }
    Write-Host "  * $($p.Name)$tag" -ForegroundColor $color
}

# --- STEP 4: Launch Action ---
if (-not $NonInteractive) {
    Write-Host "`nAction:" -ForegroundColor Yellow
    Write-Host "  [1] Launch ALL $($pages.Count) pages in default viewer (Default)" -ForegroundColor White
    Write-Host "  [2] Open month folder in File Explorer" -ForegroundColor White
    Write-Host "  [3] Both (Launch all pages AND open folder)" -ForegroundColor White
    Write-Host "  [Q] Exit without launching" -ForegroundColor DarkGray

    $actChoice = Read-Host "`nChoose action [1/2/3/Q] (Default: 1)"
    if ($actChoice -match "^[Qq]") {
        Write-Host "`nExited without launching." -ForegroundColor Yellow
        return
    }
    if ($actChoice -eq "2") {
        Write-Host "`nOpening folder in File Explorer..." -ForegroundColor Cyan
        Invoke-Item -Path $selectedMonthDir.FullName
        return
    }
    if ($actChoice -eq "3") {
        $OpenFolder = $true
    }
}

Write-Host "`nLaunching all $($pages.Count) page(s) in default PDF viewer..." -ForegroundColor Green

$opened = 0
foreach ($p in $pages) {
    $opened++
    $tag = if ($p.Name -match "_FALSE") { " (FALSE COVER)" } else { "" }
    Write-Host "  [$opened/$($pages.Count)] Opening $($p.Name)$tag..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath $p.FullName
    } catch {
        Invoke-Item -Path $p.FullName
    }
    # Slight pause to let Edge/viewer open tabs smoothly
    Start-Sleep -Milliseconds 150
}

if ($OpenFolder) {
    Invoke-Item -Path $selectedMonthDir.FullName
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host " Successfully launched all $($pages.Count) page(s) for $($selectedYearDir.Name)\$($selectedMonthDir.Name)!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
