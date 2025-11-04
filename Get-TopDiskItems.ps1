<# 
.SYNOPSIS
  List the top-N largest items on selected disks.

.DESCRIPTION
  - Enumerates fixed disks, shows an indexed list.
  - Prompts for comma-separated selection (e.g., 1,3).
  - Outputs the top largest FILES (fast).
  - Optional: include top largest FOLDERS by total size (slower) via -IncludeFolders.
  - Optional: export per-drive CSVs to ~/Downloads via -Export.

.PARAMETER Top
  Number of items to show (default 20).

.PARAMETER IncludeFolders
  Also compute top folders by cumulative size (slower).

.PARAMETER Export
  Export CSVs to the user's Downloads folder.

.EXAMPLE
  .\Get-TopDiskItems.ps1
  (interactive; top 20 files)

.EXAMPLE
  .\Get-TopDiskItems.ps1 -Top 50 -IncludeFolders -Export
#>

[CmdletBinding()]
param(
  [int]$Top = 20,
  [switch]$IncludeFolders,
  [switch]$Export
)

function Format-Bytes {
  param([long]$Bytes)
  if ($Bytes -ge 1PB) { return "{0:N2} PB" -f ($Bytes / 1PB) }
  if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
  if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
  if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
  if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
  return "$Bytes B"
}

# 1) Enumerate fixed disks (DriveType=3)
$drives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
  Sort-Object DeviceID | Select-Object DeviceID, VolumeName, Size, FreeSpace

if (-not $drives) {
  Write-Warning "No fixed disks found."
  exit 1
}

# Show indexed list
Write-Host "`n== Fixed Disks ==" -ForegroundColor Cyan
$idx = 1
$display = foreach ($d in $drives) {
  [pscustomobject]@{
    Index      = $idx++
    Drive      = $d.DeviceID
    Label      = if ($d.VolumeName) {$d.VolumeName} else {"(No Label)"}
    Size       = Format-Bytes $d.Size
    Free       = Format-Bytes $d.FreeSpace
    Used       = Format-Bytes ([long]$d.Size - [long]$d.FreeSpace)
  }
}
$display | Format-Table -AutoSize

# 2) Prompt for selection
$raw = Read-Host "`nEnter drive numbers to scan (comma-separated, e.g. 1,3,4)"
if ([string]::IsNullOrWhiteSpace($raw)) {
  Write-Warning "No selection provided. Exiting."
  exit 1
}

$chosen = $raw -split '[,\s]+' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ } |
          Where-Object { $_ -ge 1 -and $_ -le $drives.Count } | Select-Object -Unique

if (-not $chosen) {
  Write-Warning "No valid selections. Exiting."
  exit 1
}

# Helper: top files
function Get-TopFiles {
  param(
    [Parameter(Mandatory)][string]$RootPath,
    [int]$Top = 20
  )
  Write-Host "Scanning files under $RootPath (this can take time)..." -ForegroundColor Yellow
  Get-ChildItem -LiteralPath $RootPath -File -Recurse -Force -ErrorAction SilentlyContinue |
    Sort-Object Length -Descending |
    Select-Object -First $Top @{
      Name='SizeBytes';Expression={$_.Length}
    }, @{
      Name='Size';Expression={ Format-Bytes $_.Length }
    }, FullName, DirectoryName, LastWriteTime
}

# Helper: top folders by cumulative size (slower)
function Get-TopFolders {
  param(
    [Parameter(Mandatory)][string]$RootPath,
    [int]$Top = 20
  )

  Write-Host "Calculating folder sizes under $RootPath (slower)..." -ForegroundColor Yellow

  $dirs = Get-ChildItem -LiteralPath $RootPath -Directory -Force -ErrorAction SilentlyContinue

  # Progress-friendly pass
  $i = 0
  $total = $dirs.Count
  $results = foreach ($dir in $dirs) {
    $i++
    $pct = [int](($i / [math]::Max($total,1)) * 100)
    Write-Progress -Activity "Measuring folders" -Status "$pct% ($i of $total)" -PercentComplete $pct

    $sum = (Get-ChildItem -LiteralPath $dir.FullName -File -Recurse -Force -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
    [pscustomobject]@{
      Folder     = $dir.FullName
      SizeBytes  = [long]($sum)
      Size       = Format-Bytes $sum
      LastWrite  = $dir.LastWriteTime
    }
  }
  Write-Progress -Activity "Measuring folders" -Completed

  $results | Sort-Object SizeBytes -Descending | Select-Object -First $Top
}

# 3) For each selected drive, run scans
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$downloads = Join-Path $env:USERPROFILE "Downloads"

foreach ($n in $chosen) {
  $drive = $drives[$n-1]
  $root  = ($drive.DeviceID + '\')

  Write-Host "`n=== $($drive.DeviceID)  $($drive.VolumeName) ===" -ForegroundColor Cyan

  # Top Files
  $topFiles = Get-TopFiles -RootPath $root -Top $Top
  Write-Host "`n-- Top $Top Files on $($drive.DeviceID) --" -ForegroundColor Green
  $topFiles | Select-Object @{n='Rank';e={[array]::IndexOf($topFiles, $_)+1}},
                         Size, FullName, DirectoryName, LastWriteTime |
            Format-Table -AutoSize

  if ($Export) {
    $fileCsv = Join-Path $downloads ("Top{0}Files_{1}_{2}.csv" -f $Top, $drive.DeviceID.TrimEnd(':'), $timestamp)
    $topFiles | Export-Csv -NoTypeInformation -Encoding UTF8 $fileCsv
    Write-Host "Exported: $fileCsv" -ForegroundColor DarkCyan
  }

  # Top Folders (optional)
  if ($IncludeFolders) {
    $topFolders = Get-TopFolders -RootPath $root -Top $Top
    Write-Host "`n-- Top $Top Folders by Size on $($drive.DeviceID) --" -ForegroundColor Green
    $topFolders | Select-Object @{n='Rank';e={[array]::IndexOf($topFolders, $_)+1}},
                             Size, Folder, LastWrite |
                Format-Table -AutoSize

    if ($Export) {
      $folderCsv = Join-Path $downloads ("Top{0}Folders_{1}_{2}.csv" -f $Top, $drive.DeviceID.TrimEnd(':'), $timestamp)
      $topFolders | Export-Csv -NoTypeInformation -Encoding UTF8 $folderCsv
      Write-Host "Exported: $folderCsv" -ForegroundColor DarkCyan
    }
  }
}

Write-Host "`nDone." -ForegroundColor Cyan
