<#
Get-TopDiskItems.ps1

Lists the Top-N largest files (and optional folders) on one or more fixed disks.

Updates:
- Input accepts 'all' to scan every listed disk
- Ranges like '1-4' or '1 - 4'
- Comma lists with or without spaces: '1,3,5' or '1, 3, 5'
- Mixed forms: '1-3, 5, 7-8'

Usage examples:
  .\Get-TopDiskItems.ps1
  .\Get-TopDiskItems.ps1 -Top 50 -IncludeFolders -Export
#>
[CmdletBinding()]
param(
  [int]$Top = 20,
  [switch]$IncludeFolders,
  [switch]$Export
)

function Format-Bytes {
  [CmdletBinding()] param([long]$Bytes)
  if ($Bytes -ge 1PB) { return "{0:N2} PB" -f ($Bytes / 1PB) }
  if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
  if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
  if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
  if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
  return "$Bytes B"
}

# Enumerate fixed logical disks
$drives = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
          Sort-Object DeviceID |
          Select-Object DeviceID, VolumeName, Size, FreeSpace

if (-not $drives) { Write-Warning "No fixed disks found."; exit 1 }

Write-Host "`n== Fixed Disks ==" -ForegroundColor Cyan
$index = 1
$drivesList = foreach ($d in $drives) {
  [pscustomobject]@{
    Index = $index++
    Drive = $d.DeviceID
    Label = if ($d.VolumeName) { $d.VolumeName } else { "(No Label)" }
    Size  = Format-Bytes ([long]$d.Size)
    Free  = Format-Bytes ([long]$d.FreeSpace)
    Used  = Format-Bytes (([long]$d.Size) - ([long]$d.FreeSpace))
  }
}
$drivesList | Format-Table -AutoSize

function Parse-Selection {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$InputString,
    [Parameter(Mandatory)][int]$MaxIndex
  )
  $s = $InputString.Trim().ToLower()
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  if ($s -eq 'all') { return 1..$MaxIndex }

  # Normalize common separators and spaced ranges
  $s = $s -replace ';', ',' -replace ' - ', '-'
  $tokens = $s -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

  $chosen = [System.Collections.Generic.HashSet[int]]::new()
  foreach ($t in $tokens) {
    if ($t -eq 'all') { return 1..$MaxIndex }
    elseif ($t -match '^(\d+)-(\d+)$') {
      $a = [int]$Matches[1]; $b = [int]$Matches[2]
      if ($a -gt $b) { $tmp=$a; $a=$b; $b=$tmp }
      foreach ($i in $a..$b) { if ($i -ge 1 -and $i -le $MaxIndex) { $null = $chosen.Add($i) } }
    }
    elseif ($t -match '^\d+$') {
      $n = [int]$t
      if ($n -ge 1 -and $n -le $MaxIndex) { $null = $chosen.Add($n) }
    }
  }
  return $chosen.ToArray() | Sort-Object
}

function Get-TopFiles {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Root,
    [int]$Top = 20
  )
  Write-Host "Scanning files under $Root..." -ForegroundColor Yellow
  Get-ChildItem -LiteralPath $Root -File -Recurse -Force -ErrorAction SilentlyContinue |
    Sort-Object Length -Descending |
    Select-Object -First $Top @{n='SizeBytes';e={$_.Length}}, @{n='Size';e={Format-Bytes $_.Length}}, FullName, DirectoryName, LastWriteTime
}

function Get-TopFolders {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Root,
    [int]$Top = 20
  )
  Write-Host "Measuring folder sizes under $Root (this may take time)..." -ForegroundColor Yellow
  $dirs = Get-ChildItem -LiteralPath $Root -Directory -Force -ErrorAction SilentlyContinue
  $i = 0; $total = [math]::Max($dirs.Count,1)
  $rows = foreach ($dir in $dirs) {
    $i++; $pct = [int](($i/$total)*100)
    Write-Progress -Activity "Folder sizing" -Status "$pct% ($i of $total)" -PercentComplete $pct
    $sum = (Get-ChildItem -LiteralPath $dir.FullName -File -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    [pscustomobject]@{
      Folder    = $dir.FullName
      SizeBytes = [long]$sum
      Size      = Format-Bytes $sum
      LastWrite = $dir.LastWriteTime
    }
  }
  Write-Progress -Activity "Folder sizing" -Completed
  $rows | Sort-Object SizeBytes -Descending | Select-Object -First $Top
}

# Prompt and parse selection
$raw = Read-Host "`nEnter drive numbers (e.g., 1,3,5), ranges (e.g., 1-4), or 'all'"
$chosen = Parse-Selection -InputString $raw -MaxIndex $drivesList.Count
if (-not $chosen) { Write-Warning "No valid selections."; exit 1 }

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$downloads = Join-Path $env:USERPROFILE 'Downloads'

foreach ($n in $chosen) {
  $drive = $drivesList[$n-1]
  $root  = "$($drive.Drive)\\"

  Write-Host "`n=== $($drive.Drive)  $($drive.Label) ===" -ForegroundColor Cyan

  # Files
  $topFiles = Get-TopFiles -Root $root -Top $Top
  Write-Host "`n-- Top $Top Files --" -ForegroundColor Green
  $topFiles | Select-Object @{n='Rank';e={[array]::IndexOf($topFiles, $_)+1}}, Size, FullName, DirectoryName, LastWriteTime | Format-Table -AutoSize

  if ($Export) {
    $fileCsv = Join-Path $downloads ("Top{0}Files_{1}_{2}.csv" -f $Top, $drive.Drive.TrimEnd(':'), $timestamp)
    $topFiles | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $fileCsv
    Write-Host "Exported: $fileCsv" -ForegroundColor DarkCyan
  }

  # Folders (optional)
  if ($IncludeFolders) {
    $topFolders = Get-TopFolders -Root $root -Top $Top
    Write-Host "`n-- Top $Top Folders --" -ForegroundColor Green
    $topFolders | Select-Object @{n='Rank';e={[array]::IndexOf($topFolders, $_)+1}}, Size, Folder, LastWrite | Format-Table -AutoSize

    if ($Export) {
      $folderCsv = Join-Path $downloads ("Top{0}Folders_{1}_{2}.csv" -f $Top, $drive.Drive.TrimEnd(':'), $timestamp)
      $topFolders | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $folderCsv
      Write-Host "Exported: $folderCsv" -ForegroundColor DarkCyan
    }
  }
}

Write-Host "`nDone." -ForegroundColor Cyan
