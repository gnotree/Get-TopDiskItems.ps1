# Get-TopDiskItems.ps1

A PowerShell script that lists the largest files and folders on one or more disks, helping you quickly locate disk-space hogs. It automatically enumerates all fixed drives, lets you pick which to scan (e.g., `1,3,4`), and outputs results in ranked order, optionally exporting to CSV.

---

## Features

* Enumerates all fixed disks automatically
* Interactive drive selection (e.g., `1,5,6`)
* Top-N largest files (default 20)
* Optional folder size analysis (`-IncludeFolders`)
* Optional CSV export (`-Export` saves to your Downloads folder)
* Human-readable sizes (GB, MB, etc.)
* Progress bar for folder scanning
* Error-tolerant recursion (skips inaccessible paths safely)

---

## Usage

### 1. Run in PowerShell

Open PowerShell (version 7 or later recommended), then execute:

```powershell
.\Get-TopDiskItems.ps1
```

### 2. Follow prompts

The script will list available fixed drives:

```
== Fixed Disks ==
Index Drive Label       Size     Free     Used
----- ----- ----------- -------- -------- --------
1     C:    System      500 GB   150 GB   350 GB
2     D:    Storage     1 TB     450 GB   550 GB
```

When prompted, enter:

```
Enter drive numbers to scan (comma-separated, e.g. 1,3,4): 1,2
```

The script will then display the top 20 largest files for each selected drive.

---

## Parameters

| Parameter         | Description                                         | Default |
| ----------------- | --------------------------------------------------- | ------- |
| `-Top`            | Number of top items to display                      | 20      |
| `-IncludeFolders` | Include largest folders by cumulative size (slower) | Off     |
| `-Export`         | Save CSV reports to Downloads folder                | Off     |

---

## Examples

### Top 20 files (interactive)

```powershell
.\Get-TopDiskItems.ps1
```

### Include largest folders

```powershell
.\Get-TopDiskItems.ps1 -IncludeFolders
```

### Export results to CSV

```powershell
.\Get-TopDiskItems.ps1 -Export
```

### Top 50 files and folders, with CSV export

```powershell
.\Get-TopDiskItems.ps1 -Top 50 -IncludeFolders -Export
```

---

## Output Format

### Console Output

```
-- Top 20 Files on C: --
Rank Size      FullName                              DirectoryName                       LastWriteTime
---- ----      --------                              -------------                       --------------
1    42.3 GB   C:\VMs\Ubuntu.vhdx                    C:\VMs                               2025-10-22
2    18.7 GB   C:\Users\Public\Videos\OBS_Recording.mkv  ...
```

### CSV Export (if `-Export`)

```
C:\Users\<username>\Downloads\Top20Files_C_20251103_1918.csv
C:\Users\<username>\Downloads\Top20Folders_D_20251103_1918.csv
```

---

## Notes

* Folder size analysis is recursive and may take longer on large drives.
* Uses `Get-ChildItem -Recurse -Force -ErrorAction SilentlyContinue` to safely skip restricted paths.
* Works on Windows 10, 11, and Windows Server 2022 or later.
* PowerShell 7.x provides faster performance and improved formatting.

---

## Requirements

* PowerShell 5.1 or newer (7.x recommended)
* No administrator rights required, though elevated mode improves access.

---

## License

MIT License - Free to use, modify, and distribute.

---

## Author

Grant Scott Turner
[https://gtai.dev](https://gtai.dev)
Engineering and Cybersecurity Student | Domain Admin | Blue/Purple Team Developer
