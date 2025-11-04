# Get-TopDiskItems.ps1
A clean PowerShell script that:  Lists all attached fixed disks with index numbers  Prompts you to enter which disks to scan (comma-separated like 1,5,6)  Recursively finds the 20 largest files on each selected disk (fast)  Optional flag to also compute 20 largest folders by total size (slower)
