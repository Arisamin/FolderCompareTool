Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to select a folder using GUI (modern dialog that supports MTP devices)
function Select-Folder {
    param([string]$Description)
    
    # Use the Windows Shell COM object for modern folder browser
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.BrowseForFolder(0, $Description, 0x0200, 0)
    
    if ($folder) {
        $selectedPath = $folder.Self.Path
        
        # Return both the path and the folder object for MTP devices
        $result = @{
            Path = $selectedPath
            FolderObject = $folder
        }
        
        return $result
    }
    
    # Release COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    return $null
}

# Function to check if path is MTP device
function Test-IsMTPPath {
    param([string]$Path)
    return $Path -match '^::' -or $Path -match '^\\\\\?\\usb'
}

# Function to get items from Shell namespace (for MTP devices)
function Get-ShellItems {
    param(
        [Parameter(Mandatory=$true)]
        $FolderObject,
        [string]$BasePath = "",
        [ref]$AllFiles,
        [ref]$AllFolders
    )
    
    foreach ($item in $FolderObject.Items()) {
        $itemPath = if ($BasePath) { "$BasePath\$($item.Name)" } else { $item.Name }
        
        if ($item.IsFolder) {
            $AllFolders.Value += @{
                FullName = $itemPath
                Name = $item.Name
            }
            
            # Recurse into subfolder
            $subFolder = $item.GetFolder
            if ($subFolder) {
                Get-ShellItems -FolderObject $subFolder -BasePath $itemPath -AllFiles $AllFiles -AllFolders $AllFolders
            }
        } else {
            $size = 0
            
            # Try multiple methods to get file size
            try {
                # Method 1: Try to get Size property directly
                if ($item.Size -and $item.Size -gt 0) {
                    $size = $item.Size
                }
            } catch { }
            
            if ($size -eq 0) {
                try {
                    # Method 2: Try ExtendedProperty
                    $sizeProperty = $item.ExtendedProperty("System.Size")
                    if ($sizeProperty -and $sizeProperty -gt 0) {
                        $size = [long]$sizeProperty
                    }
                } catch { }
            }
            
            if ($size -eq 0) {
                try {
                    # Method 3: GetDetailsOf with column index 1 (usually Size)
                    $sizeStr = $FolderObject.GetDetailsOf($item, 1)
                    if ($sizeStr) {
                        # Try to extract number with units
                        if ($sizeStr -match '([\d,]+)\s*(KB|MB|GB)') {
                            $numStr = $matches[1] -replace ',', ''
                            $unit = $matches[2]
                            $num = [double]$numStr
                            
                            # Convert to bytes based on unit
                            switch ($unit) {
                                'KB' { $size = [long]($num * 1024) }
                                'MB' { $size = [long]($num * 1024 * 1024) }
                                'GB' { $size = [long]($num * 1024 * 1024 * 1024) }
                            }
                        }
                        # Try plain number (bytes)
                        elseif ($sizeStr -match '^[\d,]+$') {
                            $size = [long]($sizeStr -replace ',', '')
                        }
                    }
                } catch { }
            }
            
            $AllFiles.Value += @{
                FullName = $itemPath
                Name = $item.Name
                Length = $size
            }
        }
    }
}

# Function to get folder statistics (works with both regular paths and MTP devices)
function Get-FolderStats {
    param(
        [string]$Path,
        $FolderObject = $null
    )
    
    $isMTP = Test-IsMTPPath -Path $Path
    
    if ($isMTP) {
        # Use Shell COM object for MTP devices
        Write-Host "  Detected MTP/Portable device, using Shell namespace..." -ForegroundColor Gray
        
        if (-not $FolderObject) {
            Write-Host "  Error: No folder object provided for MTP device" -ForegroundColor Red
            return @{
                Path = $Path
                TotalSize = 0
                TotalSizeGB = 0
                TotalSizeMB = 0
                FileCount = 0
                FolderCount = 0
                Files = @()
                Folders = @()
            }
        }
        
        $allFiles = @()
        $allFolders = @()
        
        Get-ShellItems -FolderObject $FolderObject -AllFiles ([ref]$allFiles) -AllFolders ([ref]$allFolders)
        
        $totalSize = ($allFiles | ForEach-Object { $_.Length } | Measure-Object -Sum).Sum
        if ($null -eq $totalSize) { $totalSize = 0 }
        
        return @{
            Path = $Path
            TotalSize = $totalSize
            TotalSizeGB = [math]::Round($totalSize / 1GB, 2)
            TotalSizeMB = [math]::Round($totalSize / 1MB, 2)
            FileCount = $allFiles.Count
            FolderCount = $allFolders.Count
            Files = $allFiles
            Folders = $allFolders
        }
    } else {
        # Use regular file system cmdlets for normal paths
        $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
        $files = $items | Where-Object { -not $_.PSIsContainer }
        $folders = $items | Where-Object { $_.PSIsContainer }
        
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $totalSize) { $totalSize = 0 }
        
        return @{
            Path = $Path
            TotalSize = $totalSize
            TotalSizeGB = [math]::Round($totalSize / 1GB, 2)
            TotalSizeMB = [math]::Round($totalSize / 1MB, 2)
            FileCount = $files.Count
            FolderCount = $folders.Count
            Files = $files
            Folders = $folders
        }
    }
}

# Function to get relative paths for comparison
function Get-RelativePaths {
    param(
        [array]$Items,
        [string]$BasePath
    )
    
    $relativePaths = @()
    $isMTP = Test-IsMTPPath -Path $BasePath
    
    foreach ($item in $Items) {
        if ($isMTP) {
            # For MTP devices, the FullName is already relative
            $relativePaths += $item.FullName
        } else {
            # For regular paths, make it relative
            $relativePath = $item.FullName.Substring($BasePath.Length).TrimStart('\')
            $relativePaths += $relativePath
        }
    }
    return $relativePaths
}

# Main script
Write-Host "Folder Comparison Tool" -ForegroundColor Cyan
Write-Host "=====================" -ForegroundColor Cyan
Write-Host ""

# Select Folder A
Write-Host "Please select Folder A..." -ForegroundColor Yellow
$folderAResult = Select-Folder -Description "Select Folder A"
if (-not $folderAResult) {
    Write-Host "Folder A selection cancelled. Exiting." -ForegroundColor Red
    exit
}
$folderA = $folderAResult.Path
$folderAObject = $folderAResult.FolderObject
Write-Host "Folder A: $folderA" -ForegroundColor Green

# Select Folder B
Write-Host "Please select Folder B..." -ForegroundColor Yellow
$folderBResult = Select-Folder -Description "Select Folder B"
if (-not $folderBResult) {
    Write-Host "Folder B selection cancelled. Exiting." -ForegroundColor Red
    exit
}
$folderB = $folderBResult.Path
$folderBObject = $folderBResult.FolderObject
Write-Host "Folder B: $folderB" -ForegroundColor Green
Write-Host ""

# Gather statistics
Write-Host "Analyzing folders..." -ForegroundColor Yellow
$statsA = Get-FolderStats -Path $folderA -FolderObject $folderAObject
$statsB = Get-FolderStats -Path $folderB -FolderObject $folderBObject

# Display statistics
Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "FOLDER STATISTICS" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Folder A: $($statsA.Path)" -ForegroundColor White
Write-Host "  Total Size: $($statsA.TotalSizeMB) MB ($($statsA.TotalSizeGB) GB)" -ForegroundColor White
Write-Host "  Files: $($statsA.FileCount)" -ForegroundColor White
Write-Host "  Folders: $($statsA.FolderCount)" -ForegroundColor White
Write-Host ""

Write-Host "Folder B: $($statsB.Path)" -ForegroundColor White
Write-Host "  Total Size: $($statsB.TotalSizeMB) MB ($($statsB.TotalSizeGB) GB)" -ForegroundColor White
Write-Host "  Files: $($statsB.FileCount)" -ForegroundColor White
Write-Host "  Folders: $($statsB.FolderCount)" -ForegroundColor White
Write-Host ""

# Compare files
Write-Host "Comparing contents..." -ForegroundColor Yellow
$filesA = Get-RelativePaths -Items $statsA.Files -BasePath $folderA
$filesB = Get-RelativePaths -Items $statsB.Files -BasePath $folderB
$foldersA = Get-RelativePaths -Items $statsA.Folders -BasePath $folderA
$foldersB = Get-RelativePaths -Items $statsB.Folders -BasePath $folderB

# Find differences
$filesOnlyInA = $filesA | Where-Object { $_ -notin $filesB }
$filesOnlyInB = $filesB | Where-Object { $_ -notin $filesA }
$foldersOnlyInA = $foldersA | Where-Object { $_ -notin $foldersB }
$foldersOnlyInB = $foldersB | Where-Object { $_ -notin $foldersA }

# Display differences
Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "COMPARISON RESULTS" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "Files only in Folder A ($($filesOnlyInA.Count)):" -ForegroundColor Magenta
if ($filesOnlyInA.Count -gt 0) {
    foreach ($file in $filesOnlyInA | Sort-Object) {
        Write-Host "  + $file" -ForegroundColor Yellow
    }
} else {
    Write-Host "  (none)" -ForegroundColor Gray
}
Write-Host ""

Write-Host "Files only in Folder B ($($filesOnlyInB.Count)):" -ForegroundColor Magenta
if ($filesOnlyInB.Count -gt 0) {
    foreach ($file in $filesOnlyInB | Sort-Object) {
        Write-Host "  + $file" -ForegroundColor Yellow
    }
} else {
    Write-Host "  (none)" -ForegroundColor Gray
}
Write-Host ""

Write-Host "Folders only in Folder A ($($foldersOnlyInA.Count)):" -ForegroundColor Magenta
if ($foldersOnlyInA.Count -gt 0) {
    foreach ($folder in $foldersOnlyInA | Sort-Object) {
        Write-Host "  + $folder" -ForegroundColor Cyan
    }
} else {
    Write-Host "  (none)" -ForegroundColor Gray
}
Write-Host ""

Write-Host "Folders only in Folder B ($($foldersOnlyInB.Count)):" -ForegroundColor Magenta
if ($foldersOnlyInB.Count -gt 0) {
    foreach ($folder in $foldersOnlyInB | Sort-Object) {
        Write-Host "  + $folder" -ForegroundColor Cyan
    }
} else {
    Write-Host "  (none)" -ForegroundColor Gray
}
Write-Host ""

# Summary
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
$commonFiles = $filesA.Count - $filesOnlyInA.Count
$commonFolders = $foldersA.Count - $foldersOnlyInA.Count
Write-Host "Common files: $commonFiles" -ForegroundColor Green
Write-Host "Common folders: $commonFolders" -ForegroundColor Green
Write-Host "Files unique to A: $($filesOnlyInA.Count)" -ForegroundColor Yellow
Write-Host "Files unique to B: $($filesOnlyInB.Count)" -ForegroundColor Yellow
Write-Host "Folders unique to A: $($foldersOnlyInA.Count)" -ForegroundColor Yellow
Write-Host "Folders unique to B: $($foldersOnlyInB.Count)" -ForegroundColor Yellow
Write-Host ""

# Option to export results
$export = Read-Host "Would you like to export results to a file? (Y/N)"
if ($export -eq 'Y' -or $export -eq 'y') {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportPath = Join-Path $PSScriptRoot "FolderComparison_$timestamp.txt"
    
    $report = @"
FOLDER COMPARISON REPORT
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
===============================================

FOLDER STATISTICS
===============================================
Folder A: $($statsA.Path)
  Total Size: $($statsA.TotalSizeMB) MB ($($statsA.TotalSizeGB) GB)
  Files: $($statsA.FileCount)
  Folders: $($statsA.FolderCount)

Folder B: $($statsB.Path)
  Total Size: $($statsB.TotalSizeMB) MB ($($statsB.TotalSizeGB) GB)
  Files: $($statsB.FileCount)
  Folders: $($statsB.FolderCount)

===============================================
COMPARISON RESULTS
===============================================

Files only in Folder A ($($filesOnlyInA.Count)):
$($filesOnlyInA | Sort-Object | ForEach-Object { "  + $_" } | Out-String)

Files only in Folder B ($($filesOnlyInB.Count)):
$($filesOnlyInB | Sort-Object | ForEach-Object { "  + $_" } | Out-String)

Folders only in Folder A ($($foldersOnlyInA.Count)):
$($foldersOnlyInA | Sort-Object | ForEach-Object { "  + $_" } | Out-String)

Folders only in Folder B ($($foldersOnlyInB.Count)):
$($foldersOnlyInB | Sort-Object | ForEach-Object { "  + $_" } | Out-String)

===============================================
SUMMARY
===============================================
Common files: $commonFiles
Common folders: $commonFolders
Files unique to A: $($filesOnlyInA.Count)
Files unique to B: $($filesOnlyInB.Count)
Folders unique to A: $($foldersOnlyInA.Count)
Folders unique to B: $($foldersOnlyInB.Count)
"@
    
    $report | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "Report saved to: $reportPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "Comparison complete!" -ForegroundColor Green
