Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to select a folder using GUI
function Select-Folder {
    param([string]$Description)
    
    # Use the Windows Shell COM object for modern folder browser
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.BrowseForFolder(0, $Description, 0x0200, 0)
    
    if ($folder) {
        $selectedPath = $folder.Self.Path
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        return $selectedPath
    }
    
    # Release COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    return $null
}

# Main script
Write-Host "Find Duplicate File Names Tool" -ForegroundColor Cyan
Write-Host "===============================" -ForegroundColor Cyan
Write-Host ""

# Select root folder
Write-Host "Please select the root folder to search..." -ForegroundColor Yellow
$rootFolder = Select-Folder -Description "Select Root Folder"
if (-not $rootFolder) {
    Write-Host "Folder selection cancelled. Exiting." -ForegroundColor Red
    exit
}
Write-Host "Root Folder: $rootFolder" -ForegroundColor Green
Write-Host ""

# Scan for files
Write-Host "Scanning for files..." -ForegroundColor Yellow
$startTime = Get-Date

# Get all files with progress
$allFiles = @()
$folderCount = 0
$fileCount = 0

# First pass: count folders for better progress estimation
Write-Progress -Activity "Initializing scan..." -Status "Counting folders" -PercentComplete 0
$allFolders = @(Get-ChildItem -Path $rootFolder -Recurse -Directory -Force -ErrorAction SilentlyContinue)
$totalFolders = $allFolders.Count + 1  # +1 for root folder

Write-Progress -Activity "Scanning files" -Status "Processing folders..." -PercentComplete 0

# Scan files with progress
$allFiles = Get-ChildItem -Path $rootFolder -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
    $fileCount++
    if ($fileCount % 100 -eq 0) {
        $percent = [math]::Min(99, [math]::Round(($fileCount / [math]::Max(1, $fileCount)) * 100))
        $elapsed = (Get-Date) - $startTime
        Write-Progress -Activity "Scanning files" -Status "Found $fileCount files..." -PercentComplete $percent
    }
    $_
}

Write-Progress -Activity "Scanning files" -Completed

if ($allFiles.Count -eq 0) {
    Write-Host "No files found in the selected folder." -ForegroundColor Red
    exit
}

$scanElapsed = (Get-Date) - $startTime
Write-Host "Found $($allFiles.Count) files in $([math]::Round($scanElapsed.TotalSeconds, 1)) seconds. Analyzing..." -ForegroundColor Gray
Write-Host ""

# Group files by name (case-insensitive)
Write-Progress -Activity "Analyzing files" -Status "Grouping by name..." -PercentComplete 0
$groupedByName = $allFiles | Group-Object -Property Name
Write-Progress -Activity "Analyzing files" -Status "Finding duplicates..." -PercentComplete 50

# Find duplicates
$duplicates = $groupedByName | Where-Object { $_.Count -gt 1 }
Write-Progress -Activity "Analyzing files" -Completed

# Prepare display data (without printing yet)
$displayData = @()
if ($duplicates.Count -gt 0) {
    $processedCount = 0
    $totalDuplicates = $duplicates.Count
    
    foreach ($duplicate in $duplicates | Sort-Object Name) {
        $processedCount++
        $percent = [math]::Round(($processedCount / $totalDuplicates) * 100)
        
        Write-Progress -Activity "Preparing results" -Status "Processing $processedCount of $totalDuplicates duplicates" -PercentComplete $percent
        
        $fileName = $duplicate.Name
        $count = $duplicate.Count
        $files = $duplicate.Group
        
        # Calculate total size for this duplicate set
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
        
        $fileData = @{
            FileName = $fileName
            Count = $count
            TotalSize = $totalSize
            Locations = @()
        }
        
        foreach ($file in $files | Sort-Object FullName) {
            $relativePath = $file.FullName.Substring($rootFolder.Length).TrimStart('\')
            $sizeKB = [math]::Round($file.Length / 1KB, 2)
            $sizeMB = [math]::Round($file.Length / 1MB, 2)
            
            $fileData.Locations += @{
                Path = $relativePath
                SizeKB = $sizeKB
                SizeMB = $sizeMB
            }
        }
        
        $displayData += $fileData
    }
    
    # Sort by total size descending
    $displayData = $displayData | Sort-Object -Property TotalSize -Descending
    
    Write-Progress -Activity "Preparing results" -Completed
}

# Now display all results
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "SCAN RESULTS" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

if ($duplicates.Count -eq 0) {
    Write-Host "No duplicate file names found!" -ForegroundColor Green
} else {
    Write-Host "Found $($duplicates.Count) file name(s) with duplicates:" -ForegroundColor Magenta
    Write-Host ""
    
    foreach ($item in $displayData) {
        Write-Host "File Name: $($item.FileName)" -ForegroundColor Yellow
        Write-Host "  Occurrences: $($item.Count)" -ForegroundColor White
        Write-Host "  Locations:" -ForegroundColor White
        
        foreach ($location in $item.Locations) {
            if ($location.SizeMB -ge 1) {
                Write-Host "    - $($location.Path) ($($location.SizeMB) MB)" -ForegroundColor Cyan
            } else {
                Write-Host "    - $($location.Path) ($($location.SizeKB) KB)" -ForegroundColor Cyan
            }
        }
        Write-Host ""
    }
}

# Summary
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
$totalElapsed = (Get-Date) - $startTime
Write-Host "Total files scanned: $($allFiles.Count)" -ForegroundColor White
Write-Host "Unique file names: $($groupedByName.Count)" -ForegroundColor White
Write-Host "Duplicate file names: $($duplicates.Count)" -ForegroundColor White
if ($duplicates.Count -gt 0) {
    $totalDuplicateFiles = ($duplicates | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
    Write-Host "Total files with duplicate names: $totalDuplicateFiles" -ForegroundColor White
    
    # Calculate disk space that would be freed by keeping only one copy of each duplicate
    $spaceWasted = 0
    foreach ($duplicate in $duplicates) {
        $files = $duplicate.Group
        # Space wasted = (total size of all copies) - (size of one copy)
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
        $singleCopySize = $files[0].Length
        $spaceWasted += ($totalSize - $singleCopySize)
    }
    
    $spaceWastedMB = [math]::Round($spaceWasted / 1MB, 2)
    $spaceWastedGB = [math]::Round($spaceWasted / 1GB, 2)
    
    if ($spaceWastedGB -ge 1) {
        Write-Host "Disk space occupied by duplicates: $spaceWastedGB GB" -ForegroundColor Yellow
    } else {
        Write-Host "Disk space occupied by duplicates: $spaceWastedMB MB" -ForegroundColor Yellow
    }
    Write-Host "(This is the space you would free by keeping only one copy of each duplicate file)" -ForegroundColor Gray
}
Write-Host "Total run duration: $([math]::Round($totalElapsed.TotalSeconds, 1)) seconds" -ForegroundColor White
Write-Host ""

# Option to export results
$export = Read-Host "Would you like to export results to a file? (Y/N)"
if ($export -eq 'Y' -or $export -eq 'y') {
    $reportPath = Join-Path $PSScriptRoot "DuplicateNames.txt"
    
    # If file exists, add numbered suffix
    if (Test-Path $reportPath) {
        $counter = 1
        do {
            $reportPath = Join-Path $PSScriptRoot "DuplicateNames_$counter.txt"
            $counter++
        } while (Test-Path $reportPath)
    }
    
    $reportContent = @"
DUPLICATE FILE NAMES REPORT
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Root Folder: $rootFolder
===============================================

SUMMARY
===============================================
Total files scanned: $($allFiles.Count)
Unique file names: $($groupedByName.Count)
Duplicate file names: $($duplicates.Count)

"@

    if ($duplicates.Count -gt 0) {
        $totalDuplicateFiles = ($duplicates | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        $reportContent += "Total files with duplicate names: $totalDuplicateFiles`n"
        
        # Calculate disk space wasted
        $spaceWasted = 0
        foreach ($duplicate in $duplicates) {
            $files = $duplicate.Group
            $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
            $singleCopySize = $files[0].Length
            $spaceWasted += ($totalSize - $singleCopySize)
        }
        
        $spaceWastedMB = [math]::Round($spaceWasted / 1MB, 2)
        $spaceWastedGB = [math]::Round($spaceWasted / 1GB, 2)
        
        if ($spaceWastedGB -ge 1) {
            $reportContent += "Disk space occupied by duplicates: $spaceWastedGB GB`n"
        } else {
            $reportContent += "Disk space occupied by duplicates: $spaceWastedMB MB`n"
        }
        $reportContent += "(This is the space you would free by keeping only one copy of each duplicate file)`n`n"
        
        $reportContent += @"
===============================================
DUPLICATE FILES (Ordered by Total Size Descending)
===============================================

"@
        
        # Create sorted table
        $tableHeader = "{0,-80} {1,-12} {2,-15} {3,-15}" -f "FileName", "Occurrences", "TotalSizeMB", "TotalSizeGB"
        $tableSeparator = "-" * 122
        $reportContent += "$tableHeader`n"
        $reportContent += "$tableSeparator`n"
        
        foreach ($item in $displayData) {
            $totalSizeMB = [math]::Round($item.TotalSize / 1MB, 2)
            $totalSizeGB = [math]::Round($item.TotalSize / 1GB, 2)
            $reportContent += ("{0,-80} {1,-12} {2,-15} {3,-15}" -f $item.FileName, $item.Count, $totalSizeMB, $totalSizeGB) + "`n"
        }
        
        $reportContent += "`n"
        
        $reportContent += @"
===============================================
DUPLICATE FILES (Detailed View)
===============================================

"@
        
        $exportCount = 0
        $totalExport = $duplicates.Count
        
        foreach ($duplicate in $duplicates | Sort-Object Name) {
            $exportCount++
            if ($exportCount % 10 -eq 0 -or $exportCount -eq $totalExport) {
                $percent = [math]::Round(($exportCount / $totalExport) * 100)
                Write-Progress -Activity "Generating report" -Status "Processing $exportCount of $totalExport duplicates" -PercentComplete $percent
            }
            
            $fileName = $duplicate.Name
            $count = $duplicate.Count
            $files = $duplicate.Group
            
            $reportContent += "File Name: $fileName`n"
            $reportContent += "  Occurrences: $count`n"
            $reportContent += "  Locations:`n"
            
            foreach ($file in $files | Sort-Object FullName) {
                $relativePath = $file.FullName.Substring($rootFolder.Length).TrimStart('\')
                $sizeKB = [math]::Round($file.Length / 1KB, 2)
                $sizeMB = [math]::Round($file.Length / 1MB, 2)
                
                if ($sizeMB -ge 1) {
                    $reportContent += "    - $relativePath ($sizeMB MB)`n"
                } else {
                    $reportContent += "    - $relativePath ($sizeKB KB)`n"
                }
            }
            $reportContent += "`n"
        }
        
        Write-Progress -Activity "Generating report" -Completed
    }
    
    Write-Progress -Activity "Saving report" -Status "Writing to file..." -PercentComplete 0
    $reportContent | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Progress -Activity "Saving report" -Completed
    Write-Host "Report saved to: $reportPath" -ForegroundColor Green
    
    # Ask if user wants CSV export for automation
    if ($duplicates.Count -gt 0) {
        Write-Host ""
        $exportCsv = Read-Host "Would you like to export a CSV file for automated duplicate handling? (Y/N)"
        if ($exportCsv -eq 'Y' -or $exportCsv -eq 'y') {
            $csvPath = $reportPath -replace '\.txt$', '.csv'
            
            Write-Host ""
            Write-Host "CSV Export Options:" -ForegroundColor Cyan
            Write-Host "1. Export ALL duplicate files (including the first occurrence)" -ForegroundColor White
            Write-Host "2. Export ONLY extra copies (keep the first occurrence of each duplicate)" -ForegroundColor White
            $csvOption = Read-Host "Select option (1 or 2)"
            
            $csvData = @()
            $csvCount = 0
            $totalCsvItems = $duplicates.Count
            
            foreach ($duplicate in $duplicates) {
                $csvCount++
                if ($csvCount % 10 -eq 0 -or $csvCount -eq $totalCsvItems) {
                    $percent = [math]::Round(($csvCount / $totalCsvItems) * 100)
                    Write-Progress -Activity "Generating CSV" -Status "Processing $csvCount of $totalCsvItems duplicates" -PercentComplete $percent
                }
                
                $fileName = $duplicate.Name
                $files = $duplicate.Group | Sort-Object FullName
                
                # Determine which files to export
                $filesToExport = if ($csvOption -eq '1') {
                    $files  # All files
                } else {
                    $files | Select-Object -Skip 1  # Skip first, export rest
                }
                
                foreach ($file in $filesToExport) {
                    $csvData += [PSCustomObject]@{
                        FileName = $fileName
                        FullPath = $file.FullName
                        SizeBytes = $file.Length
                        SizeMB = [math]::Round($file.Length / 1MB, 2)
                        LastModified = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                        DuplicateCount = $duplicate.Count
                    }
                }
            }
            
            Write-Progress -Activity "Generating CSV" -Completed
            Write-Progress -Activity "Saving CSV" -Status "Writing to file..." -PercentComplete 0
            $csvData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Progress -Activity "Saving CSV" -Completed
            
            Write-Host "CSV saved to: $csvPath" -ForegroundColor Green
            Write-Host "Total rows: $($csvData.Count)" -ForegroundColor Gray
            Write-Host ""
            Write-Host "You can now pipe this CSV to a deletion script. Example:" -ForegroundColor Yellow
            Write-Host "  Import-Csv '$csvPath' | ForEach-Object { Remove-Item `$_.FullPath -WhatIf }" -ForegroundColor Cyan
            Write-Host "(Remove the -WhatIf parameter to actually delete files)" -ForegroundColor Gray
        }
    }
}

Write-Host ""
Write-Host "Scan complete!" -ForegroundColor Green
