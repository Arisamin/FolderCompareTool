# Folder Compare Tool

A collection of PowerShell scripts for folder comparison and duplicate file detection with GUI folder selection.

## Scripts Overview

- **Compare-Folders.ps1** - Compare two folders with bidirectional analysis, supports MTP devices (phones via USB)
- **Find-DuplicateNames.ps1** - Recursively find all files with duplicate names in a folder hierarchy

## Detailed Information

### Compare-Folders.ps1
Compare the contents of two folders with support for both regular Windows file system paths and MTP devices (phones, cameras connected via USB).

### Find-DuplicateNames.ps1
Find all files with duplicate names within a folder hierarchy, recursively scanning all subdirectories.

## Features

### Compare-Folders.ps1

- **GUI Folder Selection**: Native Windows folder browser dialogs for selecting folders
- **MTP Device Support**: Compare folders on phones and other portable devices connected via USB
- **Comprehensive Statistics**: 
  - Total size (MB and GB)
  - File and folder counts
  - Bidirectional comparison
- **Detailed Comparison Results**:
  - Files unique to Folder A
  - Files unique to Folder B
  - Folders unique to each location
  - Common files and folders
- **Export Functionality**: Save comparison results to a text file
- **Color-Coded Output**: Easy-to-read console output with different colors for different information types

### Find-DuplicateNames.ps1

- **GUI Folder Selection**: Choose root folder to scan using native Windows folder browser
- **Recursive Scanning**: Searches all subdirectories for files
- **Duplicate Detection**: Groups files by name (case-insensitive) and identifies duplicates
- **Detailed Results**:
  - Shows occurrence count for each duplicate filename
  - Lists all locations where each duplicate appears
  - Displays file sizes (KB/MB) for each instance
- **Summary Statistics**: Total files scanned, unique names, and duplicate counts
- **Export Functionality**: Save results to a text file with numbered suffix for multiple reports
- **Color-Coded Output**: Clear, organized console display

## Requirements

- Windows PowerShell 5.1 or later
- Windows operating system

## Usage

### Compare-Folders.ps1

Run the script:

```powershell
.\Compare-Folders.ps1
```

The script will:
1. Prompt you to select Folder A using a GUI dialog
2. Prompt you to select Folder B using a GUI dialog
3. Analyze both folders recursively
4. Display comprehensive statistics and comparison results
5. Optionally export results to a timestamped text file

### Find-DuplicateNames.ps1

Run the script:

```powershell
.\Find-DuplicateNames.ps1
```

The script will:
1. Prompt you to select the root folder to scan using a GUI dialog
2. Recursively scan all files in the folder and subdirectories
3. Identify and group files with duplicate names
4. Display detailed results with file locations and sizes
5. Optionally export results to `DuplicateNames.txt` (or numbered suffix if file exists)

## Supported Folder Types

- Local drives (C:\, D:\, etc.)
- Network paths (UNC paths)
- MTP devices (phones, tablets, cameras via USB)
- Portable storage devices

## Example Output

### Compare-Folders.ps1

```
FOLDER STATISTICS
===============================================
Folder A: C:\MyPhotos
  Total Size: 2586.95 MB (2.53 GB)
  Files: 740
  Folders: 15

Folder B: [Phone]\Internal storage\DCIM
  Total Size: 2586.95 MB (2.53 GB)
  Files: 740
  Folders: 15

SUMMARY
===============================================
Common files: 740
Common folders: 15
Files unique to A: 0
Files unique to B: 0
```

### Find-DuplicateNames.ps1

```
SCAN RESULTS
===============================================

Found 93 file name(s) with duplicates:

File Name: Camera_2022-05-25_23_24.mp4
  Occurrences: 3
  Locations:
    - 20-5-2022\Camera_2022-05-25_23_24.mp4 (1.74 MB)
    - 23-5-2022\Camera_2022-05-25_23_24.mp4 (1.74 MB)
    - 29-5-2022\Camera_2022-05-25_23_24.mp4 (1.74 MB)

SUMMARY
===============================================
Total files scanned: 3499
Unique file names: 3405
Duplicate file names: 93
Total files with duplicate names: 187
```

## Notes

### Compare-Folders.ps1
- The script performs recursive folder traversal
- File comparison is based on file names and relative paths
- For MTP devices, file size extraction uses Windows Shell namespace
- Large folders may take some time to analyze

### Find-DuplicateNames.ps1
- The script performs recursive folder traversal
- File comparison is case-insensitive (e.g., "File.txt" and "file.txt" are considered duplicates)
- Only compares filenames, not content
- Export files use `DuplicateNames.txt` as the base name, with numbered suffixes for multiple reports
- Large folder structures may take time to scan

## License

Free to use and modify.
