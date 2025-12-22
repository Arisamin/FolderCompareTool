# Folder Compare Tool

A PowerShell script for comparing the contents of two folders with GUI folder selection. Supports both regular Windows file system paths and MTP devices (phones, cameras connected via USB).

## Features

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

## Requirements

- Windows PowerShell 5.1 or later
- Windows operating system

## Usage

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

## Supported Folder Types

- Local drives (C:\, D:\, etc.)
- Network paths (UNC paths)
- MTP devices (phones, tablets, cameras via USB)
- Portable storage devices

## Example Output

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

## Notes

- The script performs recursive folder traversal
- File comparison is based on file names and relative paths
- For MTP devices, file size extraction uses Windows Shell namespace
- Large folders may take some time to analyze

## License

Free to use and modify.
