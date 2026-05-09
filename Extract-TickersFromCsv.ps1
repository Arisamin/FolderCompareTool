param(
    [string]$CsvPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

function Select-CsvFile {
    param([string]$Title = 'Select source file (CSV-formatted content)')

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = 'All files (*.*)|*.*'
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    $dialog.CheckPathExists = $true

    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return $dialog.FileName
}

function Get-CsvHeaders {
    param([Parameter(Mandatory = $true)][string]$Path)

    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path)
    $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(',')
    $parser.HasFieldsEnclosedInQuotes = $true

    try {
        return $parser.ReadFields()
    }
    finally {
        $parser.Close()
    }
}

Write-Host 'Extract Ticker Column Tool' -ForegroundColor Cyan
Write-Host '==========================' -ForegroundColor Cyan
Write-Host ''

$sourceCsvPath = $CsvPath
if ([string]::IsNullOrWhiteSpace($sourceCsvPath)) {
    $sourceCsvPath = Select-CsvFile
    if (-not $sourceCsvPath) {
        Write-Host 'No file selected. Exiting.' -ForegroundColor Yellow
        exit 0
    }
}

if (-not (Test-Path -Path $sourceCsvPath -PathType Leaf)) {
    Write-Host "File not found: $sourceCsvPath" -ForegroundColor Red
    exit 1
}

try {
    $headers = Get-CsvHeaders -Path $sourceCsvPath
}
catch {
    Write-Host "Failed to read CSV headers: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

if (-not $headers -or $headers.Count -eq 0) {
    Write-Host 'Input file appears empty or has no headers.' -ForegroundColor Red
    exit 1
}

$tickerHeaderName = $headers | Where-Object { $_ -ieq 'Ticker' } | Select-Object -First 1
if (-not $tickerHeaderName) {
    Write-Host 'Column "Ticker" was not found in the selected CSV.' -ForegroundColor Red
    exit 1
}

try {
    $rows = @(Import-Csv -Path $sourceCsvPath -ErrorAction Stop)
}
catch {
    Write-Host "Failed to read CSV rows: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$sourceDirectory = Split-Path -Path $sourceCsvPath -Parent
$sourceBaseName = [System.IO.Path]::GetFileNameWithoutExtension($sourceCsvPath)
$sourceExtension = [System.IO.Path]::GetExtension($sourceCsvPath)
$outputFileName = "${sourceBaseName}_tickers_only$sourceExtension"
$outputPath = Join-Path $sourceDirectory $outputFileName

if ($rows.Count -eq 0) {
    # Keep CSV shape even if there are no data rows.
    Set-Content -Path $outputPath -Value 'Ticker' -Encoding UTF8
    Write-Host "Source CSV contains no data rows. Created header-only file: $outputPath" -ForegroundColor Yellow
    exit 0
}

$tickerLines = New-Object System.Collections.Generic.List[string]
$tickerLines.Add('Ticker') | Out-Null

foreach ($row in $rows) {
    $rawTicker = $row.PSObject.Properties[$tickerHeaderName].Value
    $tickerText = [string]$rawTicker
    # Keep output quote-free as requested.
    $tickerText = $tickerText -replace '"', ''
    $tickerLines.Add($tickerText) | Out-Null
}

try {
    Set-Content -Path $outputPath -Value $tickerLines -Encoding UTF8
}
catch {
    Write-Host "Failed to write output CSV: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor Green
Write-Host "Input file: $sourceCsvPath" -ForegroundColor White
Write-Host "Output file: $outputPath" -ForegroundColor White
Write-Host "Rows exported: $($rows.Count)" -ForegroundColor White