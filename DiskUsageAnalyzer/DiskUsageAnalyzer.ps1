# Get the script's directory
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Prompt for folder path
$targetPath = Read-Host "Enter folder path to analyze (Leave blank for entire C:\ drive)"
if ([string]::IsNullOrWhiteSpace($targetPath)) {
    $targetPath = "C:\"
}

# Validate path
if (-not (Test-Path $targetPath)) {
    Write-Host "ERROR: The path '$targetPath' does not exist." -ForegroundColor Red
    exit
}

Write-Host "Scanning files in $targetPath... Please wait." -ForegroundColor Cyan

# Start timer
$sw = [System.Diagnostics.Stopwatch]::StartNew()

# Get files sorted by size
$files = Get-ChildItem -Path $targetPath -Recurse -File -ErrorAction SilentlyContinue |
    Sort-Object Length -Descending |
    Select-Object @{Name='SizeMB'; Expression={"{0:N2}" -f ($_.Length / 1MB)}},
                  @{Name='LastModified'; Expression={$_.LastWriteTime}},
                  @{Name='Path'; Expression={$_.FullName}}

# Define output path
$csvPath = Join-Path $scriptDirectory "FileAnalysis.csv"

# Export to CSV
Write-Host "Saving results to $csvPath ..." -ForegroundColor Yellow
$files | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

# End timer
$sw.Stop()
Write-Host "Done! Analyzed $($files.Count) files in $($sw.Elapsed.TotalSeconds) seconds." -ForegroundColor Green
Write-Host "Output saved at: $csvPath"
