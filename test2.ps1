# PowerShell Script to Match Sitecodes with GPO Names
# This script reads sitecode.csv and GPOname.csv files, matches them, and creates separate CSV files for matches and non-matches

# Define file paths - Rajesh.alda's Configuration
$sitecodeFile = "C:\Users\Rajesh.alda\Downloads\work\sitecode.csv"        # Input: Sitecodes file
$gpoFile = "C:\Users\Rajesh.alda\Downloads\work\GPOname.csv"              # Input: GPO names file
$outputFile = "C:\Users\Rajesh.alda\Downloads\result\sitecode_gpo_match_report.csv"  # Output: Single report only

# Check if input files exist
if (-not (Test-Path $sitecodeFile)) {
    Write-Error "File $sitecodeFile not found!"
    exit 1
}

if (-not (Test-Path $gpoFile)) {
    Write-Error "File $gpoFile not found!"
    exit 1
}

Write-Host "Reading input files..." -ForegroundColor Green

# Read the CSV files
try {
    $sitecodes = Import-Csv $sitecodeFile
    $gponames = Import-Csv $gpoFile
    
    Write-Host "Loaded $($sitecodes.Count) sitecodes and $($gponames.Count) GPO names" -ForegroundColor Yellow
} catch {
    Write-Error "Error reading CSV files: $_"
    exit 1
}

# Create array to store results
$results = @()

Write-Host "Processing all GPO names..." -ForegroundColor Green

# Create a list of sitecodes for quick lookup
$sitecodeList = $sitecodes | ForEach-Object { $_.sitecode.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

# Process each GPO name
foreach ($gpoRow in $gponames) {
    $gpoName = $gpoRow.GPOname.Trim()
    
    # Skip empty GPO names
    if ([string]::IsNullOrWhiteSpace($gpoName)) {
        continue
    }
    
    # Check if this GPO name contains any of our sitecodes ANYWHERE with precise word boundary matching
    $foundSitecode = $null
    $isMatched = $false
    
    foreach ($sitecode in $sitecodeList) {
        # Use word boundary matching to ensure exact sitecode match (not partial)
        # \b ensures we match whole words only, not parts of larger words
        if ($gpoName -match "\b$sitecode\b") {
            $foundSitecode = $sitecode
            $isMatched = $true
            break
        }
    }
    
    if ($isMatched) {
        # Found a matching sitecode
        $results += [PSCustomObject]@{
            GPOname = $gpoName
            Status = "YES"
            Sitecode = $foundSitecode
        }
        Write-Host "✓ $gpoName -> YES -> $foundSitecode" -ForegroundColor Green
    } else {
        # No matching sitecode found
        $results += [PSCustomObject]@{
            GPOname = $gpoName
            Status = "NO"
            Sitecode = ""
        }
        Write-Host "✗ $gpoName -> NO" -ForegroundColor Red
    }
}

# Export single output file
Write-Host "`nExporting results..." -ForegroundColor Green

try {
    if ($results.Count -gt 0) {
        $results | Export-Csv $outputFile -NoTypeInformation -Encoding UTF8
        Write-Host "✓ Exported $($results.Count) records to $outputFile" -ForegroundColor Cyan
    } else {
        Write-Host "⚠ No records to export" -ForegroundColor Yellow
    }
} catch {
    Write-Error "Error exporting CSV file: $_"
    exit 1
}

# Calculate statistics
$totalGPOs = $gponames.Count
$totalSitecodes = $sitecodes.Count
$yesCount = ($results | Where-Object { $_.Status -eq "YES" }).Count
$noCount = ($results | Where-Object { $_.Status -eq "NO" }).Count
$yesPercentage = if ($totalGPOs -gt 0) { [math]::Round(($yesCount / $totalGPOs) * 100, 2) } else { 0 }
$noPercentage = if ($totalGPOs -gt 0) { [math]::Round(($noCount / $totalGPOs) * 100, 2) } else { 0 }

# Display enhanced summary
Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "`nINPUT:" -ForegroundColor Yellow
Write-Host "• Total GPO names: $totalGPOs" -ForegroundColor White
Write-Host "• Total sitecodes: $totalSitecodes" -ForegroundColor White

Write-Host "`nRESULTS:" -ForegroundColor Yellow
Write-Host "• YES (GPOs with matching sitecodes): $yesCount out of $totalGPOs ($yesPercentage%)" -ForegroundColor Green
Write-Host "• NO (GPOs without matching sitecodes): $noCount out of $totalGPOs ($noPercentage%)" -ForegroundColor Red

Write-Host "`nOUTPUT:" -ForegroundColor Yellow
Write-Host "• Single file created: $($results.Count) records" -ForegroundColor Cyan
Write-Host "• File location: sitecode_gpo_match_report.csv" -ForegroundColor White

Write-Host "`n===========================================" -ForegroundColor Cyan

# Show sample results
if ($results.Count -gt 0) {
    Write-Host "`nSample results:" -ForegroundColor Cyan
    $results | Select-Object -First 5 | Format-Table -AutoSize
}

Write-Host "`nScript completed successfully!" -ForegroundColor Green