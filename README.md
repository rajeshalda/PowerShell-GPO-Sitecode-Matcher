# PowerShell GPO-Sitecode Matcher

A PowerShell script that matches Group Policy Object (GPO) names with sitecodes to identify which GPOs contain specific sitecodes in their names.

## 📋 Overview

This script reads two CSV files containing sitecodes and GPO names, performs intelligent matching using word boundary detection, and generates a comprehensive report showing which GPOs contain sitecodes and which don't.

## ✨ Features

- **Intelligent Matching**: Uses regex word boundary matching (`\b`) to ensure exact sitecode matches (not partial matches)
- **Comprehensive Reporting**: Generates a single CSV report with all results
- **Real-time Feedback**: Displays colored console output showing progress and matches
- **Error Handling**: Robust error checking for missing files and processing issues
- **Statistical Summary**: Provides detailed statistics and percentages
- **UTF-8 Encoding**: Ensures proper character encoding in output files

## 📁 File Structure

```
work/
├── sitecode.csv          # Input: List of sitecodes
├── GPOname.csv          # Input: List of GPO names
└── result/
    └── sitecode_gpo_match_report.csv  # Output: Complete matching report
```

## 📊 Input Files Format

### sitecode.csv
```csv
sitecode
ABC123
DEF456
GHI789
```

### GPOname.csv
```csv
GPOname
Policy_ABC123_Security
Network_Settings_DEF456
Standard_Configuration
```

## 📈 Output Format

### sitecode_gpo_match_report.csv
```csv
GPOname,Status,Sitecode
Policy_ABC123_Security,YES,ABC123
Network_Settings_DEF456,YES,DEF456
Standard_Configuration,NO,
```

## 🚀 Usage

### Prerequisites
- Windows PowerShell 5.1 or PowerShell Core 6+
- Read access to input CSV files
- Write access to output directory

### Running the Script

1. **Update file paths** in the script to match your environment:
   ```powershell
   $sitecodeFile = "C:\Your\Path\sitecode.csv"
   $gpoFile = "C:\Your\Path\GPOname.csv"
   $outputFile = "C:\Your\Path\result\sitecode_gpo_match_report.csv"
   ```

2. **Execute the script**:
   ```powershell
   .\test2.ps1
   ```

3. **Check the results** in the generated CSV file and console output.

## 🔧 Configuration

### File Paths
Update these variables at the top of the script:
- `$sitecodeFile`: Path to your sitecode CSV file
- `$gpoFile`: Path to your GPO names CSV file
- `$outputFile`: Path for the output report

### Matching Logic
The script uses word boundary matching (`\b$sitecode\b`) which ensures:
- ✅ `ABC123_Policy` matches sitecode `ABC123`
- ✅ `Policy_ABC123_Settings` matches sitecode `ABC123`
- ❌ `XABC123Y` does NOT match sitecode `ABC123`

## 📊 Output Features

### Console Output
- **Green checkmarks** (✓) for successful matches
- **Red X marks** (✗) for non-matches
- **Colored progress indicators** for different stages
- **Comprehensive summary statistics**

### CSV Report Columns
- **GPOname**: Original GPO name from input
- **Status**: "YES" if sitecode found, "NO" if not found
- **Sitecode**: The matching sitecode (empty for non-matches)

### Summary Statistics
- Total GPO names processed
- Total sitecodes available
- Number and percentage of matches
- Number and percentage of non-matches
- Sample results preview

## ⚠️ Error Handling

The script includes robust error handling for:
- **Missing input files**: Exits with error code 1
- **CSV reading errors**: Displays detailed error messages
- **Export failures**: Handles file writing issues
- **Empty data**: Gracefully handles empty input files

## 🔍 Technical Details

### Matching Algorithm
1. Loads all sitecodes into memory for efficient lookup
2. Iterates through each GPO name
3. Uses regex word boundary matching for precise detection
4. Stops at first matching sitecode (if multiple exist)
5. Records results with status and matched sitecode

### Performance Considerations
- **Memory Usage**: Loads all data into memory for fast processing
- **Time Complexity**: O(n×m) where n=GPO count, m=sitecode count
- **Suitable for**: Typical enterprise environments (thousands of GPOs/sitecodes)

## 🛠️ Customization Options

### Modify Matching Logic
To change matching behavior, update this section:
```powershell
if ($gpoName -match "\b$sitecode\b") {
    # Current: Word boundary matching
    # Alternative: Case-insensitive: -imatch
    # Alternative: Exact match: -eq
}
```

### Add Additional Columns
Extend the result object:
```powershell
$results += [PSCustomObject]@{
    GPOname = $gpoName
    Status = $isMatched ? "YES" : "NO"
    Sitecode = $foundSitecode
    # Add custom fields here
    MatchCount = $matchCount
    Timestamp = Get-Date
}
```

## 📝 License

[Specify your license here]

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## 📞 Support

For issues, questions, or contributions, please [create an issue](link-to-issues) or contact the maintainer.

---

*Script optimized for Windows environments and Group Policy management workflows.*
