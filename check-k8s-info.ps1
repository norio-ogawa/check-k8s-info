# check-k8s-info.ps1

# Extract case number

function Get-CaseNumber {
    # Define the log file path
    $logFilePath = ".get-k8s-info\get-k8s-info.log"

    # Return default if file does not exist
    if (-not (Test-Path $logFilePath)) {
        return "CS0000000"
    }

    # Read first 10 lines from the file
    $lines = Get-Content $logFilePath -TotalCount 10

    foreach ($line in $lines) {
        # Check if line starts with "CASENUMBER:"
        if ($line -match '^CASENUMBER:\s*(\S+)') {
            return $matches[1]
        }
    }

    # Return default if no matching line found
    return "CS0000000"
}

# Extract datetime

function Get-TimestampFromLog {
    # Define the log file path
    $logFilePath = ".get-k8s-info\get-k8s-info.log"

    # Month abbreviation map
    $monthMap = @{
        "Jan" = "01"; "Feb" = "02"; "Mar" = "03"; "Apr" = "04"
        "May" = "05"; "Jun" = "06"; "Jul" = "07"; "Aug" = "08"
        "Sep" = "09"; "Oct" = "10"; "Nov" = "11"; "Dec" = "12"
    }

    # If file does not exist, return current timestamp
    if (-not (Test-Path $logFilePath)) {
        return (Get-Date).ToString("yyyyMMdd_HHmmss")
    }

    # Read the second line
    $lines = Get-Content $logFilePath -TotalCount 2
    if ($lines.Count -lt 2) {
        return (Get-Date).ToString("yyyyMMdd_HHmmss")
    }

    $timestampLine = $lines[1]

    # Match timestamp format like "Tue Sep 30 04:00:01 JST 2025"
    if ($timestampLine -match '^\w{3} (\w{3}) (\d{1,2}) (\d{2}):(\d{2}):(\d{2}) \w{3,4} (\d{4})$') {
        $monthStr = $matches[1]
        $day = $matches[2].PadLeft(2, '0')
        $hour = $matches[3]
        $minute = $matches[4]
        $second = $matches[5]
        $year = $matches[6]

        if ($monthMap.ContainsKey($monthStr)) {
            $month = $monthMap[$monthStr]
            return "$year$month$day" + "_$hour$minute$second"
        }
    }

    # If timestamp not found or invalid, return current timestamp
    return (Get-Date).ToString("yyyyMMdd_HHmmss")
}

# Find the date after the specified line in the log file

function Get-LogTimestamp {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LogFile,       # Path of the log file
        [Parameter(Mandatory = $true)]
        [int]$StartLine         # Start line for date lookup (0-based)
    )

    # Read the log file as an array
    $lines = Get-Content -Path $LogFile

    # Check range�E�Eax 50 lines�E�E
    $endLine = [math]::Min($StartLine + 50, $lines.Count - 1)

    # Date patterm YYYY-MM-DD or YYYY/MM/DD�E�E
    $datePattern = '20\d{2}[-/]\d{2}[-/]\d{2}'

    for ($i = $StartLine; $i -le $endLine; $i++) {
        if ($lines[$i] -match $datePattern) {
            return $matches[0]  # Return the first matching date
        }
    }
    return "N/A"
}

# Merge temporary keyword files into a single TSV file.

function Merge-KeywordFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$OutputFileName
    )

    # Get all keyword result files and sort by name
    $files = Get-ChildItem -Path "." -Filter "_keyword*.txt" | Sort-Object Name

    # Exit if no files are found
    if ($files.Count -eq 0) {
        Write-Warning "No keyword files found."
        return
    }

    # Create output file and write TSV header
    "Pattern`tDate`tFile`tLine`tMessage" | Out-File -FilePath $OutputFileName -Encoding UTF8

    # Merge all temporary files
    foreach ($file in $files) {
        Get-Content $file.FullName | Add-Content -Path $OutputFileName
    }

    # Remove temporary files
    foreach ($file in $files) {
        Remove-Item -Path $file.FullName -Force
    }
}

# Import TSV results into the ErrorResults worksheet of an Excel file.
function Import-ErrorResultsToExcel {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ExcelFileName,

        [Parameter(Mandatory = $true)]
        [string]$TsvFileName
    )

    # Open Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        # Open workbook
        $workbook = $excel.Workbooks.Open((Resolve-Path $ExcelFileName).Path)

        # Delete existing worksheet if present
        try {
            $worksheet = $workbook.Worksheets.Item("ErrorResults")
            $worksheet.Delete()
        }
        catch {
            # Worksheet does not exist
        }

        # Create new worksheet
        $worksheet = $workbook.Worksheets.Add()
        $worksheet.Name = "ErrorResults"

        # Read TSV data
        $rows = Get-Content $TsvFileName

        $rowIndex = 1

        foreach ($row in $rows) {

            $columns = $row -split "`t"

            for ($colIndex = 0; $colIndex -lt $columns.Count; $colIndex++) {
                $worksheet.Cells.Item($rowIndex, $colIndex + 1) = $columns[$colIndex]
            }

            $rowIndex++
        }

        # Auto-fit columns
        $worksheet.Columns.AutoFit() | Out-Null

        
        # Freeze the first row
        $worksheet.Activate() | Out-Null
        $excel.ActiveWindow.SplitRow = 1
        $excel.ActiveWindow.FreezePanes = $true

        # Set background color(2)
        $worksheet.Cells.Interior.ColorIndex = 2

        # Save workbook
        $workbook.Save()
    }
    finally {
        if ($workbook) {
            $workbook.Close($true)
        }

        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# Get error patterns from Excel or use built-in defaults.
function Get-ErrorPatterns {
    param (
        [string]$ExcelFileName
    )

    # Default patterns
    $defaultPatterns = @(
        "SIGSEGV",
        "OOMKilled",
        "OutOfMemory",
        "out of memory",
        "Java heap space",
        "panic:",
        "JobExecutionException",
        "unhandled Exception",
        "SSL error",
        "SAS/TK is aborting",
        "OAuth token is expired",
        "OOMKilling",
        "Operation timed out",
        "INTERNAL_SERVER_ERROR",
        "I/O error on",
        "ServletOutputStream failed to write:",
        "Gateway Time-out",
        "No space left",
        "endpoints have no available addresses",
        "FailedScheduling",
        "Traceback"
    )

    # Excel file does not exist
    if (-not (Test-Path $ExcelFileName)) {
        Write-Host "Using default error patterns (file not found)."
        return $defaultPatterns
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open((Resolve-Path $ExcelFileName).Path)

        try {
            $worksheet = $workbook.Worksheets("ErrorPatterns")
            Write-Host "Error Patterns: Loaded error patterns from ErrorPatterns worksheet."
        }
        catch {
            Write-Host "Using default error patterns (ErrorPatterns worksheet not found)."
            return $defaultPatterns
        }

        $patterns = @()

        # Start from row 1 (no header)
        $row = 1

        while ($true) {

            $value = $worksheet.Cells.Item($row, 1).Text

            if ([string]::isNullOrWhiteSpace($value)) {
                break
            }

            $patterns += $value.Trim()
            $row++
        }

        if ($patterns.Count -eq 0) {
            Write-Host "No patterns found. Using default patterns."
            return $defaultPatterns
        }

        return $patterns
    }
    finally {

        if ($workbook) {
            $workbook.Close($false)
        }

        if ($excel) {
            $excel.Quit()
        }

        if ($worksheet) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        }

        if ($workbook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }

        if ($excel) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Look for error keywords in files

function Search-K8sErrors {
    param (
        [string]$ResultFilePath,
        [string]$SummaryFilePath,
        [string]$TemplateExcelFileName
    )
    # Keyword count
    [int]$count = 0
    [int]$fcount = 0

    # Use current directory as base
    $scriptDir = Get-Location

    # Error keywords (can be extended by user)
    $errorPatterns = Get-ErrorPatterns -ExcelFileName $TemplateExcelFileName

    # Initialize output files
    "" | Out-File -FilePath $ResultFilePath -Encoding UTF8
    "Date,Keyword,Count" | Out-File -FilePath $SummaryFilePath -Encoding UTF8

    # Get target files (.log only, including subfolders)
    $files = Get-ChildItem -Path $scriptDir -Include *.log -File -Recurse

    # Hashtable to store counts per timestamp for each keyword
    $summaryTable = @{}
    foreach ($keyword in $errorPatterns) {
        $summaryTable[$keyword] = @{}
    }

    # Search for each keyword
    foreach ($file in $files) {
        $fcount++
        $count=0
        foreach ($keyword in $errorPatterns) {
            $count++
            [string]$tempFileName = "_keyword{0:D3}.txt" -f $count
            $matches = Select-String -Path $file.FullName -Pattern $keyword -CaseSensitive:$false
            foreach ($match in $matches) {
                $line = $match.Line.Trim()
                $lineNumber = $match.LineNumber
                $fileName = Split-Path $match.Path -Leaf  # Extract file name only

                # Extract timestamp
                if ($line -match '(\d{4})[/-](\d{2})[/-](\d{2})') {
                    # Convert yyyy/MM/dd to yyyy-MM-dd
                    $timestamp = "$($matches[1])-$($matches[2])-$($matches[3])"
                } elseif ($line -match '"[@]*time[sS]tamp":"(\d{4}-\d{2}-\d{2})T') {
                    # Extract ISO format yyyy-MM-dd
                    $timestamp = $matches[1]
                } else {
                    $timestamp = "N/A"
                    $timestamp = Get-LogTimestamp $file.FullName $match.LineNumber
                }

                # Write to result file
                $line = $line -replace "`r"," " -replace "`n"," "
                [string]$outLine = "$keyword`t$timestamp`t$fileName`t$lineNumber`t$line"
                Add-Content -Path $tempFileName -Value $outLine

                # Aggregate counts
                if (-not $summaryTable[$keyword].ContainsKey($timestamp)) {
                    $summaryTable[$keyword][$timestamp] = 0
                }
                $summaryTable[$keyword][$timestamp]++
            }
        }
    }
    Write-Host " Checked Files: $fcount"
    
    # Merge temporary keyword files into a single TSV file.
    Merge-KeywordFiles -OutputFileName $ResultFilePath

    # Prepare summary output
    $summaryOutput = @()
    foreach ($keyword in $errorPatterns) {
        if ($summaryTable[$keyword].Count -eq 0) {
            $summaryOutput += [PSCustomObject]@{
                Timestamp = "N/A"
                Keyword = $keyword
                Count = 0
            }
        } else {
            foreach ($timestamp in $summaryTable[$keyword].Keys) {
                $summaryOutput += [PSCustomObject]@{
                    Timestamp = $timestamp
                    Keyword = $keyword
                    Count = $summaryTable[$keyword][$timestamp]
                }
            }
        }
    }

    # Sort and write summary file
    $summaryOutput |
        Sort-Object Timestamp, Keyword |
        ForEach-Object {
            Add-Content -Path $SummaryFilePath -Value "$($_.Timestamp),$($_.Keyword),$($_.Count)"
        }
}

# Extract case number and datetime

function Get-K8sInfoIdentifier {
    $case = Get-CaseNumber
    $timestamp = Get-TimestampFromLog
    $name = $case + "_" + $timestamp
    return $name
}

# Convert csv file to excel file

function Convert-CsvToExcel {
    param (
        [string]$CsvFileName,
        [string]$ExcelFileName,
        [string]$SheetName,
        [string]$TemplateExcelFileName
    )

    # Get the current directory
    $currentDir = Get-Location
    $csvPath      = Join-Path $currentDir $CsvFileName
    $excelPath    = Join-Path $currentDir $ExcelFileName
    $templatePath = Join-Path $currentDir $TemplateExcelFileName

    # Start Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Open the excel template file
    $workbook = $excel.Workbooks.Open($templatePath)

    # Set first sheet
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = $SheetName

    # Read the csv file
    $csvData = Import-Csv -Path $csvPath
    $row = 1

    # Write headers
    $headers = $csvData[0].PSObject.Properties.Name
    for ($col = 0; $col -lt $headers.Count; $col++) {
        $worksheet.Cells.Item($row, $col + 1).Value2 = $headers[$col]
    }

    # Write data rows
    foreach ($line in $csvData) {
        $row++
        for ($col = 0; $col -lt $headers.Count; $col++) {
            $worksheet.Cells.Item($row, $col + 1).Value2 = $line.$($headers[$col])
        }
    }

    # Format the sheet
    $lastRow = $row
    $lastCol = $headers.Count
    $range = $worksheet.Range("A1", $worksheet.Cells.Item($lastRow, $lastCol))

    $range.AutoFormat(1) | Out-Null
    $worksheet.Range("A1", $worksheet.Cells.Item(1, $lastCol)).AutoFilter() | Out-Null
    $worksheet.Cells.Interior.ColorIndex = 2

    # Save as new file name
    $workbook.SaveAs($excelPath)
    $workbook.Close($false)
    $excel.Quit()

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}


# Create a PowerShell 7 shortcut for this script if it does not already exist.

function New-CheckK8sInfoShortcut {

    $scriptPath   = $PSCommandPath
    $scriptFolder = Split-Path $scriptPath -Parent
    $scriptName   = [System.IO.Path]::GetFileNameWithoutExtension($scriptPath)

    $shortcutPath = Join-Path $scriptFolder "$scriptName.lnk"

    # Skip if shortcut already exists
    if (Test-Path $shortcutPath) {
        return
    }

    $pwshPath = (Get-Command pwsh).Source

    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($shortcutPath)

    $shortcut.TargetPath = $pwshPath
    $shortcut.Arguments = "-NoExit -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    $shortcut.WorkingDirectory = $scriptFolder
    $shortcut.IconLocation = "$pwshPath,0"

    $shortcut.Save()

    Write-Host ""
    Write-Host "A PowerShell 7 shortcut has been created."
    Write-Host "Shortcut: $shortcutPath"
    Write-Host ""
    Read-Host "Press Enter to continue"
}

# Argument check & base directory definition

$excelTemplateName = "_template.xlsx"
$templateSource = Join-Path $PSScriptRoot $excelTemplateName

if ($args.Count -eq 0) {
    Write-Host "ERROR: Please specify at least one target folder to check." -ForegroundColor Red
    New-CheckK8sInfoShortcut
    exit 1
}

if (-not (Test-Path $templateSource)) {
    Write-Host "ERROR: Required template file $templateSource was not found." -ForegroundColor Red
    Write-Host "Create an empty Excel file named $excelTemplateName and save it in the script folder." -ForegroundColor Red
    exit 1
}

foreach ($targetFolder in $args) {
    if (-not (Test-Path $targetFolder -PathType Container)) {
        Write-Host "ERROR: $targetFolder is not a folder." -ForegroundColor Red
        exit 1
    }
}

Write-Host ("    Start Time: " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))

foreach ($BaseDir in $args) {

    $excelTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $excelWorkName = "_template_${excelTimestamp}.xlsx"
    $templateDest = Join-Path $BaseDir $excelWorkname



    Write-Host " Target Folder: $BaseDir"

    # Copy Excel Template File


    Copy-Item -Path $templateSource -Destination $templateDest -Force

    # Main processing

    Push-Location $BaseDir
    try {
        $resultsPath = "error_results.txt"
        $summaryPath = "error_summary.csv"
        
        $excelPath= "error_summary.xlsx"

        Search-K8sErrors -ResultFilePath $resultsPath -SummaryFilePath $summaryPath -TemplateExcelFileName $excelWorkName

        ## Assign the result to the variable
        $sheetName = Get-K8sInfoIdentifier

        ## Call the function with explicit arguments
        Convert-CsvToExcel -CsvFileName $summaryPath -ExcelFileName $excelPath -SheetName $sheetName -TemplateExcelFileName $excelWorkName

        Import-ErrorResultsToExcel -ExcelFileName $excelPath -TsvFileName $resultsPath
        Write-Host "   Result File: $excelPath"

        # Delete temporary TSV file if it exists
        Remove-Item -Path $summaryPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $resultsPath -Force -ErrorAction SilentlyContinue
    }
    finally {
        Pop-Location
    }

    # Delete Excel Template File

    if (Test-Path $templateDest) {
        Remove-Item -Path $templateDest -Force
    }

    ## Completion message
}

Write-Host ("      End Time: " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"))
Write-Host "Done."