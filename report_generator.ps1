<#
.SYNOPSIS
    Retrieves metadata from a SharePoint document library and exports it to an Excel file in batches.

.DESCRIPTION
    This script connects to a SharePoint site using PnP PowerShell, retrieves metadata for files in a specified document library,
    and exports the data to an Excel file. The script handles large libraries by paginating through the items and logging all
    operations and errors to a text log file. The Excel file includes details such as file name, extension, path, size, and
    user information.

.PARAMETER siteUrl
    The URL of the SharePoint site containing the document library.

.PARAMETER libraryName
    The name of the document library from which to fetch data.

.PARAMETER logFileName
    The name of the log file to record script actions and errors. The file should end with '.txt'.

.PARAMETER excelFileName
    The name of the Excel file where the data will be exported. The file should end with '.xlsx'.

.EXAMPLE
    .\report_generator.ps1
    This will prompt the user for SharePoint site URL, document library name, log file name, and Excel file name. The script
    will then connect to SharePoint, fetch the document library data in batches, and save it to the specified Excel file.

.NOTES
    File extensions for log and Excel files are validated to ensure correct format.
    The script handles pagination to avoid exceeding SharePoint's list view threshold limit.
#>

# Import necessary modules
Import-Module PnP.PowerShell -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

# Function to log messages with timestamps
function Log-Message {
    param (
        [string]$Message,
        [string]$LogFile
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $Message"
    Write-Host $logEntry
    Add-Content -Path $LogFile -Value $logEntry
}

try {
    # Prompt user for input
    $siteUrl = Read-Host -Prompt "Enter SharePoint site URL"
    $libraryName = Read-Host -Prompt "Enter document library name"
    $logFileName = Read-Host -Prompt "Enter log file name (.txt)"
    $excelFileName = Read-Host -Prompt "Enter Excel file name (.xlsx)"

    # Ensure file extensions are correct
    if (-not $logFileName.EndsWith(".txt")) {
        $logFileName += ".txt"
    }
    if (-not $excelFileName.EndsWith(".xlsx")) {
        $excelFileName += ".xlsx"
    }

    # Initialize log file
    New-Item -Path $logFileName -ItemType File -Force | Out-Null
    Log-Message -Message "Starting script" -LogFile $logFileName

    # Record script start time
    $scriptStartTime = Get-Date
    Log-Message -Message "Script start time: $scriptStartTime" -LogFile $logFileName

    # Authenticate to SharePoint
    $authStartTime = Get-Date
    Log-Message -Message "Authentication started at: $authStartTime" -LogFile $logFileName
    Connect-PnPOnline -Url $siteUrl -Interactive
    $authEndTime = Get-Date
    Log-Message -Message "Authenticated to SharePoint site $siteUrl at: $authEndTime" -LogFile $logFileName

    # Initialize an Excel file and add header
    $data = @()
    $data | Export-Excel -Path $excelFileName -WorksheetName "Report" -AutoSize -TableName "DocumentLibraryData" -ClearSheet

    # Initialize variables for pagination
    $pageSize = 5000
    $itemCount = 0

    do {
        try {
            Log-Message -Message "Fetching items batch $($itemCount / $pageSize + 1)" -LogFile $logFileName

            # Retrieve items from SharePoint
            $items = Get-PnPListItem -List $libraryName -PageSize $pageSize -Fields "FileLeafRef", "FileRef", "File_x0020_Size", "Author", "Editor", "Created", "Modified"

            if ($items.Count -eq 0) {
                Log-Message -Message "No more items found." -LogFile $logFileName
                break
            }

            $batchData = @()
            foreach ($item in $items) {
                try {
                    $processStartTime = Get-Date
                    $fileName = $item["FileLeafRef"]
                    $fileSizeMB = [math]::Round($item["File_x0020_Size"] / 1MB, 2)
                    $fileExtension = [System.IO.Path]::GetExtension($fileName)
                    $filePath = $item["FileRef"]
                    $createdBy = $item["Author"].Email
                    $createdDate = $item["Created"]
                    $modifiedBy = $item["Editor"].Email
                    $modifiedDate = $item["Modified"]

                    $dataRow = [PSCustomObject]@{
                        FileName        = $fileName
                        FileSizeMB      = $fileSizeMB
                        FileExtension   = $fileExtension
                        FilePath        = $filePath
                        CreatedByEmail  = $createdBy
                        CreatedDate     = $createdDate
                        ModifiedByEmail = $modifiedBy
                        ModifiedDate    = $modifiedDate
                    }

                    $batchData += $dataRow

                    $processEndTime = Get-Date
                    Log-Message -Message "Processed file: $fileName (Start: $processStartTime, End: $processEndTime)" -LogFile $logFileName

                } catch {
                    $rowErrorMessage = $_.Exception.Message
                    Log-Message -Message "Error processing file: $fileName - $rowErrorMessage" -LogFile $logFileName
                    Write-Error "Error processing file: $fileName - $rowErrorMessage"
                }
            }

            # Write batch data to Excel
            $batchData | Export-Excel -Path $excelFileName -WorksheetName "Report" -Append -AutoSize

            $itemCount += $items.Count

        } catch {
            $queryErrorMessage = $_.Exception.Message
            Log-Message -Message "Error retrieving items: $queryErrorMessage" -LogFile $logFileName
            Write-Error "Error retrieving items: $queryErrorMessage"
            break
        }
    } while ($items.Count -eq $pageSize)  # Continue if the number of items fetched equals the page size

    # Record script end time
    $scriptEndTime = Get-Date
    Log-Message -Message "Script completed successfully at: $scriptEndTime" -LogFile $logFileName

} catch {
    $errorMessage = $_.Exception.Message
    Log-Message -Message "Error: $errorMessage" -LogFile $logFileName
    Write-Error $errorMessage
} finally {
    # Clean up and disconnect from SharePoint
    try {
        Disconnect-PnPOnline
        Log-Message -Message "Disconnected from SharePoint" -LogFile $logFileName
    } catch {
        Log-Message -Message "Error during disconnection: $_.Exception.Message" -LogFile $logFileName
    }

    $finalEndTime = Get-Date
    Log-Message -Message "Script terminated at: $finalEndTime" -LogFile $logFileName
}
