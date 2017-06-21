#detect system
$detectOS = [System.Environment]::OSVersion.Platform
if ($detectOS -eq "Unix") {
    $dirChar = "/" 
} else {
    $dirChar = "\"
}

# Project home directory
$projectHome = $HOME + $dirChar + "Desktop" +$dirChar + "FeesExtraction" + $dirChar
$projectTemp = $projectHome + "temp" + $dirChar
# $testingFile = $projectHome + "testing.txt"

# Create/Append the error log
$errorLog = $projectHome + "Error_Log.txt"
Write-Output "__________________________________________________________________" | Out-File -Append $errorLog
Get-Date -Format F | Out-file -Append $errorLog
Write-Output "__________________________________________________________________" | Out-File -Append $errorLog


# Test for the existence of the temp directory
$tmpExist = Test-Path $projectTemp
If (!($tmpExist)) {
    New-Item -ItemType directory -Path $projectTemp
    Write-Host Creating temp directory and beginning process
    Get-Date -Format F | Out-File -Append $errorLog
    Write-Output "Temp directory was deleted, renamed, or not yet created." | Out-File -Append $errorLog
    }


# Get list of documents
$startListPath = $projectHome + "feesExtractor.csv"
$startList = Test-Path $startListPath
if ($startList) {
        Write-Host "List of documents exists, good to go"
    } else {
        Write-Output "ERROR: DOCUMENT LIST DOES NOT EXIST! Export the document list from Excel in .csv (Comma Separated Values) format and save it to the prepared folder" | Out-File -Append $errorLog
        Write-Host "ERROR: DOCUMENT LIST DOES NOT EXIST! Export the document list from Excel in .csv (Comma Separated Values) format and save it to the prepared folder" -ForegroundColor white -BackgroundColor Red
        exit
    }
