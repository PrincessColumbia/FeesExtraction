# Project home directory
$projectHome = $HOME + "\Desktop\FeesExtraction\"
$projectTemp = $projectHome + "temp\"
$testingFile = $projectHome + "testing.txt"


# Create/Append the error log
$errorLog = $projectHome + "Error_Log.txt"
Write-Output "__________________________________________________________________" | Out-File -Append $errorLog
Get-Date -Format F | Out-file -Append $errorLog
Write-Output "__________________________________________________________________" | Out-File -Append $errorLog


# Test for the existence of the temp directory
$tmpExist = Test-Path $projectTemp
If ($tmpExist)
    {Write-Host Beginning process}
    else
    {New-Item -ItemType directory -Path $projectTemp
    Write-Host Creating temp directory and beginning process
    Get-Date -Format F | Out-File -Append $errorLog
    Write-Output "Temp directory was deleted, renamed, or not yet created." | Out-File -Append $errorLog
    }


# Get list of documents
$startListPath = $projectHome + "Put the list of documents to process here\ppt_prepared.csv"
$startList = Test-Path $startListPath
if ($startList) {
        Write-Host Processing list of documents
    } else {
        Write-Output "ERROR: DOCUMENT LIST DOES NOT EXIST! Export the document list from Excel in .csv (Comma Separated Values) format and save it to the prepared folder" | Out-File -Append $errorLog
        Write-Host "ERROR: DOCUMENT LIST DOES NOT EXIST! Export the document list from Excel in .csv (Comma Separated Values) format and save it to the prepared folder" -ForegroundColor white -BackgroundColor Red
        exit
    }

$importDataSet = Import-Csv $startListPath

$importDataSet | Add-Member -MemberType NoteProperty -Name CityOrEntity -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name State -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Assigned -Value $env:USERNAME
$importDataSet | Add-Member -MemberType NoteProperty -Name "Form Completed" -Value Schrodinger
$importDataSet | Add-Member -MemberType NoteProperty -Name Notes -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Downloaded -Value $null


$importDataSet | Foreach {
    
    
    $rateId = $_.LawsonDiv + "-" + $_.DivisionNo + "-" + $_.PolygonID
    $_."Rate ID" = $rateId

    $AreaWeb = $_.Area -replace ' ','%20'
    $AreaWebSub = $AreaWeb + "/"
    $areaSpaceUnderscore = $_.Area -replace ' ','_'
    $areaDashUnderscore = $_.Area -replace '-','_'
    $areaOnlyUnderscore = $areaSpaceUnderscore -replace '-','_'
    $AreaOnly = $_.Area.Substring(4)
    $AreaOnlyUnderscore = $areaSpaceUnderscore.Substring(4)

  
    $compassWebLocation = "http://compass.repsrv.com/DivisionalDocuments/"
    $compassUrl = $compassWebLocation + $AreaWebSub + $_.Name
    $_.Link = $compassUrl

    $fileSaveLocation = $projectTemp + $_.Name
    [io.file]::WriteAllText($fileSaveLocation,(Invoke-WebRequest -URI $compassUrl -UseDefaultCredentials -UseBasicParsing).content)

<#    catch { 
        if($_.Exception.Response.StatusCode -eq 'NotFound') {
            $urlFail = "Failed"
        } else {
            $urlFail = "Success"
        }
    }
    #>
<#
    if ($urlFail -eq "Failed") {
#        $_.CityOrEntity = $null
#        $_.State = $null
#        $_."Form Completed" = "No"
        $_.Notes = "Unable to find"
        $_."In Compass" = "No"
    }
    #>


    <#
    
    $workFilePath = $projectTemp + $_.Name
    $workFile = Test-Path $workFilePath
    if (!($workFile)) {
        Write-Host $_.Name ERROR: DOCUMENT DOES NOT EXIST! See the error log.
        Write-Output "$($_.Name) was not found in the temp file location. This file may not exist on the server. Double-check the name and perform a Compass search as necessary" | Out-File -Append $errorLog
        } else {
        Write-Output "File found, test passes" | Out-File -Append $testingFile
    }

    # Reminder - Keep the following line around for reference
    # Get-Content .\DIV_092_270_MUNI_Cement_City_MI.html | Select-String -Pattern "Solid\s+Waste\s+Rate"

    #Parse data from files

    # State
    #$html = New-Object -ComObject "HTMLFile";
    #$source = Get-Content -Path $workFilePath -Raw;
    #$html.IHTMLDocument2_write($source)
    [string]$parsedState = Select-String $workFilePath -Pattern "mso:State" -ErrorAction SilentlyContinue
    $parsedState = $parsedState -split ">" | Where { $_ -notmatch "DIV" }
    $parsedState = $parsedState -split "<" | Where { $_ -notmatch "State" }
    $state = $parsedState.Trim()

    $_.State = $state

    if ([string]::IsNullOrEmpty($_.State)) {
        $_.State = "Not in mso tags"
        $stateError = "Missing/Malformed State designator, check the source documents"
        if ([string]::IsNullOrEmpty($_.Notes)) {
            $_.Notes = $stateError
        } else {
            $_.Notes = $_.Notes + ", " + $stateError
    }

    # City
    if ($_.State -eq "Not in mso tags") {
        $_.CityOrEntity = "Not in mso tags"
        $cityError = "Missing/Malformed City or Entity name, check the source documents"
        if ([string]::IsNullOrEmpty($_.Notes)) {
            $_.Notes = $cityError
        } else {
            $_.Notes = $_.Notes + ", " + $cityError
        }

    } else {
        [string]$parsedCity = Select-String $workFilePath -Pattern "mso:CityOrEntity"
        $parsedCity = $parsedCity -split ">" | Where { $_ -notmatch "DIV" }
        $parsedCity = $parsedCity -split "<" | Where { $_ -notmatch "CityOrEntity" }
        $city = $parsedCity.Trim()

        $_.CityOrEntity = $city

    }
    if ([string]::IsNullOrEmpty($_.CityOrEntity)) {
        $_.CityOrEntity = "Not in mso tags"
                if ([string]::IsNullOrEmpty($_.Notes)) {
            $_.Notes = $cityError
        } else {
            $_.Notes = $_.Notes + ", " + $cityError
        }
    }
    
}
#>
}

# Test the export, this is NOT the final product, that needs to be divided out by DIV
$exportTest = $projectHome + "export-test.csv"
$importDataSet | Select-Object Name,LawsonDiv,DivisionNo,PolygonID,"Rate ID",CityOrEntity,State,Assigned,"Form Completed",Notes | Export-Csv -Append $exportTest -NoTypeInformation #$areaTrackerFilename

<#
# Clean up
Remove-Item $projectTemp -Recurse
New-Item -ItemType directory -Path $projectTemp
Remove-Variable * -ErrorAction SilentlyContinue

#>