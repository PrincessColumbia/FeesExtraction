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
$projectResources = $projectHome + "Resources"
$summaryRoutedAudit = $projectHome + "Audit" + $dirChar + "Summary Routed" + $dirChar

# Today's Date (In short format)
$today = Get-Date -Format d

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






$divisionResourcesFile = $projectHome + "Resources" + $dirChar + "rs_divlist.csv"
$divisionResourcesCSV = Import-Csv $divisionResourcesFile
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "DataEntryFileName" -Value $null
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "Data Entry Link" -Value $null
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "BillingFileName" -Value $null
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "Billing and Collections Link" -Value $null


$divisionResourcesCSV | ForEach-Object {
<# Replaced the "adaptive" code below for fixing the data to hard-coded Area information, there's too many inconsistencies in the website directory structure to formulaically process using the code below. Keeping it for informational/reference purposes
    $_.Area = $_.Area -replace ': ','-'
    $AreaWeb = $_.Area -replace '-','_'
    $AreaWeb = $AreaWeb -replace ' ','%20'
    $AreaWebSub = $AreaWeb + "/"
    $areaSpaceUnderscore = $_.Area -replace ' ','_'
    $areaDashUnderscore = $_.Area -replace '-','_'
    $areaOnlyUnderscore = $areaSpaceUnderscore -replace '-','_'
    $AreaOnly = $_.Area.Substring(4)
    $AreaOnlyUnderscore = $areaSpaceUnderscore.Substring(4)
#>
    $DataEntryName = "DIV_" + $_."Division #" + "_Data_Entry_Information.html"
    $BillingName = "DIV_" + $_."Division #" + "_Billing_and_Collection_Information.html"

    $compassWebLocationDiv = "http://compass.repsrv.com/DivisionalDocuments/"
    $genDocsFileBilling = $compassWebLocationDiv + $_.Area + "/" + $BillingName
    $genDocsFileDataEntry = $compassWebLocationDiv + $_.Area + "/" + $DataEntryName
    $_."Data Entry Link" = $genDocsFileDataEntry
    $_."Billing and Collections Link" = $genDocsFileBilling
    $_.DataEntryFileName = $DataEntryName
    $_.BillingFileName = $BillingName
}

# Get list of documents
$startListPath = $projectHome + "feesExtractor.csv"
$startList = Test-Path $startListPath
if ($startList) {
        Write-Host Processing list of documents
    } else {
        Write-Output "ERROR: DOCUMENT LIST DOES NOT EXIST! Export the document list from Excel in .csv (Comma Separated Values) format and save it to the prepared folder" | Out-File -Append $errorLog
        Write-Host "ERROR: DOCUMENT LIST DOES NOT EXIST! Export the document list from Excel in .csv (Comma Separated Values) format and save it to the prepared folder" -ForegroundColor white -BackgroundColor Red
        exit
    }


# Create the imported data array from the starting list of documents to process
$importedDocumentList = Import-Csv $startListPath

$importedDocumentList | Add-Member -MemberType NoteProperty -Name CityOrEntity -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name State -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name Assigned -Value $env:USERNAME
$importedDocumentList | Add-Member -MemberType NoteProperty -Name "Link" -Value $null -Force
$importedDocumentList | Add-Member -MemberType NoteProperty -Name "LinkSuccess" -Value $null

# Pulls the list of documents to be worked 
$importedDocumentList | ForEach-Object {

    $_.Name = $_.Name -replace " ","%20"

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

<# Changed process to no longer save the pages to the local disk.
    $fileSaveLocation = $projectTemp + $_.Name
    [io.file]::WriteAllText($fileSaveLocation,(Invoke-WebRequest -URI $compassUrl -UseDefaultCredentials -UseBasicParsing).content)
#>

}



<# OLD CODE - Created as a means to create a data table in memory and I didn't know how to do so without creating a csv on the disk first. See below the commented section for the newer, refined code.
$tempTable = $projectTemp + "temp-table.csv"
Add-Content -Path $tempTable -Value 'DivNum,DataEntryFileName,DataEntryLink,BillingFileName,BillingLink'

$divisionList | ForEach-Object {

    $divNoTemp = $_
    $dataEntryFilenameTemp = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Select-Object DataEntryFileName | Select -ExpandProperty "DataEntryFileName"
    $divDataItem1 = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Select-Object "Data Entry Link" | Select -ExpandProperty "Data Entry Link"
    $billingFilenameTemp = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Select-Object BillingFileName | Select -ExpandProperty "BillingFileName"
    $divDataItem2 = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Select-Object "Billing and Collections Link" | Select -ExpandProperty "Billing and Collections Link"

    $rowToTemp = $divNoTemp + "," + $dataEntryFilenameTemp + "," + $divDataItem1 + "," + $billingFilenameTemp + "," + $divDataItem2
    Add-Content -Path $tempTable -Value $rowToTemp

}


$divisionListWithLinks = Import-Csv $tempTable
$divisionListWithLinks | Add-Member -MemberType NoteProperty -Name "NumOfInv_Groups" -Value $null
$divisionListWithLinks | Add-Member -MemberType NoteProperty -Name "Inv_Groups_Symbols" -Value $null

# Clean up the temporary data
Remove-Item -Path $tempTable -Force

$divisionListWithLinks | ForEach-Object {
    

    $fileSaveLocationDataEntry = $projectTemp + $_.DataEntryFileName
    [io.file]::WriteAllText($fileSaveLocation,(Invoke-WebRequest -URI $_.DataEntryLink -UseDefaultCredentials -UseBasicParsing).content)
    $fileSaveLocationBilling = $projectTemp + $_.BillingFileName
    [io.file]::WriteAllText($fileSaveLocation,(Invoke-WebRequest -URI $_.BillingLink -UseDefaultCredentials -UseBasicParsing).content)

    # logic to use user input to get the invoice groups goes here
    start chrome $_.BillingLink
    $countInvGroups = Read-Host -Prompt 'How many Invoice Groups are in Billing Information - Residential/Invoice Group for the Division (see the browser that just popped up)'
    $symbolsInvGroups = Read-Host -Prompt 'What are the invoice group symbols? (ex. A,B,C,1,2,3) Separate multiple symbols with commas'
    $_.NumOfInv_Groups = $countInvGroups
    $_.Inv_Groups_Symbols = $symbolsInvGroups
    
}


#>

# Creates a subset list of $divisionResourcesCSV that strips out the data that's not being used for this session
$divisionList = $importedDocumentList.DivisionNo | Select-Object -Unique
$divisionListWithLinks = @()
$divisionList | ForEach-Object {
    
    # Gather Data
    $tempDivNo = $_
    $dataEntryFilenameTemp = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $tempDivNo } | Select-Object DataEntryFileName | Select -ExpandProperty "DataEntryFileName"
    $divDataEntryLink = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $tempDivNo } | Select-Object "Data Entry Link" | Select -ExpandProperty "Data Entry Link"
    $billingFilenameTemp = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $tempDivNo } | Select-Object BillingFileName | Select -ExpandProperty "BillingFileName"
    $divBillingLink = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $tempDivNo } | Select-Object "Billing and Collections Link" | Select -ExpandProperty "Billing and Collections Link"

    # Assign data to a new table item
    $divListObject = New-Object PSObject
    $divListObject | Add-Member -MemberType NoteProperty -Name "DivNum" -Value $_ # Reminder, $_ is the 'shortcut' variable that tells Powershell to use the current item being processed by the ForEach-Object loop as the value
    $divListObject | Add-Member -MemberType NoteProperty -Name "DataEntryFileName" -Value $dataEntryFilenameTemp
    $divListObject | Add-Member -MemberType NoteProperty -Name "DataEntryLink" -Value $divDataEntryLink
    $divListObject | Add-Member -MemberType NoteProperty -Name "BillingFileName" -Value $billingFilenameTemp
    $divListObject | Add-Member -MemberType NoteProperty -Name "BillingLink" -Value $divBillingLink
    $divListObject | Add-Member -MemberType NoteProperty -Name "NumOfInv_Groups" -Value $null
    $divListObject | Add-Member -MemberType NoteProperty -Name "Inv_Groups_Symbols" -Value $null
    
    # Add the new table item to the table
    $divisionListWithLinks += $divListObject
}


<# Keeping this code for later use in building a dedicated "Find all summary routed" tool, will need to re-think and rebuild this section

#Prepare data for the tracker
$trackerExportRefined = $projectHome + "RefinedExport.csv"

$trackerExport = @()
$importedDocumentList | ForEach-Object {
    
    # Getting the count of the container sizes
    $countingOpen = $projectTemp + $_.Name
    start chrome $_.Link
    $countOfContainerSizes = Read-Host -Prompt 'How many container sizes are there?'
    # Retrieving count from $divisionListWithLinks
    $divCompare = $_.DivisionNo
    $preCount = $divisionListWithLinks | Where-Object { $_.DivNum -eq $divCompare } | Select-Object NumOfInv_Groups | Select -ExpandProperty "NumOfInv_Groups"
    $totalLinesPerCode = [int]$countOfContainerSizes * [int]$preCount
    
    # $summaryRoutedValue = 0
    $summaryRoutedFile = $projectTemp + $_.Name
        # Solid Waste
        $summaryRoutedStringSearchSW = Select-String -Path $summaryRoutedFile -Pattern "Solid\s+Waste\s+Rate" -Context 0,2
        $summaryRoutedStringSearchSW = $summaryRoutedStringSearchSW -split "</td>" | Where { $_ -notmatch "Rate" }
        $summaryRoutedStringSearchSW = $summaryRoutedStringSearchSW -split ">" | Where { $_ -notmatch "class" }
        if ($summaryRoutedStringSearchSW -Match "table") {
            $summaryRoutedValueSW = 1
            } else {
            $summaryRoutedValueSW = 0
        }
        # Solid Waste Additional Container
        $summaryRoutedStringSearchSWAdd = Select-String -Path $summaryRoutedFile -Pattern "Solid\s+Waste\s+Additional\s+Container\s+Rental\s+Rate" -Context 0,2
        $summaryRoutedStringSearchSWAdd = $summaryRoutedStringSearchSWAdd -split "</td>" | Where { $_ -notmatch "Rate" }
        $summaryRoutedStringSearchSWAdd = $summaryRoutedStringSearchSWAdd -split ">" | Where { $_ -notmatch "class" }
        if ($summaryRoutedStringSearchSWAdd -Match "table") {
            $summaryRoutedValueSWAdd = 1
            } else {
            $summaryRoutedValueSWAdd = 0
        }
        # Recycle
        $summaryRoutedStringSearchREC = Select-String -Path $summaryRoutedFile -Pattern "Solid\s+Waste\s+Rate" -Context 0,2
        $summaryRoutedStringSearchREC = $summaryRoutedStringSearchREC -split "</td>" | Where { $_ -notmatch "Rate" }
        $summaryRoutedStringSearchREC = $summaryRoutedStringSearchREC -split ">" | Where { $_ -notmatch "class" }
        if ($summaryRoutedStringSearchREC -Match "table") {
            $summaryRoutedValueREC = 1
            } else {
            $summaryRoutedValueREC = 0
        }
        # Recycle Additional Container
        $summaryRoutedStringSearchRECAdd = Select-String -Path $summaryRoutedFile -Pattern "REC\s+Additional\s+Container\s+Rental\s+Rate" -Context 0,2
        $summaryRoutedStringSearchRECAdd = $summaryRoutedStringSearchRECAdd -split "</td>" | Where { $_ -notmatch "Rate" }
        $summaryRoutedStringSearchRECAdd = $summaryRoutedStringSearchRECAdd -split ">" | Where { $_ -notmatch "class" }
        if ($summaryRoutedStringSearchRECAdd -Match "table") {
            $summaryRoutedValueRECAdd = 1
            } else {
            $summaryRoutedValueRECAdd = 0
        }

        # Yard Waste
        $summaryRoutedStringSearchYW = Select-String -Path $summaryRoutedFile -Pattern "Solid\s+Waste\s+Rate" -Context 0,2
        $summaryRoutedStringSearchYW = $summaryRoutedStringSearchYW -split "</td>" | Where { $_ -notmatch "Rate" }
        $summaryRoutedStringSearchYW = $summaryRoutedStringSearchYW -split ">" | Where { $_ -notmatch "class" }
        if ($summaryRoutedStringSearchYW -Match "table") {
            $summaryRoutedValueYW = 1
            } else {
            $summaryRoutedValueYW = 0
        }

        # Yard Waste Additional Container
        $summaryRoutedStringSearchYWAdd = Select-String -Path $summaryRoutedFile -Pattern "Additional\s+YW\s+Container\s+Rental\s+Rate" -Context 0,2
        $summaryRoutedStringSearchYWAdd = $summaryRoutedStringSearchYWAdd -split "</td>" | Where { $_ -notmatch "Rate" }
        $summaryRoutedStringSearchYWAdd = $summaryRoutedStringSearchYWAdd -split ">" | Where { $_ -notmatch "class" }
        if ($summaryRoutedStringSearchYWAdd -Match "table") {
            $summaryRoutedValueYWAdd = 1
            } else {
            $summaryRoutedValueYWAdd = 0
        }



        $summaryRoutedValue = $summaryRoutedValueSW + $summaryRoutedValueSWAdd + $summaryRoutedValueREC + $summaryRoutedValueRECAdd + $summaryRoutedValueYW + $summaryRoutedValueYWAdd
    
    $trackerObject = New-Object PSObject
    $documentPath = $projectTemp + $_.Name

    $trackerObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name
    $trackerObject | Add-Member -MemberType NoteProperty -Name "LawsonDiv" -Value $_.LawsonDiv
    $trackerObject | Add-Member -MemberType NoteProperty -Name "DivisionNo" -Value $_.DivisionNo
    $trackerObject | Add-Member -MemberType NoteProperty -Name "PolygonID" -Value $_.PolygonID
    $trackerObject | Add-Member -MemberType NoteProperty -Name "Rate ID" -Value $_."Rate ID"
    $trackerObject | Add-Member -MemberType NoteProperty -Name "Area" -Value $_.Area
    $trackerObject | Add-Member -MemberType NoteProperty -Name "TotalLines" -Value $totalLinesPerCode
    $tempFileExist = $projectTemp + $_.Name
    $fileExistTest = Test-Path $tempFileExist
    if ($fileExistTest) {
        $trackerObject | Add-Member -MemberType NoteProperty -Name "In Compass" -Value "Yes"
        } else {
        $trackerObject | Add-Member -MemberType NoteProperty -Name "In Compass" -Value "No"
    }
    if ($summaryRoutedValue -eq 0) {
        $trackerObject | Add-Member -MemberType NoteProperty -Name "Form Completed" -Value $today
        $trackerObject | Add-Member -MemberType NoteProperty -Name "Notes" -Value "Summary Routed Account"
        Copy-Item $documentPath $summaryRoutedAudit -Force
        
        # Count of Lines Needed
        $billingDoc = $projectTemp + $_.Name
        #$invoiceGroupPrep = Select-String -Path 


        } else {
        $trackerObject | Add-Member -MemberType NoteProperty -Name "Form Completed" -Value $null
        $trackerObject | Add-Member -MemberType NoteProperty -Name "Notes" -Value $null
    }

    $trackerExport += $trackerObject
}
$trackerExport | Export-Csv -Path $trackerExportRefined -NoTypeInformation
#>


<# Broken code, not sure why the for() stepper is not working

$rateIDList = $importedDocumentList."Rate ID" | Select-Object -Unique

$dataExport = @()
$rateIDList | ForEach-Object {
    
    $rateIDalpha = $_
    $lineCount = $trackerExport | Where-Object { $_."Rate ID" -eq $rateIDalpha } | Select-Object TotalLines | Select -ExpandProperty "TotalLines"

    for ($i = $lineCount, $i -ne 0, $i--) {
    
        $dataExportObject = New-Object PSObject
        $dataExportObject | Add-Member -MemberType NoteProperty -Name "Rate_ID" -Value $_

        $dataExport += $dataExportObject
    }
}
        <# $importedDocumentList | Add-Member -MemberType NoteProperty -Name Downloaded -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name "Row Type" -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name lawson_infopro_polygon -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name is_cust_owned -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name serviceType -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name is_this_row_additional_cont -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name marketType -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name stop_cd -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name rate_type -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name district_cd -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name charge_cd -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name serv_freq -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name contract_nbr -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name contract_grp -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name quantity_threshold -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name "MIN quantity" -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name "MAX quantity" -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name "Rate applies per item" -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name size -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name cont_type -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name extra_unit_type -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name extra_units -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name extra_units_size -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name delivery -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name removal -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name serv_int_fee -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name cont_rep_fe -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name cont_exch_fee -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name Late_fee -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name other_fees -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name other_fees_desc -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name FRF -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name FRF_exempt_cd -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name ERF -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name ERF_exempt_cd -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name Admin -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name Admin_exempt_cd -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name service_reinst_fee -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name bill_freq -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name price -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name discounted_rate -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name save_rate -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name account_type -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name rev_dist_code -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name month_in_adv -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name ACQ_code -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name invoice_group -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name is_CRC_Consolidated -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name sold_as -Value $null
$importedDocumentList | Add-Member -MemberType NoteProperty -Name bundle_id -Value $null
# $importedDocumentList | Add-Member -MemberType NoteProperty -Name Notes -Value $null

#



$dataExportFileLocation = $projectHome + "DataExport.csv"
$dataExport | Export-Csv -Path $dataExportFileLocation -NoTypeInformation
#>

<#
# Open the divisional documents in Chrome
$divisionListWithLinks | ForEach-Object {

    start chrome $_.BillingLink
    #start chrome $_.DataEntryLink

}
#>

$runTestPrompt = Read-Host -Prompt "Test to make sure the documents will open? (Y/N only, default is 'N')"

if ($runTestPrompt -eq "Y") {
# Open the Resi documents in Chrome
$importedDocumentList | ForEach-Object {

    start chrome $_.Link
    $linkTest = Read-Host -Prompt "Did the document open? (Y/N only, default is 'N')"
    if ($linkTest -eq "Y") {
        $_.LinkSuccess = "Y"    
    } else {
        $_.LinkSuccess = "N"
    }
}

# Return a list of documents that did not open
$failedURLFilePath = $projectHome + $dirChar + "FailedToOpen.csv"
$failedURLs = $importedDocumentList | Where-Object { $_.LinkSuccess -eq "N" }
$failedURLs | Export-Csv -Path $failedURLFilePath -NoTypeInformation

}


# Load Data Entry information from Resources files into system memory
$dataEntryTableTempVar = $projectResources + $dirChar + "dataentry.csv"
$dataEntryTable = Import-Csv $dataEntryTableTempVar
#Fix the City capitalization

$dataEntryTable | ForEach-Object {
    $words = $_.City
    $TextInfo = (Get-Culture).TextInfo
    $fixedCity = $TextInfo.ToTitleCase($words.ToLower())
    $_.City = $fixedCity
}

# Check to see if any new divisions need entering into the Data Entry table
$tempList1 = $dataEntryTable.DIV | Select-Object -Unique
$tempList2 = $importedDocumentList.DivisionNo | Select-Object -Unique
$tempTestResult = Compare-Object $tempList1 $tempList2 | Where-Object SideIndicator -eq "=>" | ForEach-Object { $_.InputObject }
if ($tempTestResult -gt 0) {
    Write-Host "New divisions found in assigned work, retrieving Data Entry documents and HTML Table to CSV conversion tool" -BackgroundColor Red -ForegroundColor Green
    $tempTestResult | ForEach-Object {
        $openDataEntry = $_
        $openDataEntryLink = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $openDataEntry } | Select-Object "Data Entry Link" | Select -ExpandProperty "Data Entry Link"
        start chrome $openDataEntryLink
    }
    start chrome http://convertcsv.com/html-table-to-csv.htm
}


# A custom function to quickly find the Rev_Dist_Code in the Data Entry table
Function ResiCodeLookup ($city, $state) {
    $tmpTable = $dataEntryTable | Where-Object { $_.City -eq $city }
    $tmpTable | Where-Object { $_.State -eq $state } | Select-Object DIV,Lawson,City,State,Residential
}

# A custom function to quickly open a Data Entry document when only the division number is known
Function OpenDataEntry ($divnum) {
    $tmpLink = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divnum } | Select-Object "Data Entry Link" | Select -ExpandProperty "Data Entry Link"
    start chrome $tmpLink
}

Function QuickDocsOpen ($docName) {
    $divNum = $docName.Substring(4,3)
    $tmpBillingLink = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNum } | Select-Object "Billing and Collections Link" | Select -ExpandProperty "Billing and Collections Link"
    $tmpDocLink = $importedDocumentList | Where-Object { $_.Name -eq $docName } | Select-Object Link | Select -ExpandProperty "Link"
    start chrome $tmpBillingLink
    start chrome $tmpDocLink
}

Function OpenWholeDivision ($divNum) {
 
    $billingLinkTemp = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNum } | Select-Object "Billing and Collections Link" | Select -ExpandProperty "Billing and Collections Link"
    $sessionOpenListTemp = $importedDocumentList | Where-Object { $_.DivisionNo -eq $divNum }
    $sessionOpenList = $sessionOpenListTemp | Where-Object { [string]::IsNullOrEmpty($_."Form Completed") -eq "True" }
    $sessionOpenList | ForEach-Object {
        start chrome $_.Link
    }
    start chrome $billingLinkTemp
}

Function RefreshDataEntry {
    # Load Data Entry information from Resources files into system memory
    $dataEntryTableTempVar = $projectResources + $dirChar + "dataentry.csv"
    $dataEntryTable = Import-Csv $dataEntryTableTempVar
    #Fix the City capitalization

    $dataEntryTable | ForEach-Object {
        $words = $_.City
        $TextInfo = (Get-Culture).TextInfo
        $fixedCity = $TextInfo.ToTitleCase($words.ToLower())
        $_.City = $fixedCity
    }
}