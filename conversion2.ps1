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


$importDataSet = Import-Csv $startListPath

<#NOTE -

Move most (or ALL) of these fields to a new table object

#>

$importDataSet | Add-Member -MemberType NoteProperty -Name CityOrEntity -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name State -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Assigned -Value $env:USERNAME
# $importDataSet | Add-Member -MemberType NoteProperty -Name "Form Completed" -Value Schrodinger
# $importDataSet | Add-Member -MemberType NoteProperty -Name Notes -Value $null

<# $importDataSet | Add-Member -MemberType NoteProperty -Name Downloaded -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name "Row Type" -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name lawson_infopro_polygon -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name is_cust_owned -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name serviceType -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name is_this_row_additional_cont -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name marketType -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name stop_cd -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name rate_type -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name district_cd -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name charge_cd -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name serv_freq -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name contract_nbr -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name contract_grp -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name quantity_threshold -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name "MIN quantity" -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name "MAX quantity" -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name "Rate applies per item" -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name size -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name cont_type -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name extra_unit_type -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name extra_units -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name extra_units_size -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name delivery -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name removal -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name serv_int_fee -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name cont_rep_fe -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name cont_exch_fee -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Late_fee -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name other_fees -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name other_fees_desc -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name FRF -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name FRF_exempt_cd -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name ERF -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name ERF_exempt_cd -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Admin -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Admin_exempt_cd -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name service_reinst_fee -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name bill_freq -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name price -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name discounted_rate -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name save_rate -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name account_type -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name rev_dist_code -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name month_in_adv -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name ACQ_code -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name invoice_group -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name is_CRC_Consolidated -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name sold_as -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name bundle_id -Value $null
# $importDataSet | Add-Member -MemberType NoteProperty -Name Notes -Value $null

#>

$divisionResourcesFile = $projectHome + "Resources" + $dirChar + "rs_divlist.csv"
$divisionResourcesCSV = Import-Csv $divisionResourcesFile
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "DataEntryFileName" -Value $null
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "Data Entry Link" -Value $null
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "BillingFileName" -Value $null
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "Billing and Collections Link" -Value $null


$divisionResourcesCSV | ForEach-Object {

 $_.Area = $_.Area -replace ': ','-'

    $AreaWeb = $_.Area -replace ' ','%20'
    $AreaWebSub = $AreaWeb + "/"
    $areaSpaceUnderscore = $_.Area -replace ' ','_'
    $areaDashUnderscore = $_.Area -replace '-','_'
    $areaOnlyUnderscore = $areaSpaceUnderscore -replace '-','_'
    $AreaOnly = $_.Area.Substring(4)
    $AreaOnlyUnderscore = $areaSpaceUnderscore.Substring(4)

    $DataEntryName = "DIV_" + $_."Division #" + "_Data_Entry_Information.html"
    $BillingName = "DIV_" + $_."Division #" + "_Billing_and_Collection_Information.html"

    $compassWebLocationDiv = "http://compass.repsrv.com/DivisionalDocuments/"
    $genDocsFileBilling = $compassWebLocationDiv + $AreaWeb + "/DIV_" + $_."Division #" + "_Billing_and_Collection_Information.html"
    $genDocsFileDataEntry = $compassWebLocationDiv + $AreaWeb + "/" + $DataEntryName
    $_."Data Entry Link" = $genDocsFileDataEntry
    $_."Billing and Collections Link" = $genDocsFileBilling
    $_.DataEntryFileName = $DataEntryName
    $_.BillingFileName = $BillingName
}


$importDataSet | ForEach-Object {

    $_.Name = $_.Name -replace " ","_"
    $_.Name = $_.Name -replace "__","_"

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


}

$divisionList = $importDataSet.DivisionNo | Select-Object -Unique
$tempTable = $projectTemp + "temp-table.csv"
Add-Content -Path $tempTable -Value 'DivNum,DataEntryFileName,DataEntryLink,BillingFileName,BillingLink'

$divisionList | ForEach-Object {

    $divNoTemp = $_
    $dataEntryFilenameTemp = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Get-Member -Name "DataEntryFileName"
    $dataEntryFilenameTemp = $dataEntryFilenameTemp -split "=" | Where { $_ -notmatch "string" }
    $divDataItem1 = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Get-Member -Name "Data Entry Link"
    $divDataItem1 = $divDataItem1 -split "=" | Where { $_ -notmatch "string" }
    $billingFilenameTemp = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Get-Member -Name "BillingFileName"
    $billingFilenameTemp = $billingFilenameTemp -split "=" | Where { $_ -notmatch "string" }
    $divDataItem2 = $divisionResourcesCSV | Where-Object { $_."Division #" -eq $divNoTemp } | Get-Member -Name "Billing and Collections Link"
    $divDataItem2 = $divDataItem2 -split "=" | Where { $_ -notmatch "string" }


    $rowToTemp = $divNoTemp + "," + $dataEntryFilenameTemp + "," + $divDataItem1 + "," + $billingFilenameTemp + "," + $divDataItem2
    Add-Content -Path $tempTable -Value $rowToTemp

}


$divisionListWithLinks = Import-Csv $tempTable
$divisionListWithLinks | Add-Member -MemberType NoteProperty -Name "NumOfInv_Groups" -Value $null
$divisionListWithLinks | Add-Member -MemberType NoteProperty -Name "Inv_Groups_Symbols" -Value $null


$divisionListWithLinks | ForEach-Object {
    

    $fileSaveLocationDataEntry = $projectTemp + $_.DataEntryFileName
    [io.file]::WriteAllText($fileSaveLocation,(Invoke-WebRequest -URI $_.DataEntryLink -UseDefaultCredentials -UseBasicParsing).content)
    $fileSaveLocationBilling = $projectTemp + $_.BillingFileName
    [io.file]::WriteAllText($fileSaveLocation,(Invoke-WebRequest -URI $_.BillingLink -UseDefaultCredentials -UseBasicParsing).content)

    # logic to use user input to get the invoice groups goes here
    start chrome $fileSaveLocationBilling
    $countInvGroups = Read-Host -Prompt 'How many Invoice Groups are in Billing Information - Residential/Invoice Group for Division' $_.DivNum
    $symbolsInvGroups = Read-Host -Prompt 'What are the invoice group symbols? (ex. A,B,C,1,2,3) Separate multiple symbols with commas'
    $_.NumOfInv_Groups = $countInvGroups
    $_.Inv_Groups_Symbols = $symbolsInvGroups
}



#Prepare data for the tracker
$trackerExportRefined = $projectHome + "RefinedExport.csv"

$trackerExport = @()
$importDataSet | ForEach-Object {

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

<#
Area
Link
In Compass
Form Completed
Notes
#>

# Test the export, this is NOT the final product, that needs to be divided out by DIV
$trackerTest = $projectHome + "tracker-test.csv"
$compiledResults = $projectHome + "compiled-results.csv"
$importDataSet | Select-Object Name,LawsonDiv,DivisionNo,PolygonID,"Rate ID",CityOrEntity,State,Assigned,"Form Completed",Notes | Export-Csv -Append $trackerTest -NoTypeInformation #$areaTrackerFilename
$importDataSet | Select-Object "Row Type","Rate ID" | Export-Csv -Append -NoTypeInformation -Path $compiledResults