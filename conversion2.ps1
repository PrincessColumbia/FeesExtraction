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

$importDataSet | Add-Member -MemberType NoteProperty -Name CityOrEntity -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name State -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Assigned -Value $env:USERNAME
$importDataSet | Add-Member -MemberType NoteProperty -Name "Form Completed" -Value Schrodinger
$importDataSet | Add-Member -MemberType NoteProperty -Name Notes -Value $null
$importDataSet | Add-Member -MemberType NoteProperty -Name Downloaded -Value $null
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
$importDataSet | Add-Member -MemberType NoteProperty -Name Notes -Value $null

$divisionResourcesFile = $projectHome + "Resources\rs_divlist_20170608.csv"
$divisionResourcesCSV = Import-Csv $divisionResourcesFile
$divisionResourcesCSV | Add-Member -MemberType NoteProperty -Name "Link" -Value $null
$divisionResourcesCSV | foreach {

 $_.Area = $_.Area -replace ': ','-'

}

$divisionList = $importDataSet.DivisionNo | Select-Object -Unique
New-Object System.Data.DataTable "$divListwithLinks"
$divListwithLinks | Add-Member -MemberType NoteProperty -Name "DivNum" -Value $null
$divListwithLinks | Add-Member -MemberType NoteProperty -Name "Link" -Value $null

$divisionList | foreach {

    $row = $divListwithLinks.NewRow()
    $row.DivNum = $_
    $row.Link = $_ + " test"

}

$importDataSet | foreach {

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



# Test the export, this is NOT the final product, that needs to be divided out by DIV
$trackerTest = $projectHome + "tracker-test.csv"
$compiledResults = $projectHome + "compiled-results.csv"
$importDataSet | Select-Object Name,LawsonDiv,DivisionNo,PolygonID,"Rate ID",CityOrEntity,State,Assigned,"Form Completed",Notes | Export-Csv -Append $trackerTest -NoTypeInformation #$areaTrackerFilename
$importDataSet | Select-Object "Row Type","Rate ID" | Export-Csv -Append -NoTypeInformation 