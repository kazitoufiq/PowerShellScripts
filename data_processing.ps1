#To save password to txt file so that script can use it during run time
#read-host -assecurestring | convertfrom-securestring | out-file H:\cred.txt

Remove-Item "OPRDailyResults*.csv" -recurse
Remove-Item "IncidentSummaryReports*.csv" -recurse

$password = get-content H:\cred.txt | convertto-securestring

$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist "kazi.wadud",$password


$StartDate = (Get-Date).AddDays(-7).ToString('M/d/yyyy')
$EndDate = (Get-Date).AddDays(-1).ToString('M/d/yyyy')


$StartDate1 = (Get-Date).AddDays(-7).ToString('dd-MMM-yyyy')
$EndDate1 = (Get-Date).AddDays(-1).ToString('dd-MMM-yyyy')


$td = (Get-Date).ToString('dd-MMM-yyyy')
"Today is $td"


$pathdir='J:\xxxx\xxxx\11_Analytics\Weekly Performance Dashboard - Automated'

"Backing Up Master File --  OPR Daily Results"

Copy-Item -Path $pathdir'\OPR Daily Results.csv' -Destination $pathdir'\OPR Daily Results.Backup.csv' 

"Backing Up Master File --  Incident Summary Reports"

Copy-Item -Path $pathdir'\Incident Summary report.csv' -Destination $pathdir'\Incident Summary report.Backup.csv' 

"Downloading PRS data from $StartDate1 to $EndDate1"

$output = "J:\xxxx\xxxxx\11_Analytics\Weekly Performance Dashboard - Automated\IncidentSummaryReports_$StartDate1_$EndDate1 .csv"

$url = "http://xxxx/ReportServer?%2fPRS%2fIncident+Reports%2fIncident+Summary+report&rs:Command=Render&rs:Format=csv&datefrom=$StartDate%2012:00:00%20AM&dateto=$EndDate%2012:00:00%20AM"

Invoke-WebRequest -Uri $url -OutFile $output -Credential $credentials

"Download Completed Incident Summary Report"

"Append to Incident Master File"

Import-Csv $pathdir'\IncidentSummaryReports*.csv' | Export-Csv -Append $pathdir'\Incident Summary report.csv' -NoTypeInformation

#Import-Csv $pathdir'\Incident Summary report.csv' | Group Incident_No2,Incident_Date | Where {$_.Count -eq 1} | Select -expandProperty group | Export-Csv $pathdir'\Incident Summary report_uniq.csv' -NoTypeInformation


## To overcome the date format issue, need to open the CSV and then close it 
$xl = new-object -comobject excel.application
$xl.visible = $False
$Workbook = $xl.workbooks.open("$pathdir\Incident Summary report.csv")
$Worksheets = $Workbooks.worksheets
$Workbook.SaveAs("$pathdir\Incident Summary report_corrected.csv")
$Workbook.Saved = $True
$xl.Quit()


Import-Csv $pathdir'\Incident Summary report_corrected.csv' | sort Incident_No2,Incident_Date,Resp -unique | Export-Csv $pathdir'\Incident Summary report_unique.csv' -NoTypeInformation


#### OPR


$output = "J:\xxxx\xxx\11_Analytics\Weekly Performance Dashboard - Automated\OPRDailyResults_$StartDate1_$EndDate1 .csv"

$url="http://xxxx/ReportServer?%2fPRS%2fOPR+Daily+Results&rs:Command=Render&rs:format=csv&DateFrom=$StartDate&DateTo=$EndDate"

Invoke-WebRequest -Uri $url -OutFile $output -Credential $credentials

"Download Completed OPR Daily Results"

"Append to OPR Master File"

Import-Csv $pathdir'\OPRDailyResults*.csv' | Export-Csv -Append $pathdir'\OPR Daily Results.csv'  -NoTypeInformation


## To overcome the date format issue, need to open the CSV and then close it 
$xl = new-object -comobject excel.application
$xl.visible = $False
$Workbook = $xl.workbooks.open("$pathdir\OPR Daily Results.csv")
$Worksheets = $Workbooks.worksheets
$Workbook.SaveAs("$pathdir\OPR Daily Results_corrected.csv")
$Workbook.Saved = $True
$xl.Quit()

Import-Csv $pathdir'\OPR Daily Results_corrected.csv' | sort start_date -unique | Export-Csv $pathdir'\OPR Daily Results_unique.csv' -NoTypeInformation

#Import-Csv $pathdir'\OPR Daily Results.csv' | Group start_date | Where {$_.Count -eq 1} | Select -expandProperty group | Export-Csv $pathdir'\OPR Daily Results_uniq.csv' -NoTypeInformation



Remove-Item "OPR Daily Results.csv" -recurse
Remove-Item "OPR Daily Results_corrected.csv" -recurse
Rename-item "OPR Daily Results_unique.csv"  -NewName "OPR Daily Results.csv"



Remove-Item "Incident Summary report.csv" -recurse
Remove-Item "Incident Summary report_corrected.csv" -recurse 
Rename-item "Incident Summary report_unique.csv" -NewName "Incident Summary report.csv"



"Process Completed"
