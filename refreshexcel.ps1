#Set the file path (can be a network location)
$filePath = "C:\Users\kazi.wadud\Desktop\TEST.xlsx"

#Create the Excel Object
$excelObj = New-Object -ComObject Excel.Application

#Make Excel visible. Set to $false if you want this done in the background
#$excelObj.Visible = $true

$excelObj.Visible = $false

#Open the workbook
$workBook = $excelObj.Workbooks.Open($filePath)

#Focus on the top row of the "Data" worksheet
#Note: This is only for visibility, it does not affect the data refresh
$workSheet = $workBook.Sheets.Item("Data")
$workSheet.Select()

#Refresh all data in this workbook
$workBook.RefreshAll()

#Save any changes done by the refresh
$workBook.Save()

#$workBook.Close()

#Uncomment this line if you want Excel to close on its own
#$excelObj.Quit()