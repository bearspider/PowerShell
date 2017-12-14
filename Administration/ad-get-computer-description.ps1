$DEBUG = 0
import-module activedirectory
$strPath="f:\scripts\server_report.xlsx"
$objExcel=New-Object -ComObject Excel.Application
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$objExcel.Visible=$false
$WorkBook=$objExcel.Workbooks.Open($strPath)
$worksheet = $workbook.sheets.item("Sheet1")
$intRowMax =  ($worksheet.UsedRange.Rows).count
$Columnnumber = 1
$description_column = 2
$saveas_path = "f:\scripts\server_report_complete.xlsx"

for($intRow = 1 ; $intRow -le $intRowMax ; $intRow++)
{	
	$name = $worksheet.cells.item($intRow,$ColumnNumber).value2	
	write-progress -Activity "Probing Active Directory" -status $name -percentComplete (($intRow / $intRowMax)*100)
	$short_name = $name.split('.')[0]	
	$description = get-adcomputer -filter "Name -Like '$short_name'" -properties description | select -expandproperty Description
	if(!$description)
	{
		$description = "Blank"
	}	
	if($DEBUG)
	{
		"Querying $name ..."
		"Short Name: $short_name"
		"Description: $description"
	}
	$worksheet.cells.item($intRow, $description_column) = $description

}
$workbook.saveas($saveas_path)
$objexcel.quit()
