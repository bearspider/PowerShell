#VMWare Report Query
#Tool Usage .\computer-grab.ps1 -computersFile yourfile.txt
param(
	[array]$vcenters,
	[string]$computersFile
	);

$date = get-date
# Adds the base cmdlets
if((Get-PSSnapin -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) -eq $null)
{
	Add-PSSnapin VMware.VimAutomation.Core
}

function Write-Header ($worksheet)
{
	$worksheet.cells.Item(1, 1) = "Machine Name"
	$worksheet.cells.Item(1, 2) = "Power Status"
	$worksheet.cells.Item(1, 3) = "Notes"
	$range = $worksheet.Range("A1","N1")
	$range.Interior.ColorIndex = 19
	$range.Font.ColorIndex = 14
	$range.Font.Bold = $True
	$worksheet.application.activewindow.splitcolumn=0
	$worksheet.application.activewindow.splitrow=1
	$worksheet.application.activewindow.freezepanes = $True
}

#####################################################################################################################	
#							MAIN PROGRAM																			#
#####################################################################################################################
#pull in entire contents of file and store in computerList
$computerList = get-content $computersFile
#instantiate the excel object
$excel = New-Object -ComObject Excel.Application
#set global parameters for excel
$Excel.Visible = $True
$workbook = $Excel.Workbooks.Add()
$sheet1 = $workbook.ActiveSheet
$sheet1.name = "VMWare"

#intRow tracks which row to write the current entry, we start at 2 because the header row is at 1
$wRow = 2

#write the header row to the worksheets
write-header -worksheet $sheet1

disconnect-viserver -erroraction silentlycontinue -server * -Force -Confirm:$False
Set-PowerCliConfiguration  -Confirm:$false -DefaultVIServerMode Multiple
ForEach ($vcenter in $vcenters)
{
    connect-viserver $vcenter
	ForEach ($computer in $computerList)
	{
		write-host "Looking up $computer"
		$sheet1.cells.Item($wRow, 1) = ($computer.toupper())
		if($vmFile = get-vm -erroraction SilentlyContinue $computer)
		{
			$sheet1.cells.Item($wRow, 2) = [string]$vmFile.PowerState
			$sheet1.cells.Item($wRow, 3) = $vmFile | select-object -expandproperty notes
		}
		else
		{
			$sheet1.cells.Item($wRow, 2) = "Not in VMware"
		}
		$wRow += 1
	}
}

#autofit the entire worksheet.
(($sheet1.UsedRange).EntireColumn.Autofit()) | out-null
