#Disk Capacity Tool
#Tool Usage .\computer-grab.ps1 -computersFile yourfile.txt
param(
	[string]$computersFile
	);
	
#pull in entire contents of file and store in computerList
$computerList = get-content $computersFile
#instantiate the excel object
$excel = New-Object -ComObject Excel.Application
#set global parameters for excel
$Excel.Visible = $True
$workbook = $Excel.Workbooks.Add()
$sheet = $workbook.ActiveSheet

#intRow tracks which row to write the current entry, we start at 2 because the header row is at 1
$intRow = 2

#write the header row to the current worksheet
$sheet.cells.Item(1, 1) = "Machine Name"
$sheet.cells.Item(1, 2) = "OS Type"
$sheet.cells.Item(1, 3) = "IP Address"
$sheet.cells.Item(1, 4) = "Computer Description"
$sheet.cells.Item(1, 5) = "Drive"
$sheet.cells.Item(1, 6) = "File System"
$sheet.cells.Item(1, 7) = "Description"
$sheet.cells.Item(1, 8) = "Total Size (GBs)"
$sheet.cells.Item(1, 9) = "Free Space (GBs)"
$sheet.cells.Item(1, 10) = "Free Space Percentage"
$sheet.cells.Item(1, 11) = "Used Space (GBs)"
$sheet.cells.Item(1, 12) = "Memory (GBs)"
$sheet.cells.Item(1, 13) = "CPU"
#$sheet.cells.Item(1, 14) = "Software"

#crawl through each computer entry in $computerList file
foreach ($computer in $ComputerList)
{
	#ping test connection if the computer is even reachable
	If(test-connection -computer $computer -count 1 -quiet)
	{
		#If the computer is reachable, is WMI available?
		if((Get-wmiobject Win32_TimeZone -computer $computer -erroraction SilentlyContinue).caption)
		{
			#Highlight the row that contains the computer
			$cellrange = $sheet.Range("A$intRow","N$intRow")
			$cellrange.Interior.ColorIndex = 33
			$cellrange.Font.ColorIndex = 2
			#Grab WMI information
			$sheet.cells.Item($intRow, 1) = ($computer.toupper())
			$colDisks = Get-WmiObject Win32_LogicalDisk -computer $computer -filter "DriveType=3" -erroraction SilentlyContinue
			$colMemory = (Get-WmiObject Win32_PhysicalMemory -computer $computer -property Capacity -erroraction SilentlyContinue).Capacity
			$colCPU = (Get-WmiObject Win32_Processor -computer $computer -property Name -erroraction SilentlyContinue).name
			$colOS = (Get-WmiObject Win32_OperatingSystem -computer $computer -property Caption -erroraction SilentlyContinue).caption
			#$colSoftware = Get-WmiObject Win32_Product -computer $computer -erroraction SilentlyContinue
			$colIPAddr = (Get-WmiObject Win32_NetworkAdapterConfiguration -computer $computer -filter "IPEnabled = True"	-property IPAddress -erroraction SilentlyContinue ).ipaddress
			#begin writing WMI information into spreadsheet
			$sheet.cells.Item($intRow, 2) = ($colOS.toUpper())
			$sheet.cells.Item($intRow, 4) = ((get-adcomputer -filter 'name -like $computer' -properties description).description)
			#A computer may have multiple NICs or IP addresses
			#multiple addresses will be entered in separate rows and the the current counter incremented at the end.
			$ipRow = $intRow
			ForEach($ipaddr in $colIPAddr)
			{
				if($ipRow -gt $intRow)
				{
					$cellrange = $sheet.Range("A$ipRow","N$ipRow")
					$cellrange.Interior.ColorIndex = 34
				}
				$sheet.cells.Item($ipRow, 3) = $ipaddr
				$ipRow += 1
			}
			#For computers with multiple cpus, each cpu will occupy one row, the counter is unique to this object
			$cpuRow = $intRow
			ForEach($cpu in $colCPU)
			{
				if($cpuRow -gt $intRow)
				{
					$cellrange = $sheet.Range("A$cpuRow","N$cpuRow")
					$cellrange.Interior.ColorIndex = 34
				}
				$sheet.cells.Item($cpuRow, 13) = $cpu
				$cpuRow += 1
			}
			#The memory will be in multiple chunks, add them all together to get total memory
			$totalMemory = 0
			ForEach($objMemory In $colMemory)
			{
				$totalMemory += $objMemory
			}
			$sheet.cells.Item($intRow, 12) = ($totalMemory/1GB)
			#Each disk will have its information written on a separate row
			$diskRow = $intRow
			ForEach($objDisk In $colDisks)
			{
				if($diskRow -gt $intRow)
				{
					$cellrange = $sheet.Range("A$diskRow","N$diskRow")
					$cellrange.Interior.ColorIndex = 34
				}
				$sheet.cells.Item($diskRow, 5) = $objDisk.DeviceID
				$sheet.cells.Item($diskRow, 6) = ($objDisk.FileSystem).toupper()
				$sheet.cells.Item($diskRow, 7) = $objDisk.VolumeName
				$sheet.cells.Item($diskRow, 8) = ($objDisk.Size/1GB)
				$sheet.cells.Item($diskRow, 9) = ($objDisk.FreeSpace)/1GB
				$sheet.cells.Item($diskRow, 10) = "{0:P0}" -f (($objDisk.FreeSpace)/($objDisk.Size))
				$sheet.cells.Item($diskRow, 11) = (($objDisk.Size) - ($objDisk.FreeSpace))/1GB
				$diskRow += 1
			}
			#increment the row counter and then determine if there were other rows written.
			#If other rows were written and the counter is larger than the next row, we don't want to overwrite the data,
			#so we increment to the largest row
			$intRow += 1
			if($ipRow -gt $intRow)
			{
				$intRow = $ipRow
			}
			if($cpuRow -gt $intRow)
			{
				$intRow = $cpuRow
			}
			if($diskRow -gt $intRow)
			{
				$intRow = $diskRow
			}
		}
		Else
		{
			#if a computer responds to pings but not wmi, it's probably not windows
			$sheet.cells.Item($intRow, 1) = $computer.toUpper()
			$sheet.cells.Item($intRow, 3) = "Non-Windows Computer"
			$intRow += 1
		}
	}
	Else
	{
		#The computer did not respond to pings at all
		$sheet.cells.Item($intRow, 1) = $computer.toUpper()
		$sheet.cells.Item($intRow, 3) = "No Ping Response"
		$intRow += 1
	}
}
#Select the top row, change the font and cell color, bold the font, and autofit the entire worksheet.
$range = $sheet.Range("A1","N1")
$range.Interior.ColorIndex = 19
$range.Font.ColorIndex = 14
$range.Font.Bold = $True
$sheet.Columns.Autofit()
#Freeze the top row of the worksheet
#$sheet.application.activewindow.freezepanes = $True

