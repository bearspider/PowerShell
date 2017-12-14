$vcenters = @()
$emailaddress = ""
$license_table = @{}
$date = get-date
# Adds the base cmdlets
if((Get-PSSnapin -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) -eq $null)
{
	Add-PSSnapin VMware.VimAutomation.Core
}
# Set to multiple VC Mode 
if(((Get-PowerCLIConfiguration).DefaultVIServerMode) -ne "Multiple") { 
    Set-PowerCLIConfiguration -DefaultVIServerMode Multiple | Out-Null 
}

# Make sure you connect to your VCs here
ForEach ($vcenter in $vcenters)
{
    connect-viserver $vcenter
}
# Get the license info from each VC in turn 
$vSphereLicInfo = @() 
$ServiceInstance = Get-View ServiceInstance 
Foreach ($LicenseMan in Get-View ($ServiceInstance | Select -First 1).Content.LicenseManager) { 
    Foreach ($License in ($LicenseMan | Select -ExpandProperty Licenses )) { 
        $Details = "" |Select VC, Name, Key, Capacity, Assigned, Unit, ExpirationDate , Information 
        $Details.VC = ([Uri]$LicenseMan.Client.ServiceUrl).Host 
        $Details.Name = $License.Name 
        $Details.Key = $License.LicenseKey 
		$Details.Unit = $License.CostUnit
        $Details.Capacity= $License.Total 
        $Details.Assigned= $License.Used 
        $Details.Information= $License.Labels | Select -expand Value 
        $Details.ExpirationDate = $License.Properties | Where { $_.key -eq "expirationDate" } | Select -ExpandProperty Value 
        $vSphereLicInfo += $Details
		
		if($license_table.ContainsKey($Details.Name))
		{
			$temp_array = ($license_table.Get_Item($Details.Name))
			$Capacity = $temp_array[0] + $Details.Capacity
			$Assigned = $temp_array[1] + $Details.Assigned
			$push_array = @($Capacity,$Assigned,$Details.Unit)
			$license_table.Set_Item($Details.Name,$push_array)
		}
		else
		{
			$license_array = @($Details.Capacity,$Details.Assigned,$Details.Unit)
			$license_table.Set_Item($Details.Name,$license_array)
		}
    } 
} 
$convert_table = @()

#Convert hash to table(array)
foreach($hash in $license_table.getenumerator())
{
	$temp_convert = "" | select Name, Capacity, Assigned, Units
	$temp_convert.Name = $hash.Name	
	$temp_convert.Capacity = ($hash.value)[0]
	$temp_convert.Assigned = ($hash.value)[1]
	$temp_convert.Units = ($hash.value)[2]
	$convert_table += $temp_convert
}

#Export entire array to CSV file
$vSphereLicInfo | export-csv \\webservershare\wwwroot\VMLicenseReport.csv

#Generate email message and send
$message_body = $convert_table | convertto-html
send-mailmessage -to $emailaddress -from "" -subject "VMware Licensing" -body ($message_body | out-string) -BodyAsHTML -attachments "\\webservershare\wwwroot\VMLicenseReport.csv" -smtpserver mail.server.com

#Generate html page and post it to webserver
$HTMLHeader = @"
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="javascript/sorttable.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" media="all" href="css/tableorganize.css" />
<p>Report Last Ran: $date</p>
"@
$Report = $convert_table | ConvertTo-Html -Title "ReportTitle" -Head $HTMLHeader -As Table



#Replace HTML Tags for the appropriate CSS
$Report = $Report | ForEach-Object {$_ -replace "<body>", '<body id="body">'}
$Report = $Report | ForEach-Object {$_ -replace "<table>", '<a href="http://webserver.host.com/vmlicensereport.csv">Download report</a><table>'}
$Report = $Report | ForEach-Object {$_ -replace "<table>", '<table class="sortable" id="table" cellspacing="0">'}

#Output to File
$Report = $Report | Set-Content \\webservershare\wwwroot\LicenseReport.html
