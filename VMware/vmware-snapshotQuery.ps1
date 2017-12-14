[array]$vcenters = "vcenter"
$emailaddress = "recipient@mailserver.com"

if((Get-PSSnapin -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) -eq $null)
{
	Add-PSSnapin VMware.VimAutomation.Core
}
disconnect-viserver -erroraction silentlycontinue -server * -Force -Confirm:$False
Set-PowerCliConfiguration -Confirm:$false -DefaultVIServerMode Multiple
ForEach ($vcenter in $vcenters)
{
	connect-viserver $vcenter
}
$returnArray = @()
ForEach ($snapshot in (get-vm | get-snapshot))
{
	$snapshotSummary = "" | select virtualMachine,name,description,sizegb
	$snapshotSummary.virtualMachine = $snapshot.vm
	$snapshotSummary.name = $snapshot.name
	$snapshotSummary.description = $snapshot.description
	$snapshotSummary.sizegb = "{0:N2}" -f $snapshot.sizegb
	$returnArray += $snapshotSummary
}


$HTMLHeader = @"
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="javascript/sorttable.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" media="all" href="css/tableorganize.css" />
"@



$Report = $returnArray | ConvertTo-Html -Title "ReportTitle" -Head $HTMLHeader -As Table

#Replace HTML Tags for the appropriate CSS
$Report = $Report | ForEach-Object {$_ -replace "<body>", '<body id="body">'}
$Report = $Report | ForEach-Object {$_ -replace "<table>", '<table class="sortable" id="table" cellspacing="0">'}
send-mailmessage -to $emailaddress -from "vSphereMaintenance@mailserver.com" -subject "Active Snapshots" -bodyashtml -body "$Report" -smtpserver "mail.server.com"
