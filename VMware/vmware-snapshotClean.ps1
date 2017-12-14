param(
	[string]$vm
	)
[array]$vcenters = ""

if((Get-PSSnapin -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) -eq $null)
{
	Add-PSSnapin VMware.VimAutomation.Core
}
Set-PowerCliConfiguration -Confirm:$false -DefaultVIServerMode Multiple
disconnect-viserver -erroraction silentlycontinue -server * -Force -Confirm:$False
ForEach ($vcenter in $vcenters)
{
	    connect-viserver $vcenter
}
[datetime]$today = get-date -uformat %D
[datetime]$date = (get-date).AddDays(-7).ToString("MM/dd/yy")

$snapshots = get-snapshot -vm $vm
foreach($snap in $snapshots.description)
{
	write-host "Snapshot: $snap"
	$span = (new-timespan -start $date -end ([datetime]$snap)).Days > 7
	write-host "Span: $span"
	if ((new-timespan -start $date -end ([datetime]$snap)).Days > 7)
	{
		write-host "Deleting $snap"
		#remove-snapshot $snap -Confirm:$false
	}
}