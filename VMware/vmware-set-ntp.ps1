$newNTPServer = "ntp.host.com"

foreach ($esxhost in get-vmhost)
{
	write-host "Changing NTP on Server: $esxhost";
	Get-VmHostService -VMHost $esxhost | Where-Object {$_.key -eq "ntpd"} | Start-VMHostService;
	Get-VmHostService -VMHost $esxhost | Where-Object {$_.key -eq "ntpd"} | Set-VMHostService -policy "automatic";

}