param(
	[array]$vcenters
	);
$file_location = "c:\temp\affinityreport.txt"
write-host "Deleting old assessment report"
if(test-path $file_location)
{
	remove-item -Confirm:$False $file_location
}
foreach($vcenter in $vcenters)
{
	disconnect-viserver * -Force -Confirm:$False -erroraction SilentlyContinue
	write-host "Connecting to vCenter Server: "$vcenter
	connect-viserver $vcenter
	foreach($cluster in get-cluster)
	{
		
		$groups = ($cluster.extensiondata.configurationex.group | where-object { $_.name -like "sql*" })
		if($groups)
		{
			add-content -path $file_location "Cluster: $cluster"
			foreach($group in $groups )
			{
				add-content -path $file_location "->$($group.name)"
				if($group.vm)
				{
					foreach($rule in $group.vm)
					{
						$hostvm = get-vm -id $rule
						add-content -path $file_location "-->$hostvm"
					}
				}
				if($group.host)
				{
					foreach($rule in $group.host )
					{
						$host_esxi = get-vmhost -id $rule
						add-content -path $file_location "-->$($host_esxi.name)"
					}
				}
			}
			add-content -path $file_location "`n"
		}
	}
}
write-host "Assessment Complete"