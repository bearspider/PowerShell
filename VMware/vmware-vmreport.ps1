param(
	[array]$vcenters,
	[string]$emailaddress,
	[switch]$affinityreport
	);
$date = get-date
# Adds the base cmdlets
if((Get-PSSnapin -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue) -eq $null)
{
	Add-PSSnapin VMware.VimAutomation.Core
}
#Input: hostname of ESXi server, (True/False)is the ESXi server in a cluster 
#Output: An array of VM Summaries
#Description: This function queries a VMhost for each VM running on it, the amount of physical CPU sockets and cores
#             It compiles the information into a table showing the following:
#			  Virtual Server | Cluster it belongs to | ESXi Host Virtual Server is running on | Physical CPU Sockets | Cores per CPU
function Get-hostReport ($vmhost, $isInCluster, $numHostsCluster, $onvcenter)
{
	$returnArray = @()
	#write-host "Host: $vmhost"
	#Get ESXi host information
	$hostInfo = get-view $vmHost
	#Number of Physical CPUs in the ESXi Host
	$numCPUs = $vmHost.ExtensionData.Summary.Hardware.NumCpuPkgs
	#Number of cores per Physical CPU
	$numCores = $vmHost.ExtensionData.Summary.Hardware.NumCpuCores/$numCPUs
	#Number of CPU Threads
	$numThreads = $vmHost.ExtensionData.Summary.Hardware.NumCpuThreads
	
	ForEach ($VM in ($vmHost | get-vm | where{$_.PowerState -eq "PoweredOn"} ))
	{
		#write-host "VM: $VM"
		#Get Virtual Server information
		$VMView = $VM | Get-View
		#Create table with the below column headings
		$VMSummary = "" | Select Server, OS, ClusterName, vCenterName, HostName, HostsInCluster, Memory, CPUSockets, CPUCores, CPUThreads, vCPUSockets, vCPUCores, Storage, Description
		#Assign number of hosts in cluster to HostsInCluster
		$VMSummary.HostsInCluster = $numHostsCluster
		#Get VM name, assign it to Server column
		$VMSummary.Server = $VMView.Name
		#Assign vCenterName that the guest VM is located in
		$VMSummary.vCenterName = $onvcenter
		#Get OS Type, assign it to OS Column
		if((get-vmguest $VM).OSFullName -ne $null)
		{
			$VMSummary.OS = (get-vmguest $VM).OSFullName
		}
		#Get ESXi hostname, assign it to HostName column
		$VMSummary.HostName = $hostInfo.Name
		#Determine if the ESXi host is in a cluster, if it is assign ClusterName column the Cluster Name
		#if the ESXi host is not in a cluster, assign the ClusterName as "No Cluster"
		if($isInCluster)
		{
			$VMSummary.ClusterName = $Cluster.Name
		}
		else
		{
			$VMSummary.ClusterName = "No Cluster"
		}
		#Assign the memory to the Memory column
		$VMSummary.Memory = [math]::Ceiling($VM.MemoryGB)
		#Assign the number of physical CPU sockets to CPUSockets column
		$VMSummary.CPUSockets = $numCPUs
		#Assign the number of cores per CPU to CPUCores column
		$VMSummary.CPUCores = $numCores
		#Assign the number of CPU Threads to threads column
		$VMSummary.CPUThreads = $numThreads
		#Assign the number of cores per vCPU to vCPUCores column
		$VMSummary.vCPUCores = $VMView.config.hardware.NumCoresPerSocket
		#Assign the number of virtual CPU sockets to vCPUSockets column
		if($VMSummary.vCPUCores -lt 1)
		{
			"vCPUCores less than 0: $VMSummary.Server" | add-content c:\temp\vmreportErrors.txt
		}
		else
		{
			$VMSummary.vCPUSockets = $VMView.config.hardware.NumCPU/$VMSummary.vCPUCores
		}
		#Assign the Description to Description Column
		$VMSummary.Description = $VM | select-object -expandproperty notes
		#Assign the Storage info to Storage Column
		$VMSummary.Storage = [math]::Ceiling($VM.ProvisionedSpaceGB)
		#Push the table into the array
		$returnArray += $VMSummary		
	}
	#Return the array of tables to the script
	return,$returnArray
}

#############################################################################
#Start of Main Program
#############################################################################
#$colItems = (get-datacenter).count
#$count = 1
$myCol = @()
#Query for datacenters
#File location of affinity report
$file_location = "\\webserver\wwwroot\affinityreport.txt"
if($affinityreport)
{
	write-host "Deleting old assessment report"
	if(test-path $file_location)
	{
		remove-item -Confirm:$False $file_location
	}
	add-content -path $file_location $date
}
disconnect-viserver -erroraction silentlycontinue -server * -Force -Confirm:$False
Set-PowerCliConfiguration  -Confirm:$false -DefaultVIServerMode Multiple
ForEach ($vcenter in $vcenters)
{
    connect-viserver $vcenter
	ForEach ($datacenter in Get-Datacenter)
	{
		#write-host "Compiling data for Datacenter: $datacenter"
		#write-progress -Activity "Probing Datacenter" -status $datacenter -percentComplete (($count / $colItems)*100)
		#If there are any clusters in the datacenter, then separate each step out
		if($datacenter | get-cluster)
		{
			#For each cluster in the datacenter, get each VMHost
			ForEach ($Cluster in ($datacenter | Get-Cluster))
			{
				if($affinityreport)
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
				#For each VMHost in the cluster, get report for that host and push it into Array
				$clusterCount = ($Cluster | get-vmhost | measure-object | select-object -expandproperty count)
				ForEach ($vmHost in ($cluster | get-vmhost))
				{
					$mycol += Get-hostReport -vmhost $vmhost -isincluster $True -numHostsCluster $clusterCount -onvcenter $vcenter
				}
			}
			#For each folder that has hosts not in a cluster, get the folders
			ForEach ($folder in (($datacenter | get-folder -type "HostAndCluster" | select-object -expandproperty Name) -ne "host"))
			{
				#For each folder, grab the VMHost in the folder, get the VMHost report and put it into Array
				if([string]$folder -ne "False")
				{
					ForEach ($vmHost in (get-folder $folder | get-vmhost))
					{
						$mycol += Get-hostReport -vmhost $vmhost -isincluster $True -numHostsCluster "0" -onvcenter $vcenter
					}
				}
			}
		}
		#Else if no clusters exist, just grab all the vmhosts in the datacenter
		else
		{
			#Get information for each VMHost in the datacenter
			ForEach ($vmHost in ($datacenter | get-vmhost))
			{
				$mycol += Get-hostReport -vmhost $vmhost -isincluster $False -numHostsCluster "0" -onvcenter $vcenter
			}
		}
		$count += 1 
	}
}
#Export entire array to CSV file
$myCol | export-csv \\webserver\wwwroot\VMStorageReport.csv
send-mailmessage -to $emailaddress -from "VMReports@mailserver.com" -subject "VMs" -attachments "\\webserver\wwwroot\VMStorageReport.csv",$file_location -smtpserver mail.server.com

$HTMLHeader = @"
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="javascript/sorttable.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" media="all" href="css/tableorganize.css" />
<p>Report Last Ran: $date</p>
"@

$Report = $myCol | ConvertTo-Html -Title "ReportTitle" -Head $HTMLHeader -As Table

#Replace HTML Tags for the appropriate CSS
$Report = $Report | ForEach-Object {$_ -replace "<body>", '<body id="body">'}
$Report = $Report | ForEach-Object {$_ -replace "<table>", '<a href="http://webserver.host.com/vmstoragereport.csv">Download report</a><table>'}
$Report = $Report | ForEach-Object {$_ -replace "<table>", '<table class="sortable" id="table" cellspacing="0">'}

#Output to File
$Report = $Report | Set-Content \\webserver\wwwroot\Report.html

