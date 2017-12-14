#Get Description from AD and add it to the VMware Note Field
import-module ActiveDirectory
$errorActionPreference = "SilentlyContinue"

foreach ($virtMach in (get-vm | where-object {$_.guest.OSFullName -like "*Microsoft*"} | select-object name,notes ))
{
	$machineName = ($virtMach | select-object -expandproperty name)
	$description = (get-adcomputer -filter 'name -like $machineName' -properties description | select-object -expandproperty description)
	set-vm -confirm:$false -vm $machineName -description ""
	set-vm -confirm:$false -vm $machineName -description $description	
	$description = ""
}
