write-host "Current Power Settings"
Get-VMHost | Sort | Select Name, @{ N="CurrentPolicy"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.ShortName}}, @{ N="CurrentPolicyKey"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.Key}}, @{ N="AvailablePolicies"; E={$_.ExtensionData.config.PowerSystemCapability.AvailablePolicy.ShortName}}

write-host "Changing Power Settings"
$views  = (Get-VMHost | Get-View)
foreach ($view in $views)
{
	(Get-View $view.ConfigManager.PowerSystem).ConfigurePowerPolicy(1)
}
write-host "New Power Settings"
Get-VMHost | Sort | Select Name, @{ N="CurrentPolicy"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.ShortName}}, @{ N="CurrentPolicyKey"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.Key}}, @{ N="AvailablePolicies"; E={$_.ExtensionData.config.PowerSystemCapability.AvailablePolicy.ShortName}}


