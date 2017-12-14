ForEach ($vmhost in (get-vmhost))
{
	$view = Get-View $vmhost
	(Get-View $view.ConfigManager.PowerSystem).ConfigurePowerPolicy(1)
}