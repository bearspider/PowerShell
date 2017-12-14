Class Storsimple
{
    [ValidateNotNullOrEmpty()][string]$DeviceName
    [ValidateNotNullOrEmpty()][string]$SerialNumber
    [ValidateNotNullOrEmpty()][string]$Function
    [ValidateNotNullOrEmpty()][string]$MaxCapacity
    [ValidateNotNullOrEmpty()][string]$AvailableLocalCapacity
    [ValidateNotNullOrEmpty()][string]$AvailableTieredCapacity
    [ValidateNotNullOrEmpty()][string]$LocalCapacityRemaining
    [ValidateNotNullOrEmpty()][string]$TieredCapacityRemaining
    [ValidateNotNullOrEmpty()][string]$ProvisionedStorage
    [ValidateNotNullOrEmpty()][string]$UsedStorage
    [ValidateNotNullOrEmpty()][string]$PercentProvisioned
    [ValidateNotNullOrEmpty()][string]$PercentUsed
    Print()
    {
        write-host "Device: $($this.DeviceName)"
        write-host "Serial Number: $($this.SerialNumber)"
        write-host "Function: $($this.Function)"
        write-host "Max Capacity: $($this.MaxCapacity)"
        write-host "Available Local Capacity: $($this.AvailableLocalCapacity)"
        write-host "Available Tiered Capacity: $($this.AvailableTieredCapacity)"
        write-host "Local Remaining Capacity: $($this.LocalCapacityRemaining)"
        write-host "Tiered Remaining Capacity: $($this.TieredCapacityRemaining)"
        write-host "Provisioned Storage: $($this.ProvisionedStorage)"
        write-host "Used Storage: $($this.UsedStorage)"
        write-host "Percent Provisioned: $($this.PercentProvisioned)"
        write-host "Percent Used: $($this.PercentUsed)`r`n"
    }
    [string]Output()
    {
        $output = "Device: $($this.DeviceName)`r`nSerial Number: $($this.SerialNumber)`r`nFunction: $($this.Function)`r`nMax Capacity: $($this.MaxCapacity)`r`nAvailable Local Capacity: $($this.AvailableLocalCapacity)`r`nAvailable Tiered Capacity: $($this.AvailableTieredCapacity)`r`nRemaining Local Capacity: $($this.LocalCapacityRemaining)`r`nRemaining Tiered Capacity: $($this.TieredCapacityRemaining)`r`nProvisioned Storage: $($this.ProvisionedStorage)`r`nUsed Storage: $($this.UsedStorage)`r`nPercent Provisioned: $($this.PercentProvisioned)`r`nPercent Used: $($this.PercentUsed)`r`n`r`n"

        return $output
    }
    [array]ToArray()
    {
        $output = @($this.DeviceName, $this.SerialNumber,$this.Function,$this.MaxCapacity,$this.AvailableLocalCapacity,$this.AvailableTieredCapacity,$this.LocalCapacityRemaining,$this.TieredCapacityRemaining,$this.ProvisionedStorage,$this.UsedStorage,$this.PercentProvisioned,$this.PercentUsed)

        return $output
    }
    [String]ToJSON()
    {
        $rval = @{
            $this.DeviceName = @{
                "Serial Number" = $this.SerialNumber;
                "Function" = $this.Function;
                "Max Capacity" = $this.MaxCapacity;
                "Available Local Capacity" = $this.AvailableLocalCapacity;
                "Available Tiered Capacity" =  $this.AvailableTieredCapacity;
                "Remaining Local Capacity" = $this.LocalCapacityRemaining;
                "Remaining Tiered Capacity" = $this.TieredCapacityRemaining;
                "Provisioned Storage" = $this.ProvisionedStorage;
                "Used Storage" = $this.UsedStorage;
                "Percent Provisioned" = $this.PercentProvisioned;
                "Percent Used" = $this.PercentUsed;
            }
        }
        return ($rval | ConvertTo-JSON -depth 8);
    }
}