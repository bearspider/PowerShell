 Get-VM | Get-View | Select Name, `
@{N="CpuHotAddEnabled";E={$_.Config.CpuHotAddEnabled}}, `
@{N="CpuHotRemoveEnabled";E={$_.Config.CpuHotRemoveEnabled}}, `
@{N="MemoryHotAddEnabled";E={$_.Config.MemoryHotAddEnabled}} | Export-Csv f:\scripts\hotplug.csv
