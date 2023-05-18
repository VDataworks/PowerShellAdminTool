Install-Module PSHyperVTools
Install-Module ImportExcel
Import-Module PSHyperVTools
Import-Module ImportExcel

Function Get-VMResourceUsage {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$VMName,
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )
 
    $query = "SELECT PercentProcessorTime, PercentMemoryUsed FROM Msvm_ComputerSystem WHERE ElementName = '$VMName'"
 
    Get-WmiObject -Query $query -ComputerName $ComputerName -Namespace root\virtualization\v2 | Select-Object -ExpandProperty Properties | Select-Object -Property Name, Value
}

# Define the array to store the VM information
$vmInfo = @()

# Loop through each Hyper-V server
$serverList = "localhost"
$vmList = @()
foreach ($server in $serverList) {
    $vmMetrics = Get-CimInstance -ComputerName $server -ClassName "Msvm_MetricForVirtualMachine"
    $vms = Get-VM -ComputerName $server
    foreach ($vm in $vms) {
        $vmDisk = Get-VHD -Path $vm.VhdPath | Select-Object Name, Path, VhdType, ParentPath, VhdFormat, ComputerName
        $settings = $vm | Get-VMIntegrationService
        $ipAddress = ($vm | Get-VMNetworkAdapter).IPv4Addresses.IPAddress
        $macAddress = ($vm | Get-VMNetworkAdapter).MACAddress
        $resourceUsage = Get-VMResourcePool -VMName $vm.Name -ComputerName $server | Select-Object -Property Name, CPUUsage, MemoryAssigned, MemoryUsage
        $vmInfo = [PSCustomObject]@{
            ServerName      = $server
            Name            = $vm.Name
            State           = $vm.State
            DynamicMemory   = $vm.DynamicMemoryEnabled
            Memory          = $vm.MemoryAssigned/1MB
            ProcessorCount  = $vm.ProcessorCount
            Uptime          = $vm.Uptime
            DiskName        = $vmDisk.Name
            DiskPath        = $vmDisk.Path
            DiskType        = $vmDisk.VhdType
            ParentDiskPath  = $vmDisk.ParentPath
            DiskFormat      = $vmDisk.VhdFormat
            DiskComputer    = $vmDisk.ComputerName
            MetricsAverageCpu = ($vmMetrics | Where-Object { $_.ElementName -eq $vm.Name -and $_.MetricDefinitionId -eq "Microsoft:Hyper-V:Virtual Processor\CPU Usage (%)"}).MetricValue
            MetricsMemoryDemand = ($vmMetrics | Where-Object { $_.ElementName -eq $vm.Name -and $_.MetricDefinitionId -eq "Microsoft:Hyper-V:Virtual Memory\Memory Demand (KB)"}).MetricValue/1KB
            MetricsMemoryAssigned = ($vmMetrics | Where-Object { $_.ElementName -eq $vm.Name -and $_.MetricDefinitionId -eq "Microsoft:Hyper-V:Virtual Memory\Memory Assigned (KB)"}).MetricValue/1KB
            MetricsNetworkRx = ($vmMetrics | Where-Object { $_.ElementName -eq $vm.Name -and $_.MetricDefinitionId -eq "Microsoft:Hyper-V:Virtual Network Adapter(*)\Bytes Received/sec"}).MetricValue
            MetricsNetworkTx = ($vmMetrics | Where-Object { $_.ElementName -eq $vm.Name -and $_.MetricDefinitionId -eq "Microsoft:Hyper-V:Virtual Network Adapter(*)\Bytes Sent/sec"}).MetricValue
            Settings = $settings
            IpAddress = $ipAddress
            MACAddress = $macAddress
            ResourceUsage = $resourceUsage
        }
        $vmList += $vmInfo
    }
}

# Export the VM information to a formatted Excel file
$outputFile = "C:\temp\vmInfo.xlsx"
$vmList | Export-Excel -Path $outputFile -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
