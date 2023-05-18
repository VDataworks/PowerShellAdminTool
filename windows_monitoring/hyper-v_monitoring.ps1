Install-Module PSHyperVTools

Import-Module PSHyperVTools

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

# Create Excel object
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Create Excel headers
$worksheet.Cells.Item(1,1) = "Hyper-V Host"
$worksheet.Cells.Item(1,2) = "VM Name"
$worksheet.Cells.Item(1,3) = "VM Memory (GB)"
$worksheet.Cells.Item(1,4) = "VM vCPUs"
$worksheet.Cells.Item(1,5) = "VM Status"
$worksheet.Cells.Item(1,6) = "VM Uptime"
$worksheet.Cells.Item(1,7) = "VM Host"
$worksheet.Cells.Item(1,8) = "VM Integration Services Version"
$worksheet.Cells.Item(1,9) = "VM IP Address"
$worksheet.Cells.Item(1,10) = "VM MAC Address"
$worksheet.Cells.Item(1,11) = "VM Resource Usage CPU (%)"
$worksheet.Cells.Item(1,12) = "VM Resource Usage Memory (%)"

$row = 2

$servers = "localhost"  # replace with your Hyper-V server names

# loop through each server and get VM info
foreach ($server in $servers) {
    Write-Host "Getting VMs from $server"
    $vms = Get-VM -ComputerName $server

    # loop through each VM and get settings and resource usage
    foreach ($vm in $vms) {
        $settings = $vm | Get-VMIntegrationService
        $ipAddress = ($vm | Get-VMNetworkAdapter).IPv4Addresses.IPAddress
        $macAddress = ($vm | Get-VMNetworkAdapter).MACAddress
        $resourceUsage = Get-VMResourcePool -VMName $vm.Name -ComputerName $server | Select-Object -Property Name, CPUUsage, MemoryAssigned, MemoryUsage

        $worksheet.Cells.Item($row,1) = $vm.Name
        $worksheet.Cells.Item($row,2) = $vm.Name
        $worksheet.Cells.Item($row,3) = $vm.MemoryAssigned/1GB
        $worksheet.Cells.Item($row,4) = $vm.ProcessorCount
        $worksheet.Cells.Item($row,5) = $vm.State
        $worksheet.Cells.Item($row,6) = $vm.Uptime
        $worksheet.Cells.Item($row,7) = $host
        $worksheet.Cells.Item($row,8) = $settings.version
        $worksheet.Cells.Item($row,9) = $ipAddress
        $worksheet.Cells.Item($row,10) = $macAddress
        $worksheet.Cells.Item($row,11) = $resourceUsage.CPUUsage
        $worksheet.Cells.Item($row,12) = $resourceUsage.MemoryUsage/$vm.MemoryAssigned*100

        $row++
    }
}

# Auto-fit columns and save Excel file
$worksheet.Columns.AutoFit() | Out-Null
$workbook.SaveAs("C:\temp\Test1_VMReport.xlsx")
$excel.Quit()