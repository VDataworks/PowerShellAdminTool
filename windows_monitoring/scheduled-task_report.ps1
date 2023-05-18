# Confirm ExecutionPolicy change
Set-ExecutionPolicy -Force Bypass
Set-ExecutionPolicy -Force Bypass

# Clear prompt
Clear-Host
# Clear cached variables and modules
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear();

# Import required modules
if (!(Get-Module -Name ActiveDirectory)) {
    Write-Host "ActiveDirectory module not found. Installing module..."
    Install-Module -Name ActiveDirectory -Force
}
if (!(Get-Module -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing module..."
    Install-Module -Name ImportExcel -Force
}

## VARIABLES
$ADFilter = "*"
$HostEnable = "True"
$ADSearchBase = "OU=,DC=,DC=com" ### THIS NEED TO BE MODIFY BEFORE RUN
$TaskNameToSearch = "Stop PC" ### THIS NEED TO BE MODIFY BEFORE RUN

# Get all computers from AD
$Computers = Get-ADComputer -Filter {Name -like $ADFilter -and Enabled -eq $HostEnable} -SearchBase $ADSearchBase -Properties Name | Sort-Object

# Script progress
$TotalCounter = ($Computers | Measure-Object).Count
$Counter = 1

# Test connection to each computer and retrieve scheduled task details
$Results = foreach ($Computer in $Computers) {
    Write-Progress "Checking Scheduled Task $($TaskNameToSearch) on : ..." -Status $Computer.Name -Id 1 -PercentComplete (($Counter / $TotalCounter) * 100)
    if (Test-Connection -ComputerName $Computer.Name -Count 1 -Quiet) {
        Write-Host "Connection successful"
        $task = Get-ScheduledTask -TaskName $TaskNameToSearch -CimSession $Computer.Name -ErrorAction SilentlyContinue
        if ($task) {
            [PSCustomObject]@{
                ComputerName = $Computer.Name
                TaskName = $task.TaskName
                State = $task.State
                LastRunTime = $task.LastRunTime
                NextRunTime = $task.NextRunTime
            }
        }
    } else {
        Write-Host "Connection failed"
    }
    $Counter++
}

# Export results to Excel file
$Results | Export-Excel -Path "C:\path\to\file.xlsx" -AutoSize -AutoFilter