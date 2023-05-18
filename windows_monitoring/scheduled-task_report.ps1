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
$GetResults = @()
$SetResults = @()
$GetErrors = @()
$SetErrors = @()
$OutputFileName = ""

# Get all computers from AD
$Computers = Get-ADComputer -Filter {Name -like $ADFilter -and Enabled -eq $HostEnable} -SearchBase $ADSearchBase -Properties Name | Sort-Object

# Script progress
$TotalCounter = ($Computers | Measure-Object).Count
$Counter = 1

# Test connection to each computer
foreach ($Computer in $Computers) {
    Write-Progress "Checking Scheduled Task $($TaskNameToSearch) on : ..." -Status $Computer.Name -Id 1 -PercentComplete (($Counter / $TotalCounter) * 100)

    # Get all details of a specific scheduled task
    try {
        if (Test-Connection -ComputerName $Computer.Name -Count 1 -Quiet) {
            $task = Get-ScheduledTask -TaskName $TaskNameToSearch -CimSession $Computer.Name -ErrorAction SilentlyContinue
            if ($task) {
                $GetResults += [PSCustomObject]@{
                    ComputerName = $Computer.Name
                    TaskName = $task.TaskName
                    State = $task.State
                    LastRunTime = $task.LastRunTime
                    NextRunTime = $task.NextRunTime
                    Settings = $task.Settings
                    Trigger = $task.Triggers
                    Actions = $task.Actions
                    TimeTrigger = $task.Triggers.CalendarTrigger.StartBoundary
                    ScheduleType = "ScheduleByDay"
                    DaysInterval = $task.Triggers.CalendarTrigger.ScheduleByDay.DaysInterval
                }
            }
        }
    }
    catch {
        $GetErrors += [PSCustomObject]@{
            ComputerName = $Computer.Name
            ErrorEncountered = $error[0].Exception.Message
        }
    }

    # Edit the trigger of the task
    try {
        if (Test-Connection -ComputerName $Computer.Name -Count 1 -Quiet) {
            $task = Get-ScheduledTask -TaskName $TaskNameToSearch -CimSession $Computer.Name -ErrorAction SilentlyContinue
            if ($task) {
                $SetResults += [PSCustomObject]@{
                    ComputerName = $Computer.Name
                    TaskName = $task.TaskName
                    OldTimeTrigger = $task.Triggers.CalendarTrigger.StartBoundary
                    OldScheduleType = "ScheduleByDay"
                    OldDaysInterval = $task.Triggers.CalendarTrigger.ScheduleByDay.DaysInterval
                }

                $trigger = New-ScheduledTaskTrigger -Daily -At "22:00"
                Set-ScheduledTaskTrigger -Task $task -Trigger $trigger -CimSession $computer.Name

                $SetResults | Add-Member -MemberType NoteProperty -Name "NewTimeTrigger" -Value $task.Triggers.CalendarTrigger.StartBoundary
                $SetResults | Add-Member -MemberType NoteProperty -Name "NewScheduleType" -Value "ScheduleByDay"
                $SetResults | Add-Member -MemberType NoteProperty -Name "NewDaysInterval" -Value $task.Triggers.CalendarTrigger.ScheduleByDay.DaysInterval
            }
        }
    }
    catch {
        $SetErrors += [PSCustomObject]@{
            ComputerName = $Computer.Name
            ErrorEncountered = $error[0].Exception.Message
        }
    }

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
$GetResults | Export-Excel -WorksheetName "TaskDetails" -Path "C:\$($OutputFileName).xlsx" -AutoSize -AutoFilter -BoldTopRow
$GetErrors | Export-Excel -WorksheetName "ConnectionErrors" -Path "C:\$($OutputFileName).xlsx" -AutoSize -AutoFilter -BoldTopRow
$SetResults | Export-Excel -WorksheetName "TaskUpdate" -Path "C:\$($OutputFileName).xlsx" -AutoSize -AutoFilter -BoldTopRow
$SetErrors | Export-Excel -WorksheetName "EditTaskErrors" -Path "C:\$($OutputFileName).xlsx" -AutoSize -AutoFilter -BoldTopRow
