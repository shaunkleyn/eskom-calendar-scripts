#$tasks = Get-ScheduledTask -TaskName "LOAD*"
$logfile = "$env:computername.log"

$tasks = Get-ScheduledTask | Where-Object taskname -match 'load'
$today = Get-Date

foreach ($task in $tasks) {
    $taskDatePart = $task.TaskName.Split("_")[1]
    if ($null -ne $taskDatePart) {
        $taskDate=[Datetime]::ParseExact($taskDatePart, 'yyyyMMdd-HHmm', $null)
        if ($taskDate -lt $today) {
            Write-Output "Old date"
            WriteLog "The script is run"
        }
    }
}

function WriteLog {
    Param ([string]$LogString)
    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $LogMessage = "$Stamp $LogString"
    Add-content $logfile -value $LogMessage
}