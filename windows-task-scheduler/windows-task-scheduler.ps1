#$tasks = Get-ScheduledTask -TaskName "LOAD*"
# $logfile = "$env:computername.log"

$today = Get-Date
$tz = Get-TimeZone
# Get Eskom calendar

$data = @{}

# Create Table
    $tbl  = New-Object System.Data.DataTable "Events"
    $col1 = New-Object System.Data.DataColumn "Subject"
    $col2 = New-Object System.Data.DataColumn "Body"
    $col3 = New-Object System.Data.DataColumn "Start"
    $col4 = New-Object System.Data.DataColumn "End"
    $col5 = New-Object System.Data.DataColumn "Location"
    $tbl.Columns.Add($col1)
    $tbl.Columns.Add($col2)
    $tbl.Columns.Add($col3)
    $tbl.Columns.Add($col4)
    $tbl.Columns.Add($col5)

# Get / Create a folder for the Load Shedding tasks
# https://devblogs.microsoft.com/scripting/use-powershell-to-create-scheduled-tasks-folders/
Function New-ScheduledTaskFolder {
    Param ($taskpath)
    $ErrorActionPreference = "stop"
    $scheduleObject = New-Object -ComObject schedule.service
    $scheduleObject.connect()
    $rootFolder = $scheduleObject.GetFolder("\")
    Try {$null = $scheduleObject.GetFolder($taskpath)}
    Catch { $null = $rootFolder.CreateFolder($taskpath) }
    Finally { $ErrorActionPreference = "continue" } 
}

Function Create-AndRegisterApplogTask {
    Param ($taskname, $taskpath)
    $action = New-ScheduledTaskAction -Execute 'Powershell.exe' `
        -Argument '-NoProfile -WindowStyle Hidden -command "& {get-eventlog -logname Application -After ((get-date).AddDays(-1)) | Export-Csv -Path c:\fso\applog.csv -Force -NoTypeInformation}"'
    $trigger =  New-ScheduledTaskTrigger -Daily -At 9am
    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName `
    $taskname -Description "Daily dump of Applog" -TaskPath $taskpath
}

Function Create-NewApplotTaskSettings {
    Param ($taskname, $taskpath)
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -Hidden -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -RestartCount 3
    Set-ScheduledTask -TaskName $taskname -Settings $settings -TaskPath $taskpath
}

Function Remove-OldScheduledTasks {
    Param ($taskpath)
    try {       
        $tasks = Get-ScheduledTask -TaskPath $taskpath #| Where-Object taskname -match 'loads'
    
        foreach ($task in $tasks) {
            $taskDatePart = $task.TaskName.Split("_")[1]
            if ($null -ne $taskDatePart) {
                $taskDate=[Datetime]::ParseExact($taskDatePart, 'yyyyMMdd-HHmm', $null)
                if($null -ne $taskDate) {
                    if ($taskDate -lt $today) {
                        try { Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false }
                        catch { Write-Output "Unable to delete task $task.TaskName.  Please delete it manually." }
                    }
                }
            }
        }
    }
    catch { $ErrorActionPreference = "continue" }
    Finally { $ErrorActionPreference = "continue" } 
}


Function Get-Calendar {
    Param ($calendarUrl)

    $csvresponse = ([System.Text.Encoding]::ASCII.GetString((invoke-webrequest "https://github.com/beyarkay/eskom-calendar/releases/download/latest/machine_friendly.csv" | select-object -property Content).Content))
    $csvcontent = $csvresponse | ConvertFrom-Csv -Delim ',' |
    ForEach-Object {
        # $Name += $_.start
        # $Phone += $_.finsh

        if ($_.area_name -eq "gauteng-ekurhuleni-block-3") {

            # Add to Table
            $row = $tbl.NewRow()
            $row.Subject = ""
            $row.Body = ""
            $row.Start = "$($_.start)"
            $row.End = "$($_.finsh)"
            $row.Location = "$($_.area_name)"
            $tbl.Rows.Add($row)
    
            # Clear variables
            $Start = ""
            $End = ""
            $Title = ""  
            $Description = ""
            $Location = ""
            $EndDate = ""
            $EndTime = ""         
            $StartDate = ""
            $StartTime = ""
        }
    }
    #####################
    # ICS to Data Table #
    #####################
    # $response = (invoke-webrequest "https://github.com/beyarkay/eskom-calendar/releases/download/latest/gauteng-ekurhuleni-block-3.ics" | Select-Object -Property Content).Content
    $response = (invoke-webrequest "$calendarUrl" | Select-Object -Property Content).Content
    $content = ([System.Text.Encoding]::ASCII.GetString($response))

    # Process ICS into Powershell table
    ForEach ($line in $($content -split "`r`n"))
    {
        # Split key:value
        if($line.Contains(':')){
            
            $z = @{ $line.split( ':')[0] =( $line.split( ':')[1]).Trim() }
            # Begin
            if ($z.keys -eq "BEGIN"){         
            
            }

            # Get start date
            if ($z.keys -eq "DTSTART;VALUE=DATE") {
                $Start = $z.values -replace "r\n\s"
                $Start = [datetime]::ParseExact($Start,"yyyyMMdd" ,$null)
                $StartDate = $Start.ToLocalTime().ToShortDateString()           
                $StartDate = get-date $StartDate -Format yyyy-MM-dd
                $StartTime = "00:00:00"           
            }

            if ($z.keys -eq "DTSTART") {
                $Start = $z.values -replace "r\n\s"           
                $Start = $Start -replace "T"           
                $Start = $Start -replace "Z"           
                $Start = [datetime]::ParseExact($Start,"yyyyMMddHHmmss" ,$null) 
                $ remoteTime= $Start.AddHours( - ($tz.BaseUtcOffset.totalhours))          
                $StartDate = $Start.ToShortDateString()           
                # $StartTime = $Start.ToLongTimeString()
                $StartTime = $remoteTime.ToLocalTime()
                $StartDate = get-date $StartDate -Format yyyy-MM-dd
            }

            # Get end date
            if ($z.keys -eq "DTEND") {
                $End = $z.values -replace "\r\n\s"           
                $End = $End -replace "T"           
                $End = $End -replace "Z"           
                $End = [datetime]::ParseExact($End,"yyyyMMddHHmmss" ,$null)           
                $EndDate = $End.ToShortDateString()           
                $EndTime = $End.ToLongTimeString()                      
                $EndDate = get-date $EndDate -Format yyyy-MM-dd                      
            }

            if ($z.keys -eq "DTEND;VALUE=DATE") {
                $End = $z.values -replace "r\n\s"
                $End = [datetime]::ParseExact($End,"yyyyMMdd" ,$null)
                $EndDate = $End.ToShortDateString()
                $EndDate = get-date $EndDate -Format yyyy-MM-dd
                $EndTime = "00:00:00"      
            }
        
            # Get summary
            if ($z.keys -eq "SUMMARY") {           
                $Title = $z.values -replace "\r\n\s"           
                $Title = $z.values -replace ",","-"        
            }

            # Get description
            if ($z.keys -eq "DESCRIPTION") {
                $Description = $z.values -replace "\r\n\s"           
                $Description = $Description -replace "<p>"
                $Description = $Description -replace "</p>"        
                $Description = $Description -replace "<div>&nbsp;</div>"  
                $Description = $Description -replace "<div>"
                $Description = $Description -replace "</div>"
            }

            # Get location
            if ($z.keys -eq "LOCATION") {           
                $Location = $z.values -replace "\r\n\s"           
                $Location = $z.values -replace ",","-"           
            }    

            # End of event
            if ($z.keys -eq "END") {

                # Check Subject exists
                if($Title -ne ""){
                
                    # Add to Table
                    $row = $tbl.NewRow()
                    $row.Subject = "$Title"
                    $row.Body = "$Description"
                    $row.Start = "$($StartDate)T$($StartTime)"
                    $row.End = "$($EndDate)T$($EndTime)"
                    $row.Location = "$Location"
                    $tbl.Rows.Add($row)
                }

                # Clear variables
                $Start = ""
                $End = ""
                $Title = ""  
                $Description = ""
                $Location = ""
                $EndDate = ""
                $EndTime = ""         
                $StartDate = ""
                $StartTime = ""
            }
        }
    }
}

Function Register-FutureScheduledTasks {
    Param ($taskpath)
    #Loop through each row of data and create a new file
    #The dataset contains a column named FileName that I am using for the name of the file
    foreach($row in $tbl.Rows) { 
        if ($row.Start -gt $today) {
            $action = New-ScheduledTaskAction -Execute "cmd.exe"
            $trigger = New-ScheduledTaskTrigger -Once -At $row.Start
            # $principal = "Contoso\Administrator"
            $settings = New-ScheduledTaskSettingsSet
            # $task = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger -Settings $settings
            $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings
            $dateStr = Get-Date $row.Start -Format "yyyyMMdd-HHmm"
            Register-ScheduledTask "LoadShedding_${dateStr}" -InputObject $task -TaskPath $taskpath 
        }
    }
}


### ENTRY POINT ###
$taskname = "LoadShedding_yyyyMMdd-HHmm"
$taskpath = "LoadShedding"
$calendarUrl = "https://github.com/beyarkay/eskom-calendar/releases/download/latest/gauteng-ekurhuleni-block-3.ics"

Get-Calendar -calendarUrl $calendarUrl
New-ScheduledTaskFolder -taskname $taskname -taskpath $taskpath
Remove-OldScheduledTasks -taskpath "\$taskpath\"

Register-FutureScheduledTasks -taskpath "\$taskpath\"

If(Get-ScheduledTask -TaskName $taskname -EA 0) {
    Unregister-ScheduledTask -TaskName $taskname -Confirm:$false
}

Create-AndRegisterApplogTask -taskname $taskname -taskpath $taskpath | Out-Null
Create-NewApplotTaskSettings -taskname $taskname -taskpath $taskpath | Out-Null






# Function Remove-OldScheduledTasks {
#     Param ($taskpath)
#     $tasks = Get-ScheduledTask -TaskPath $taskpath #| Where-Object taskname -match 'loads'

#     foreach ($task in $tasks) {
#         $taskDatePart = $task.TaskName.Split("_")[1]
#         if ($null -ne $taskDatePart) {
#             $taskDate=[Datetime]::ParseExact($taskDatePart, 'yyyyMMdd-HHmm', $null)
#             if($null -ne $taskDate) {
#                 if ($taskDate -lt $today) {
#                     try { Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false }
#                     catch { Write-Output "Could not delete task $task.TaskName.  Please delete it manually." }
#                 }
#             }
#         }
#     }
    

#     foreach($row in $tbl.Rows) { 
#         if ($row.Start -gt $today) {
#             $action = New-ScheduledTaskAction -Execute "cmd.exe"
#             $trigger = New-ScheduledTaskTrigger -OneTime
#             # $principal = "Contoso\Administrator"
#             $settings = New-ScheduledTaskSettingsSet
#             # $task = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger -Settings $settings
#             $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings
#             Register-ScheduledTask T1 -InputObject $task
#         }
#     }
# }








# $content = $file
# $content
# $file
# $data = @{}
# # $content | foreach-Object {
# ForEach ($line in $($file -split "`r`n")) {
#     # $line = $_
#     try {
#         Write-Output $line
#         $key = ($line.split( ':').split( ';')[0])
#         $value = ( $line.split( ':')[1]).Trim()
#         if ($key) {
#             $data[$key] = $value
#         }
#     }
#     catch {
#     }   
# }

# $Body = [regex]::match($content, '(?<=\DESCRIPTION:).+(?=\DTEND:)', "singleline").value.trim()
#     $Body = $Body -replace "\r\n\s"
#     $Body = $Body.replace("\,", ",").replace("\n", " ")
#     $Body = $Body -replace "\s\s"

#     $Start = ($data.getEnumerator() | ? { $_.Name -eq "DTSTART"}).Value -replace "T"
#     $Start = [datetime]::ParseExact($Start , "yyyyMMddHHmmss" , $null )

#     $End = ($data.getEnumerator() | ? { $_.Name -eq "DTEND"}).Value -replace "T"
#     $End = [datetime]::ParseExact($End , "yyyyMMddHHmmss" , $null )

#     $Subject = ($data.getEnumerator() | ? { $_.Name -eq "SUMMARY"}).Value
#     $Location = ($data.getEnumerator() | ? { $_.Name -eq "LOCATION"}).Value
#     $eventObject = [PSCustomObject]@{
#         Body     = $Body
#         Start    = $Start
#         End      = $End
#         Subject  = $Subject
#         Location = $Location
#     }


$data = Invoke-WebRequest -Uri "https://github.com/beyarkay/eskom-calendar/releases/download/latest/gauteng-ekurhuleni-block-3.ics" -Method get
# Write-Output $data.content
[string]$d = get-content $data.content -Raw
Write-Output $d

$tasks = Get-ScheduledTask | Where-Object taskname -match 'loads'


foreach ($task in $tasks) {
    $taskDatePart = $task.TaskName.Split("_")[1]
    if ($null -ne $taskDatePart) {
        $taskDate=[Datetime]::ParseExact($taskDatePart, 'yyyyMMdd-HHmm', $null)
        if ($taskDate -lt $today) {
            
            Write-Output "Old date"
            try {

                Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false
            }
            catch {
                Write-Output "Could not delete task $task.TaskName.  Please delete it manually."
                # $ts = New-Object -ComObject "Schedule.Service"
                # $ts.Connect()
                # $t = $ts.GetFolder("").GetTask($task.TaskName)
                # Write-Host $t.GetSecurityDescriptor(0xF)
                # ConvertFrom-SddlString -Sddl ($t.GetSecurityDescriptor(0xF)) -Type RegistryRights
            }
        }
    }
}


# function WriteLog {
#     Param ([string]$LogString)
#     $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
#     $LogMessage = "$Stamp $LogString"
#     Add-content $logfile -value $LogMessage
# }