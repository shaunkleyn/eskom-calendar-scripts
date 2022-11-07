import datetime
import win32com.client
# imports
from datetime import datetime, timedelta
from pathlib import Path
import os
import pytz
from ics import Calendar
import requests
import re

# Retrieve the calendar
url = "https://github.com/beyarkay/eskom-calendar/releases/download/latest/gauteng-ekurhuleni-block-3.ics"
calendar = Calendar(requests.get(url).text)

# TODO Delete this
# Loop through events to test
for event in calendar.events:
    if event.begin.timestamp() > datetime.now().timestamp():
        print(str(event.begin) + ' to ' + str(event.end))


# TODO Delete all past / expired events

# TODO Create tasks for future events

# TODO Delete any event that no longer has a calendar entry


# Windows Task Scheduler
scheduler = win32com.client.Dispatch('Schedule.Service')
scheduler.Connect()

folders = [scheduler.GetFolder('\\')]
while folders:
    folder = folders.pop(0)
    folders += list(folder.GetFolders(0))
    for task in folder.GetTasks(0):
        print('Name       : %s' % task.Name)
        print('Path       : %s' % task.Path)
        print('Last Run   : %s' % task.LastRunTime)
        print('Last Result: %s' % task.LastTaskResult)
        match = re.search('LoadShedding(_[0-9]{8}-[0-9]{4})', task.Name)
        if match:
            taskNamePart = str.split(task.Name, '_')[1]
            taskDate = datetime.strptime(taskNamePart, '%Y%m%d-%H%M')
            if taskDate.timestamp() < datetime.now().timestamp():
                # get the folder to delete the task from
                #task_folder = scheduler.GetFolder(task.Path)
                f = scheduler.GetFolder(task.Path)
                f.DeleteTask(task.Name, 0)



def getOldTasks():
    #scheduler.Connect()
    folders = [scheduler.GetFolder('\\')]
    while folders:
        folder = folders.pop(0)
        folders += list(folder.GetFolders(0))
        for task in folder.GetTasks(0):
            print('Name       : %s' % task.Name)
            print('Path       : %s' % task.Path)
            print('Last Run   : %s' % task.LastRunTime)
            print('Last Result: %s' % task.LastTaskResult)
            



########### EXAMPLE CODE ##########
task_def = scheduler.NewTask(0)

# Defining the Start time of job
start_time = datetime.now() + datetime.date.timedelta(minutes=1)

# For Daily Trigger set this variable to 2 ; for One time run set this value as 1
TASK_TRIGGER_DAILY = 1
trigger = task_def.Triggers.Create(TASK_TRIGGER_DAILY)

#Repeat for a duration of number of day
num_of_days = 10
trigger.Repetition.Duration = "P"+str(num_of_days)+"D"

#use PT2M for every 2 minutes, use PT1H for every 1 hour
trigger.Repetition.Interval = "PT2M"
trigger.StartBoundary = start_time.isoformat()

# Create action
TASK_ACTION_EXEC = 0
action = task_def.Actions.Create(TASK_ACTION_EXEC)
action.ID = 'TRIGGER BATCH'
action.Path = 'cmd.exe'
action.Arguments ='/c start "" "C:\\Ajay\\Desktop\\test.bat"'

# Set parameters
task_def.RegistrationInfo.Description = 'Test Task'
task_def.Settings.Enabled = True
task_def.Settings.StopIfGoingOnBatteries = False

# Register task
# If task already exists, it will be updated
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_NONE = 0
root_folder.RegisterTaskDefinition(
    'Test Task',  # Task name
    task_def,
    TASK_CREATE_OR_UPDATE,
    '',  # No user
    '',  # No password
    TASK_LOGON_NONE
)