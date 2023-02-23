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
import icalendar
import configparser
import pathlib
import logging

#python -m venv venv
#Windows: venv\Scripts\activate
#Mac/Linux: source venv/bin/activate
#python -m pip install
#pip install pywin32
#pip install ics
#pip install requests
#pip install icalendar



config_path = pathlib.Path(__file__).parent.absolute() / "config.ini"
config = configparser.ConfigParser()
config.read(config_path)



tasks_folder_name = config['TaskScheduler']['tasks_folder']
calendar_url = config['TaskScheduler']['calendar_url']
task_program_path = config['TaskScheduler']['task_program_path']
task_program_arguments = config['TaskScheduler']['task_program_arguments']


logging.basicConfig(filename='log.txt', encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
# create logger
logger = logging.getLogger('')
logger.setLevel(logging.DEBUG)

# create console handler and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# create formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# add formatter to ch
ch.setFormatter(formatter)

# add ch to logger
logger.addHandler(ch)



# Windows Task Scheduler
scheduler = win32com.client.Dispatch('Schedule.Service')
scheduler.Connect()
root_folder = scheduler.GetFolder('\\')

# TODO Delete all past / expired events
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
                folder.DeleteTask(task.Name, 0)


# TODO Create tasks for future events

# Retrieve the calendar
# url = "https://github.com/beyarkay/eskom-calendar/releases/download/latest/gauteng-ekurhuleni-block-3.ics"
calendar = Calendar(requests.get(calendar_url).text)

# TODO Delete any event that no longer has a calendar entry

def add_task(event):
    # Extract the necessary information from the event
    summary = 'LoadShedding_' + event.begin.strftime('%Y%m%d-%H%M')
    start_time = event.begin - timedelta(minutes=5)
    end_time = event.end
    description = event.description

    print(event.begin - timedelta(minutes=5))

    # Create a new Task Scheduler object
    scheduler = win32com.client.Dispatch("Schedule.Service")

    # Connect to the Task Scheduler
    scheduler.Connect()

    # Create a new Task Folder to store the tasks
    root_folder = scheduler.GetFolder('\\')
    # tasks_folder_name = tasks_folder
    tasks_folder_exists = False
    for folder in root_folder.GetFolders(0):
        if str(folder.Name).lower() == str(tasks_folder_name).lower():
            tasks_folder = folder
            tasks_folder_exists = True
            break
    if not tasks_folder_exists:
        tasks_folder = root_folder.CreateFolder(tasks_folder_name, None)

    # Create a new Task Definition
    task_definition = scheduler.NewTask(0)
    task_definition.RegistrationInfo.Description = f'{event.name} from {event.begin.strftime("%H:%M")} to {event.end.strftime("%H:%M")}'
    task_definition.Settings.Enabled = True
    task_definition.Settings.StartWhenAvailable = True
    task_definition.Settings.Hidden = False

    # Create a new Trigger for the Task
    trigger = task_definition.Triggers.Create(1)
    trigger.StartBoundary = start_time.strftime('%Y-%m-%dT%H:%M:%S')
    trigger.EndBoundary = end_time.strftime('%Y-%m-%dT%H:%M:%S')
    trigger.Enabled = True

    # Create a new Action for the Task
    action = task_definition.Actions.Create(0)
    action.Path = task_program_path
    action.Arguments = task_program_arguments

    # Save the Task
    tasks_folder.RegisterTaskDefinition(summary, task_definition, 6, None, None, 3)


# folders = [scheduler.GetFolder('\\')]
# while folders:
#     folder = folders.pop(0)
#     folders += list(folder.GetFolders(0))
#     for task in folder.GetTasks(0):
#         print('Name       : %s' % task.Name)
#         print('Path       : %s' % task.Path)
#         print('Last Run   : %s' % task.LastRunTime)
#         print('Last Result: %s' % task.LastTaskResult)
#         match = re.search('LoadShedding(_[0-9]{8}-[0-9]{4})', task.Name)
#         if match:
#             taskNamePart = str.split(task.Name, '_')[1]
#             taskDate = datetime.strptime(taskNamePart, '%Y%m%d-%H%M')
#             if taskDate.timestamp() > datetime.now().timestamp():
#                 # get the folder to delete the task from
#                 #task_folder = scheduler.GetFolder(task.Path)
#                 f = scheduler.GetFolder(task.Path)
#                 f.DeleteTask(task.Name, 0)




# Add a scheduled task for each event in the calendar
# for component in calendar.subcomponents:
#     if component.name == 'VEVENT':
#         add_task(component)

# TODO Delete this
# Loop through events to test
for event in calendar.events:
    if event.begin.timestamp() > datetime.now().timestamp():
        add_task(event)
        print(event.description)
        print(event.categories)
        print(event.extra)
        print(event.name)
        print(str(event.begin) + ' to ' + str(event.end))


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
            match = re.search('LoadShedding(_[0-9]{8}-[0-9]{4})', task.Name)
            if match:
                taskNamePart = str.split(task.Name, '_')[1]
                taskDate = datetime.strptime(taskNamePart, '%Y%m%d-%H%M')
                if taskDate.timestamp() < datetime.now().timestamp():
                    folder.DeleteTask(task.Name, 0)





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