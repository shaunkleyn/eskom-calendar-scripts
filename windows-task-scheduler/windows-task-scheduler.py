
import win32com.client
# imports
from datetime import datetime, timedelta
from pathlib import Path
import os
import pytz
from ics import Calendar, Event
import requests
import re
import icalendar
import configparser
import pathlib
import logging
import pytz
from dateutil.tz import gettz

#python -m venv venv
#Windows: venv\Scripts\activate
#Mac/Linux: source venv/bin/activate
#python -m pip install
#pip install pywin32
#pip install ics
#pip install requests
#pip install icalendar

log_path = os.path.join(str(pathlib.Path(__file__).parent.absolute()), str(os.path.basename(__file__).replace('.py', '.log')))
logging.basicConfig(filename=log_path, encoding='utf-8', level=logging.DEBUG, format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')
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



#config
config_path = pathlib.Path(__file__).parent.absolute() / "config.ini"
logger.info(f'using config: "{str(config_path)}"')
config = configparser.ConfigParser()
config.read(config_path)


tasks_folder_name = config['TaskScheduler']['tasks_folder']
calendar_url = config['TaskScheduler']['calendar_url']
task_program_path = config['TaskScheduler']['task_program_path']
task_program_arguments = config['TaskScheduler']['task_program_arguments']



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
        logger.info('Name       : %s' % task.Name)
        print('Path       : %s' % task.Path)
        print('Last Run   : %s' % task.LastRunTime)
        print('Last Result: %s' % task.LastTaskResult)
        match = re.search('LoadShedding(_[0-9]{8}-[0-9]{4})', task.Name)
        if match:
            taskNamePart = str.split(task.Name, '_')[1]
            taskDate = datetime.strptime(taskNamePart, '%Y%m%d-%H%M')
            if taskDate.timestamp() < datetime.now().timestamp():
                logger.info(f'Removing old event: {task.Name} from folder "{folder.Name}"')
                folder.DeleteTask(task.Name, 0)


# TODO Create tasks for future events

# Retrieve the calendar
# url = "https://github.com/beyarkay/eskom-calendar/releases/download/latest/gauteng-ekurhuleni-block-3.ics"
calendar = Calendar(requests.get(calendar_url).text)

# TODO Delete any event that no longer has a calendar entry


def remove_emojis(data):
    emoj = re.compile("["
        u"\U0001F600-\U0001F64F"  # emoticons
        u"\U0001F300-\U0001F5FF"  # symbols & pictographs
        u"\U0001F680-\U0001F6FF"  # transport & map symbols
        u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
        u"\U00002500-\U00002BEF"  # chinese char
        u"\U00002702-\U000027B0"
        u"\U00002702-\U000027B0"
        u"\U000024C2-\U0001F251"
        u"\U0001f926-\U0001f937"
        u"\U00010000-\U0010ffff"
        u"\u2640-\u2642" 
        u"\u2600-\u2B55"
        u"\u200d"
        u"\u23cf"
        u"\u23e9"
        u"\u231a"
        u"\ufe0f"  # dingbats
        u"\u3030"
                      "]+", re.UNICODE)
    return re.sub(emoj, '', data)


def add_task(event, tzinfo):
    summary = 'LoadShedding_' + event.begin.strftime('%Y%m%d-%H%M')
    start_time = event.begin.datetime.astimezone(local_tz) - timedelta(minutes=5)
    end_time = event.end.datetime.astimezone(local_tz)
    clean_event_name = remove_emojis(event.name)

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
    task_definition.RegistrationInfo.Description = f'{clean_event_name} from {start_time.strftime("%H:%M")} to {end_time.strftime("%H:%M")}'
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
    action.Arguments = f'"{clean_event_name} from {start_time.strftime("%H:%M")} to {end_time.strftime("%H:%M")}"'
    if task_program_arguments != '':
        action.Arguments = task_program_arguments.replace('_eventname_', clean_event_name).replace('_eventstart_', start_time.strftime("%H:%M")).replace('_eventend_', end_time.strftime("%H:%M"))

    # Save the Task
    tasks_folder.RegisterTaskDefinition(summary, task_definition, 6, None, None, 3)

now = datetime.now()
local_now = now.astimezone()
local_tz = local_now.tzinfo

# TODO Delete this
# Loop through events to test
for event in calendar.events:
    if event.end.datetime.astimezone(local_tz).timestamp() > datetime.now().astimezone(local_tz).timestamp():
        add_task(event, local_tz)

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
