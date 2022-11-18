# Google Calendar

## Overview
Being a data hoarder I hate losing historical data as it might just come in handy at some point in time.  So I created this script to copy load-shedding events from the machine_friendly.csv file into my Google calendar.  By doing so, I will have all historical load-shedding events for my area even when the csv file is updated.

## Requirements
You would have to enable the Google Calendar API in a Google Cloud project to be able to use this script.  See this for more info on how to do it: https://developers.google.com/calendar/api/quickstart/python

## What does this script do?
- Inserts future load-shedding events that are in the csv file but not in your calendar.
- Deletes future events that are in your calendar but not in the csv file. Sometimes (ughm *seldom*) Eskom would reduce the load-shedding stage.  This will result in some future schedules being changed / removed.  

## How to use it
- Clone the repo or [download](https://github.com/shaunkleyn/eskom-calendar-scripts/archive/refs/heads/main.zip) it as a zip package and extract it to wherever you want
- Open the destination folder and type ```cmd``` in the address bar to open the command prompt
- Create a virtual environment by running ```python -m venv gc-venv``` (this keeps the script's packages and its runtime environment totally separate from your host machine’s environment)
- Activate the virtual environment using ```.\gc-venv\Scripts\activate```
- Install the required packages by running ```python -m pip install -r requirements.txt```
- Create a scheduled task in Windows to run the script at a preferred interval so that it can update your Google calendar.  See this for more info on how to run a Python script using Windows Task Scheduler: https://www.jcchouinard.com/python-automation-using-task-scheduler/