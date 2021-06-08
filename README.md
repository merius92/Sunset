# Intro
This is a collection of Python scripts designed to change the Windows background image, and to play a sound at the start of the night. At the end of the night (the beginning of the day), it reverts the background image, plays a sound and closes the light.
Another functionality of these scripts is that the background image changes during the day whenever the METAR conditions are non-VMC during the day, or if there is an issue with the METAR (AUTO, out-of-sync METAR etc.)

## Applications

The script was designed to be ran in an Air Traffic Control Tower environment, a workplace where it's necessary to know when the night has began/day has started, in order to perform specific actions (open the runway lights during the night, notice VFR traffic that the civil day has ended). 

Instead of checking the published data manually, this collection of scripts checks the data and perform a action in order to make the Air Traffic Controller aware that the day or night has begun.

## Script Description

The **sunset.pyw** script should be scheduled to be ran each minute in the background by Windows Task Scheduler. The script is designed to perform two main actions:
* 5 minutes before sunset, it plays an alarm sound in order to alert the user that a period change is about to happen in the near future.
* It evaluates if in the last 60 seconds there has been a period change (ex: day has become night or vice-versa) and runs the actions: changes the wallpaper, plays a beep (depending on the time of the day).

The **start_up.pyw** script should be scheduled to be ran on log on in the background by Windows Task Scheduler. This script runs a specific action depending on the time of the day (changes the background image and opens).

## Prerequisites
* Windows 10 operating system
* Microsoft Excel spreadsheet with the dates in a year in column A (starting 1st of Jan, going to 31st of Dec), and sunset + sunrise info in columns B and C (HH:MM:SS format) - sample data in **myfile.xlsx**
* 2 tasks added in **Task Scheduler**: one for **sunset.pyw**, one for **start_up.pyw**. Models can be found in tasks folder
* Pip requirements file added via pip install
* Time of the computer in local time (but the script can also be adapted to work in UTC, if necessary)
