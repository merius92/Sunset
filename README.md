# Sunset
This is a collection of Python scripts designed to change the Windows background image, to play a sound and to open a locally connected light bulb/light band at the start of the night. At the end of the night (the beginning of the day), it reverts the background image, plays a sound and closes the light.

## Applications

The script was designed to be ran in an Air Traffic Control Tower environment, a workplace where it's necessary to know when the night has began/day has started, in order to perform specific actions (open the runway lights during the night, notice VFR traffic that the civil day has ended). 

Instead of checking the published data manually, this collection of scripts checks the data and perform a action in order to make the Air Traffic Controller aware that the day or night has begun.

## Script Description

The **sunset.pyw** script should be scheduled to be ran each minute in the background by Windows Task Scheduler. The script is designed to perform two main actions:
* 5 minutes before sunset, it plays an alarm sound and opens the light (with a red color) in order to alert the user that a period change is about to happen in the near future.
* it evaluates if in the last 60 seconds there has been a period change (ex: day has become night or vice-versa) and runs the actions: changes the wallpaper, plays a beep, and opens/closes the light (depending on the time of the day)

The **start_up.pyw** script should be scheduled to be ran on log on in the background by Windows Task Scheduler. This script runs a specific action depending on the time of the day (changes the background image and opens/closes the light).

## Prerequisites
* Windows 10 computer
* Microsoft Excel spreadsheet with the dates in a year in column A (starting 1st of Jan, going to 31st of Dec), and sunset + sunrise info in columns B and C (HH:MM:SS format) - sample data in **myfile.xlsx**.
* 2 tasks added in **Task Scheduler**: one for **sunset.pyw**, one for **start_up.pyw**. Models can be found in tasks folder.

For light opening/closing functionalities:
* **Xiaomi Yeelight** gear
* A computer that is connected wirelessly to the same network as the Xiaomi Yeelight gear (if not possible, a Xiaomi Gateway is necessary)
