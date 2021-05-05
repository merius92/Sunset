# Sunset
This is a phython changes background image, plays a sound and opens the light at the start of the night. At the end of the night and the beginning of the day, it reverts the background image, plays a sound and closes the light.

The sunset.pyw script should be scheduled to be ran each minute in the background by Windows Task Scheduler. This script evaluates if in the last 60 seconds there has been a period change (ex: day has become night or vice-versa) and runs the actions.

The start_up.pyw script should be scheduled to be ran on log on in the background by Windows Task Scheduler. This script runs a specific action depending on the time of the day (changes the background image and opens/closes the light).
