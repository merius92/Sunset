from datetime import datetime, time
import pandas
import ctypes
import time

file_path = "table.xlsx" #Excel Spreadsheet file path
data = pandas.read_excel(file_path, header=0) #Header on line 0
day = datetime.now().timetuple().tm_yday #Today as day number in reference to 1st of Jan

#sr and ss are column names in the Excel spreadsheet
#-1 to account for 0 based indexing
sunrise = data["sr"][day-1] #Today's sunrise time
sunset = data["ss"][day-1] #Today's sunset time
sunrise_next = data["sr"][day] #Tomorrow's sunrise time
now = datetime.now().time() #Time right now


#Function to convert time objects into integers
def seconds_in_time(time_value: time):
    return (time_value.hour * 60 + time_value.minute) * 60 + time_value.second

seconds_in_a_day = 86400
delta = (seconds_in_time(sunset) - seconds_in_time(now)) #Variable that calculates the time until the next sunset
delta_next = seconds_in_a_day - (seconds_in_time(now) - seconds_in_time(sunrise_next)) #Variable that calculates the time until the next sunrise

#Function to determine day/night
def period():
    if now > sunrise and now < sunset:
        return('day')
    else:
        return('night')

#Function to call the variables in the while loop
def current_parameters():
    return day
    return sunrise
    return sunset
    return sunrise_next
    return period()
    return delta
    return delta_next

path = 'C:\\wallpapers\\' + period() + '.jpg' #The path to the wallpapers being used

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, path, 3) #SystemParametersInfoW for x64 architecture
    
#The script will be ran just after user login (WIN TASK SCHEDULER)
#This statement will run if the trigger event (user login) is between sunrise and sunset (during the day)
while now > sunrise and now < sunset:
    time.sleep(delta)
    now = datetime.now().time() #Time right now
    current_parameters()
    path = 'C:\\wallpapers\\' + period() + '.jpg'
    changeBG(path)
    time.sleep(delta_next)
    now = datetime.now().time() #Time right now
    current_parameters()
    path = 'C:\\wallpapers\\' + period() + '.jpg'
    changeBG(path)

#This statement will run if the trigger event (user login) is between sunset and sunrise (during the night)
while now < sunrise or now > sunset:
    time.sleep(delta_next)
    current_parameters()
    path = 'C:\\wallpapers\\' + period() + '.jpg'
    changeBG(path)
    time.sleep(delta)
    current_parameters()
    path = 'C:\\wallpapers\\' + period() + '.jpg'
    changeBG(path)



