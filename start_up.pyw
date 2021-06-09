#This script is supposed to run on Windows 10 user log on

from datetime import datetime, time
import pandas
import ctypes
from blinkstick import blinkstick
from pathlib import Path


bstick = blinkstick.find_first() #Assigns a variable name to the blinkstick


#Opens all the 32 LEDs on white color
def bstick_turn_on():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x, name="white")
        x+=1

#Closes all the 32 LEDs
def bstick_turn_off():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x)
        x+=1

file_path = "myfile.xlsx"

#Current year
year =  datetime.today().year

if(year % 4 != 0):
    data = pandas.read_excel(file_path, sheet_name="nonleap_year", header=0) #Header on line 0
else:
    data = pandas.read_excel(file_path, sheet_name="leap_year", header=0)

#Today as day number in reference to 1st of Jan
day = datetime.now().timetuple().tm_yday

#Today's parameters
#sr and ss are column names in the Excel spreadsheet
#Minus 1 to account for 0 based indexing
sunrise = data["sr"][day-1]
sunset = data["ss"][day-1] 

#Time right now
now = datetime.now().time()

#Function to convert time objects into integers
def seconds_in_time(time_value: time):
    return (time_value.hour * 60 + time_value.minute) * 60 + time_value.second

#Setting up the day_night variable depending on the now variable
#delta calculates the difference in seconds between now and sunset -during night- and sunrise -during day-
if now > sunrise and now < sunset:
    day_night = 'day'
    delta = (seconds_in_time(now) - seconds_in_time(sunrise))
else:
    day_night = 'night'
    delta = (seconds_in_time(now) - seconds_in_time(sunset))
    
#The path to the folder
abs_path = Path().resolve()
#The path to the wallpapers being used
target_path = abs_path / 'wallpapers' / f'{day_night}.jpg'
#Converting the path to string
path = str(target_path)

#Function defined to perform an action (open/close the light) depending on the time of the day
def on_off():
    if now > sunrise and now < sunset:
        return bstick_turn_off()
    else: 
        return bstick_turn_on()

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, path, 3) #SystemParametersInfoW for x64 architecture

#Wallpaper when code is ran user log on
changeBG(path)
on_off()




