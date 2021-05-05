#This script is supposed to run on Windows 10 user log on

from datetime import datetime, time
import pandas
import ctypes
from yeelight import Bulb


file_path = "myfile.xlsx"
data = pandas.read_excel(file_path, header=0) #Header on line 0

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
    delta = (seconds_in_time(now)-seconds_in_time(sunrise))
else:
    day_night = 'night'
    delta = (seconds_in_time(now)-seconds_in_time(sunset))
    
#The path to the wallpapers being used
path = 'C:\\Users\\mariu\\Desktop\\Sunset\\wallpapers\\'+ day_night +'.jpg'
mylight = Bulb("192.168.1.7") #IP of the Xiaomi lightbulb 

#Function defined to perform an action (open/close the light) depending on the time of the day
def on_off():
    if now > sunrise and now < sunset:
        return mylight.turn_off()
    else: 
        return mylight.turn_on()

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, path, 3) #SystemParametersInfoW for x64 architecture

#Wallpaper when code is ran user log on
changeBG(path)
on_off()




