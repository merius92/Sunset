from datetime import datetime, time
import pandas
import ctypes

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
path = 'C:\\wallpapers\\'+ day_night +'.jpg'
SPI_SETDESKWALLPAPER = 20

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, path, 3) #SystemParametersInfoW for x64 architecture

#Wallpaper changes only if delta is less than 60 seconds. 
if delta < 60:
    changeBG(path)





