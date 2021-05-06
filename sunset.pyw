from datetime import datetime, time
import pandas
import ctypes
from playsound import playsound
from yeelight import Bulb

light_bulb_ip = "192.168.1.7" #IP of the Xiaomi light bulb
mylight = Bulb(light_bulb_ip) #The light bulb is defined here

file_path = "myfile.xlsx" #sunrise/sunset file path
data = pandas.read_excel(file_path, header=0) #Header on line 0

#Today as day number in reference to 1st of Jan
day = datetime.now().timetuple().tm_yday

#Today's parameters
#sr and ss are column names in the Excel spreadsheet for sunrise and sunset respectively
#Minus 1 to account for 0 based indexing
sunrise = data["sr"][day - 1]
sunset = data["ss"][day - 1] 

#Time right now
now = datetime.now().time()

#Function to convert time objects into integers
def seconds_in_time(time_value: time):
    return (time_value.hour * 60 + time_value.minute) * 60 + time_value.second

#Variable for a moment in time 5 minutes before the sunset
sunset_minus_five = seconds_in_time(sunset) - 60 * 5

#Setting up the day_night variable depending on the now variable
#delta calculates the difference in seconds between now and sunset -during night- and sunrise -during day-
#A negative value for delta means that now variable is equal to any moment between midnight and the sunrise  
if now > sunrise and now < sunset:
    day_night = 'day'
    delta = (seconds_in_time(now) - seconds_in_time(sunrise))
else:
    day_night = 'night'
    delta = (seconds_in_time(now) - seconds_in_time(sunset))

#delta_notification calculates the difference in seconds between now and sunset_minus_five
delta_notification = seconds_in_time(now) - sunset_minus_five
    
#The path to the wallpapers being used
path = 'C:\\Users\\mariu\\Desktop\\Sunset\\wallpapers\\'+ day_night +'.jpg'

#Function defined to perform an action (open/close the light) depending on the time of the day
def on_off():
    if now > sunrise and now < sunset:
        return mylight.turn_off()
    else: 
        return mylight.turn_on()

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, path, 3) #SystemParametersInfoW for x64 architecture

#An alarm sound is played and a red light turns on if delta_notification is less than 60 seconds AND delta_notification is greater than 0
if delta_notification < 60 and delta_notification > 0:
    playsound('alarm.wav') #Plays the sound 
    mylight.set_rgb(255, 0, 0) #Set red color for the light bulb
    mylight.turn_on()
      
#Wallpaper changes, a three-beep sound is played, and light turns on only if delta is less than 60 seconds AND delta is greater than 0
#In order for the light to turn on, the script should be ran on a computer that is on the same network as the light bulb (wireless)
if delta < 60 and delta > 0:
    changeBG(path)
    playsound('sound.mp3') #Plays the sound
    mylight.set_rgb(253, 216, 68) #Set yellow color for the light bulb
    on_off()
