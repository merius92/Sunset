from datetime import datetime, time
import pandas
import ctypes
from playsound import playsound
from PythonMETAR import *
from lxml import html
import requests
import re

#Test metar
#('LFQN','METAR LFQN 201630Z 18005KT 6000 -SHRA SCT030 BKN050 18/12 Q1014')
#'METAR COR LRSV 091630Z 13008KT CAVOK 16/02 Q1022='

aerodrome = 'LROP'
source = 'https://flightplan.romatsa.ro/init/meteows/getopmet?ad=' + aerodrome
page = requests.get(source)
tree = html.fromstring(page.content)
raw_report = tree.xpath('//div[@class="report"]/text()')[0]
report = re.sub(' +', ' ', raw_report)[4:-3] #regex eliminates some of the whitespaces, while [4:-3] finishes the job
print(report)

report_metar = Metar(aerodrome, report)

#Checks if there is a time difference greater than 32 minutes between NOW and the time METAR was latest submitted
#It basically checks if the site is updated accordingly
#time_zone_dif is the time (in hours) difference between the computer's time zone and UTC

#Function to convert time objects into integers
def seconds_in_time(time_value: time):
    return (time_value.hour * 60 + time_value.minute) * 60 + time_value.second

#Conversion of METAR time in seconds
properties_time = report_metar.getAttribute('date_time')
(day, hour, minutes) = properties_time
metar_time = int(hour) * 60 * 60 + int(minutes) * 60

#Time zone difference
time_zone_dif = 3

#Actual moment in seconds
now = datetime.now().time()
seconds_now = seconds_in_time(now) - time_zone_dif * 60 * 60
seconds_in_a_day = 24 * 60 * 60

if seconds_now > 0:
    delta_metar = seconds_now - metar_time
else:
    delta_metar = seconds_in_a_day - metar_time - seconds_now
    
properties_auto = report_metar.getAttribute('auto')

#Checks if METAR report is outdated (last issue time was less/more than 32 minutes ago)
def time_checker():
    if properties_auto == True:
        return 'AUTO METAR'
    else:    
        if delta_metar < 32 * 60:
            return 'METAR_OK'
        else:
            return 'Check METAR time'

print(delta_metar)        
        
properties_cld = report_metar.getAttribute('cloud')
properties_vis = report_metar.getAttribute('visibility')
properties_wind = report_metar.getAttribute('wind')

#Modify the visibility function in order to bypass the library in case there is a CAVOK and WIND VARIATION [None] and return 9999 regardless of the conditions
def adjusted_vis():
    if properties_vis is not None:
        return properties_vis
    else:
        return 9999

#Check if there are IMC cloud conditions; if there is at least 1 True value in the list, then there are non-VMC cloud conditions
#The ELSE statement works for the situation in which there is no cloud group reported in the METAR/SPECI and automatically returns a list containing False
def cld_chck_vmc_minima():
    if properties_cld is not None:
        return [i['code'] in ('BKN', 'OVC') and i['altitude'] < 1500 for i in properties_cld]
    else: 
        return [False]

#Check if there are IMC cloud conditions; if there is at least 1 True value in the list, then there are total IMC cloud conditions
#The ELSE statement works for the situation in which there is no cloud group reported in the METAR/SPECI and automatically returns a list containing False    
def cld_chck_special_vfr_minima():
    if properties_cld is not None:
        return [i['code'] in ('BKN', 'OVC') and i['altitude'] < 600 for i in properties_cld]
    else: 
        return [False]

#Defining VMC and IMC according to https://www.easa.europa.eu/sites/default/files/dfu/Easy%20Access%20Rules%20for%20Standardised%20European%20Rules%20of%20the%20Air%20%28SERA%29.pdf  
def vmc_imc():
    if time_checker() == 'METAR_OK':     
        if adjusted_vis() > 5000 and True not in cld_chck_vmc_minima():
            return 'VMC'
        elif True in cld_chck_special_vfr_minima() or adjusted_vis() < 800:
            return 'Total IMC'
        elif True in cld_chck_special_vfr_minima() or adjusted_vis() < 1500:
            return 'Special VFR for heli only'
        elif True in cld_chck_vmc_minima() or adjusted_vis() < 5000:
            return 'Special VFR'
        else:
            return 'Check conditions'
    elif time_checker() == 'AUTO METAR':
        return 'AUTO'
    elif time_checker() == 'Check METAR':
        return 'Check METAR'
    else:
        return 'METAR ERROR'

#Day/Night

file_path = "myfile.xlsx" #sunrise/sunset file path
data = pandas.read_excel(file_path, header=0) #Header on line 0

#Today as day number in reference to 1st of Jan
day = datetime.now().timetuple().tm_yday

#Today's parameters
#sr and ss are column names in the Excel spreadsheet for sunrise and sunset respectively
#Minus 1 to account for 0 based indexing
sunrise = data["sr"][day - 1]
sunset = data["ss"][day - 1] 

#Function to convert time objects into integers
def seconds_in_time(time_value: time):
    return (time_value.hour * 60 + time_value.minute) * 60 + time_value.second

notification_minutes = 5
notification_seconds = notification_minutes * 60
#Variable for a moment in time 5 minutes before the sunset
sunset_minus_five = seconds_in_time(sunset) - notification_seconds

#Setting up the day_night variable depending on the now variable
#delta calculates the difference in seconds between now and sunset -during night- and sunrise -during day-
#A negative value for delta means that now variable is equal to any moment between midnight and the sunrise  
if now > sunrise and now < sunset:
    day_night = 'day'
    delta = (seconds_in_time(now) - seconds_in_time(sunrise))
else:
    day_night = 'night'
    delta = (seconds_in_time(now) - seconds_in_time(sunset))
    
def vmc_imc_day_night():
    if day_night == 'night' and vmc_imc() == 'METAR OK':
        return 'night'
    else:
        return vmc_imc()

#delta_notification calculates the difference in seconds between now and sunset_minus_five
delta_notification = seconds_in_time(now) - sunset_minus_five

#The path to the wallpapers being used
path = 'C:\\Users\\mariu\\Desktop\\Sunset\\wallpapers_desktop_only\\'+ day_night +'\\'+ vmc_imc_day_night() +'.jpg'

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, path, 3) #SystemParametersInfoW for x64 architecture

#An alarm sound is played and a red light turns on if delta_notification is less or equal than 15 seconds AND delta_notification is greater than -30 delta_notification <= 15 and delta_notification => -30:
if delta_notification < 60 and delta_notification > -1:
    playsound('alarm.wav') #Plays the sound

#Wallpaper changes, a three-beep sound is played, and light turns on only if delta is less than 60 seconds AND delta is greater than -1
#In order for the light to turn on, the script should be ran on a computer that is on the same network as the light bulb (wireless)
if (delta < 60 and delta > -1) or ((delta_metar < 60 and delta_metar > -1) and vmc_imc != "METAR OK"):
    changeBG(path) #Wallpaper change
    playsound('sound.mp3') #Plays the sound
