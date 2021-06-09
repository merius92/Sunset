from datetime import datetime, time
import pandas
import ctypes
from playsound import playsound
from blinkstick import blinkstick
from PythonMETAR import *
from lxml import html
import requests
import re

bstick = blinkstick.find_first() #Assigns a variable name to the blinkstick

#Metar

#Define light colors

def bstick_turn_red():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x, name="red")
        x+=1

def bstick_turn_purple():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x, name="purple")
        x+=1

def bstick_turn_blue():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x, name="blue")
        x+=1
        
def bstick_turn_yellow():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x, name="yellow")
        x+=1

#Closes all the 32 LEDs
def bstick_turn_off():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x)
        x+=1
        
def red():
    return bstick_turn_red()

def purple():
    return bstick_turn_purple()

def blue():
    return bstick_turn_blue()

def yellow():
    return bstick_turn_yellow()

def off():
    return bstick_turn_off()

#Test metar
#('LFQN','METAR LFQN 201630Z 18005KT 6000 -SHRA SCT030 BKN050 18/12 Q1014')
#'METAR COR LRSV 091630Z 13008KT CAVOK 16/02 Q1022='

aerodrome = 'LROD'
source = 'https://flightplan.romatsa.ro/init/meteows/getopmet?ad=' + aerodrome
page = requests.get(source)
tree = html.fromstring(page.content)
raw_report = tree.xpath('//div[@class="report"]/text()')[0]
report = re.sub(' +', ' ', raw_report)[4:-3] #regex eliminates some of the whitespaces, while [4:-3] finishes the job
print(report)

report_metar = Metar(aerodrome, report)

def seconds_in_time(time_value: time):
    return (time_value.hour * 60 + time_value.minute) * 60 + time_value.second

#Checks if there is a time difference greater than 32 minutes between NOW and the time METAR was latest submitted
#It basically checks if the site is updated accordingly
#time_zone_dif is the time (in hours) difference between the computer's time zone and UTC
def time_checker():
    properties_time = report_metar.getAttribute('date_time')
    (day, hour, minutes) = properties_time
    metar_time = int(hour)*60*60 + int(minutes)*60
    time_zone_dif = 3
    now = datetime.now().time()
    seconds_now = seconds_in_time(now) - time_zone_dif*60*60
    seconds_in_a_day = 24*60*60
    if seconds_now > 0:
        delta = seconds_now - metar_time
    else:
        delta = seconds_in_a_day - metar_time - seconds_now
    properties_auto = report_metar.getAttribute('auto')
    if properties_auto == True:
        return 'AUTO METAR'
    else:    
        if delta < 32*60:
            return 'METAR_OK'
        else:
            return 'Check METAR time'

#Testing material
#report_metar1 = Metar('LFQN','METAR LFQN 201630Z 18005KT 0500 -SHRA 18/12 Q1014')
#report_metar2 = Metar('LFQN','METAR LFQN 201630Z 18005KT 1000 -SHRA SCT005 BKN100 18/12 Q1014')
#report_metar3 = Metar('LFQN','METAR LFQN 201630Z 18005KT 2000 -SHRA SCT005 BKN100 18/12 Q1014')
#report_metar4 = Metar('LFQN','METAR LFQN 201630Z 18005KT 6000 -SHRA SCT005 BKN100 18/12 Q1014')
#report_metar5 = Metar('LFQN','METAR LFQN 201630Z 18005KT 1300 -SHRA SCT005 BKN003 18/12 Q1014')
#report_metar6 = Metar('LFQN','METAR LFQN 201630Z 18005KT 4000 -SHRA SCT005 BKN003 18/12 Q1014')
#report_metar7 = Metar('LFQN','METAR LFQN 201630Z 18005KT 6000 -SHRA 18/12 Q1014')
#report_metar8 = Metar('LRSV','METAR COR LRSV 091630Z 13008KT CAVOK 16/02 Q1022=')
#report_metar9 = METAR LRBM 110830Z 19003KT 160V220 CAVOK 23/11 Q1016=

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
        return 'ERROR'

#Day/Night

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

#Opens all the 32 LEDs on orange color (for 5 minutes pre-event notification)
def bstick_notification():
    x=0
    while x < 32:
        bstick.set_color(channel=0, index=x, name="lime")
        x+=1

file_path = "myfile.xlsx" #sunrise/sunset file path

#Current year
year =  datetime.today().year

if(year % 4 != 0):
    data = pandas.read_excel(file_path, sheet_name="nonleap_year", header=0) #Header on line 0
else:
    data = pandas.read_excel(file_path, sheet_name="leap_year", header=0)

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

#delta_notification calculates the difference in seconds between now and sunset_minus_five
delta_notification = seconds_in_time(now) - sunset_minus_five
    
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

#Function defined to perform turn on the lights on orange color for 5 minutes notification
def notification_on():
    return bstick_notification()

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, path, 3) #SystemParametersInfoW for x64 architecture

#An alarm sound is played and a red light turns on if delta_notification is less or equal than 15 seconds AND delta_notification is greater than -30 delta_notification <= 15 and delta_notification => -30:
if delta_notification < 60 and delta_notification > -1:
    playsound('alarm.wav') #Plays the sound

if delta_notification < 60 and delta_notification > -1 and vmc_imc() == 'VMC':
    notification_on() #Turns on the orange lights

#Wallpaper changes, a three-beep sound is played, and light turns on only if delta is less than 60 seconds AND delta is greater than -1
#In order for the light to turn on, the script should be ran on a computer that is on the same network as the light bulb (wireless)
if delta < 60 and delta > -1:
    changeBG(path) #Wallpaper change
    playsound('sound.mp3') #Plays the sound

if delta < 60 and delta > -1 and vmc_imc() == 'VMC':
    on_off() #Turns on/off the lights

if vmc_imc() == 'Total IMC':
    red()
elif vmc_imc() == 'Special VFR for heli only':
    purple()
elif vmc_imc() == 'Special VFR':
    blue()
elif vmc_imc() == 'Check METAR time':
    yellow()
elif vmc_imc() == 'ERROR':
    yellow()
elif vmc_imc() == 'AUTO METAR':
    yellow()
elif vmc_imc() == 'Check conditions':
    yellow()
elif vmc_imc() == 'VMC' and delta_notification > 0 and delta_notification < notification_seconds - 1:
    notification_on()
else:
    on_off()