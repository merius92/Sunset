from datetime import datetime, time
import pandas
import ctypes
from playsound import playsound
from PythonMETAR import *
from lxml import html
import requests
import re
from pathlib import Path

aerodrome = 'LROD'
source = 'https://flightplan.romatsa.ro/init/meteows/getopmet?ad=' + aerodrome
page = requests.get(source)
tree = html.fromstring(page.content)
raw_report = tree.xpath('//div[@class="report"]/text()')[0]
report = re.sub(' +', ' ', raw_report)[4:-3] #regex eliminates some of the whitespaces, while [4:-3] finishes the job

report_metar = Metar(aerodrome, report)

#Function to convert time objects into integers
def seconds_in_time(time_value: time):
    return (time_value.hour * 60 + time_value.minute) * 60 + time_value.second

properties_time = report_metar.getAttribute('date_time')
(day, hour, minutes) = properties_time
metar_time = int(hour) * 60 * 60 + int(minutes) * 60
time_zone_dif = 3
now = datetime.now().time()
seconds_now = seconds_in_time(now) - time_zone_dif * 60 * 60
seconds_in_a_day = 24 * 60 * 60

if seconds_now > 0:
    delta_metar = seconds_now - metar_time
else:
    delta_metar = seconds_in_a_day - metar_time - seconds_now

#Checking if the METAR is AUTO  
properties_auto = report_metar.getAttribute('auto')

#Function to check if the METAR is AUTO, is valid or is out-of-sync
def time_checker():
    if properties_auto == True:
        return 'AUTO METAR'
    else:    
        if delta_metar < 32 * 60:
            return 'METAR_OK'
        else:
            return 'Check METAR time'    
        
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

#The path the folder
abs_path = Path().resolve()
#The path to the wallpapers being used
target_path = abs_path / 'wallpapers' / day_night / f'{vmc_imc_day_night()}.jpg'
#Converting the path to string
path = str(target_path)

#Function to change the wallpaper
def changeBG(path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, path, 3) #SystemParametersInfoW for x64 architecture

#Wallpaper when code is ran user log on
changeBG(path)
