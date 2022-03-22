import decodeclimate as dc
import os 
from datetime import date, timedelta


data_path = "/home/rd/data/AM4/METAR/"
dir_list = os.listdir(data_path)
 

list_airport = open(os.path.dirname(os.path.realpath(__file__))+'/dir_airport_climate', "r")
list_airport = list_airport.read().splitlines()
mon = '03'
year = '2022'

for airport in list_airport:
    dir_airport = os.path.join(data_path, airport)
    dir_work = os.path.join(dir_airport, year)
    os.chdir(dir_work)
    a = dc.decode(airport)
    a.filerw()
    a.model_a()
    a.model_b()
    a.model_c()
    a.model_d()
    a.model_e()
    a.model_f()
    a.model_g()
    print('finish_'+airport)

#fish