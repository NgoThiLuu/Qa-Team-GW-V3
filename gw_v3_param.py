from ast import Str
import  json,os,re
from pathlib import Path
from sys import platform
from gw_v3_set_up import driver,datetime

json_file = os.path.dirname(Path(__file__).absolute())+"\\gw_v3_data.json"
if platform == "linux" or platform == "linux2": 
    json_file =  json_file.replace("\\","/") 
with open(json_file) as json_file:
    data = json.load(json_file)


       


   
            
    