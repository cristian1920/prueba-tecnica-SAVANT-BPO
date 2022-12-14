from genericpath import exists
import win32com.client 
import datetime
import mailbox
import time
from email.mime.base import MIMEBase
from importlib.abc import FileLoader
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate, parseaddr 
from email import encoders
from selenium import webdriver
import os
import win32com
import requests
# from webdriver_manager.chrome import ChromeDriverManager
import chromedriver_autoinstaller
import os
import urllib.request, json
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select

fecha = datetime.datetime.now()
fecha = str(fecha)
anio = datetime.datetime.strptime(fecha,'%Y-%m-%d %H:%M:%S.%f').strftime('%Y')
mes = datetime.datetime.strptime(fecha,'%Y-%m-%d %H:%M:%S.%f').strftime('%m')
dia = datetime.datetime.strptime(fecha,'%Y-%m-%d %H:%M:%S.%f').strftime('%d')

response = requests.get('http://10.8.8.16/v1/RPA_RIS/'+anio+'-'+mes+'-'+dia, headers={'Authorization': '$1$pt/.8s2.$XlXAfZxXhzUV7YSlu2.e30'})
dato=response.json()
for i in dato['data']['estudios']:
    documento=i['documento']
    paciente=i['paciente']
    erp=i['erp']
    servicio=i['servicio']
    estudio=i['estudio']
    fecha_atencion=i['fecha_atencion']
    hora_atencion=i['hora_atencion']
# if (response.status_code != 200):
#     print("API is not working. status code :" + str(response.status_code))
# else:
#     print("API is working. status code :" + str(response.status_code))
#     datas = json.loads(response.text)
#     for value in datas['data']:
#         print(values['estudios'])
    # print(pacientes)
        #print(values['column_name'])
# a=json.loads(response)
# print(a['documento'])
# for i in dato:
#     documento=dato['documento']
#     paciente=dato['paciente']
#     Entidad=dato['erp']
#     servicio=dato['servicio']
#     estudio=dato['estudio']
#     fecha_atencion=dato['fecha_atencion']
#     hora_atencion=dato['hora_atencion']
#     print(documento)