from UnixExcel import *
import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
from webbrowser import Chrome
import pyexcel
import xlwt
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, ElementNotVisibleException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC

class Driver_1(object):

    #Declaracion de Variables
    op = webdriver.ChromeOptions()
    op.add_experimental_option('excludeSwitches', ['enable-logging'])
    #op.add_argument('--headless')
    #op.add_argument('--disable-gpu')
    op.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=op)
    #Escoger preferencias del WebDriver
    def __init__(self):  
        os.system ("cls")
        
    #Cargar la pagina solicitada
    def load_page(self,tipoDocumento):
         try:
            if tipoDocumento == "ESTADOS": self.driver.get("https://samairj.consejodeestado.gov.co/Vistas/utiles/WEstados")
            elif tipoDocumento == "TRASLADOS": self.driver.get("https://samairj.consejodeestado.gov.co/Vistas/utiles/WTraslados.aspx")
            elif tipoDocumento == "EDICTOS": self.driver.get("https://samairj.consejodeestado.gov.co/Vistas/utiles/WONotificaciones.aspx?guid=Vedictos")
            elif tipoDocumento == "FIJACIONES": self.driver.get("https://samairj.consejodeestado.gov.co/Vistas/utiles/WFijacionLista.aspx")
            elif tipoDocumento == "SENTENCIAS": self.driver.get("https://samairj.consejodeestado.gov.co/Vistas/utiles/WONotificaciones.aspx?guid=VEstadoSentencia")
            WebDriverWait(self.driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="titulosede"]')))
         except TimeoutException:
            print('line: 38 error: No se cargo la pagina, TimeoutException')
         except:
            print('Error Interno')