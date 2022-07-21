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
from selenium.webdriver.common.keys import Keys

class Consultor_1():
    
    #Ingresar a "Todos los procesos"
    def ingresar_corporacion(driver,corporacion):
        Select(driver.find_element(By.XPATH,'//*[@id="MainContent_LstCorpHabilitada"]')).select_by_visible_text(corporacion)
        try:
            driver.find_element(By.XPATH,'//*[@id="MainContent_ImgBuscar2"]').click()
        except: pass

    #Ingresar el radicado en la caja
    def ingresar_juzgado(driver,juzgado,contador,ciudad):
        driver.implicitly_wait(3)
        Select(driver.find_element(By.XPATH,'//*[@id="MainContent_LstCoorporacion"]')).select_by_visible_text(juzgado)
        try:
            driver.find_element(By.XPATH,'//*[@id="MainContent_ImgBuscar3"]').click()
        except: pass
        if ciudad != "CONSEJO" and contador==1:
            if ciudad != "SECCIONES":
                driver.find_element(By.XPATH,'//*[@id="MainContent_ChkSeccion"]').click()
            

    def verificar_fecha(driver,fecha_actual):
        fecha = driver.find_element(By.XPATH,'//*[@id="MainContent_LstUEstados"]').text
        if fecha[2:12] == fecha_actual or fecha[3:13] == fecha_actual or fecha[1:11] == fecha_actual: return fecha_actual
        return fecha[2:12]
        
    def descargar(driver,datos):
    
        driver.find_element(By.XPATH,'//*[@id="MainContent_CmdBuscar"]').click()
        WebDriverWait(driver,timeout=4).until(EC.presence_of_element_located((By.XPATH,'//*[@id="MainContent_PanelProvidencias"]/table/tbody/tr[2]/td')))
        
        table_actos = driver.find_element(By.XPATH,'//*[@id="MainContent_GvProvidencias"]/tbody')
        allrows = table_actos.find_elements(By.TAG_NAME,"tr")[:]
        for tr in allrows:
            lista_td = []
            allcols = tr.find_elements(By.TAG_NAME,"td")[1:-1]
            for j in range(len(allcols)):
                lista_td.append(allcols[j].text)
            datos.append(lista_td)
        print("Guardado")
        
        
        
