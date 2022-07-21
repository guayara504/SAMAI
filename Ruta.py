import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
import pyexcel
import xlwt

class rutaCiudad:
    
    
    def __init__(self):
        pass
    
    def dife_fecha(self):
        mes= time.strftime("%m")
        if mes == '01':
            mes = 'ENERO'
        if mes == '02':
            mes = 'FEBRERO'
        if mes == '03':
            mes = 'MARZO'
        if mes == '04':
            mes = 'ABRIL'
        if mes == '05':
            mes = 'MAYO'
        if mes == '06':
            mes = 'JUNIO'
        if mes == '07':
            mes = 'JULIO'
        if mes == '08':
            mes = 'AGOSTO'
        if mes == '09':
            mes = 'SEPTIEMBRE'
        if mes == '10':
            mes = 'OCTUBRE'
        if mes == '11':
            mes = 'NOVIEMBRE'
        if mes == '12':
            mes = 'DICIEMBRE'
        return mes
    
    def crear_carpetas(self,dia,mes,ano,ciudad):
        try:
                os.mkdir(f'.\\{ano}')
        except:
                pass

        mes= self.dife_fecha()
        try:
                os.mkdir(f'.\\{ano}\\{mes}')
        except:
                pass

        try:
                os.mkdir(f'.\\{ano}\\{mes}\\{dia}')
        except:
                pass
        
        try:
                if ciudad == "CONSEJO":
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\CALI')
                elif ciudad == "SECCIONES":
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\BOGOTA')
                else:  
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\{ciudad}')
                    
        except:
                pass
        
        try:
                if ciudad == "CONSEJO":
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO')
                if ciudad == "SECCIONES":
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES')
                else:
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS')
        except:
                pass
        
        try:
                if ciudad == "CONSEJO":
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO\\NH')
                if ciudad == "SECCIONES":
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\NH')
                else:
                    os.mkdir(f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\NH')
        except:
                pass
        