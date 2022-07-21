import msvcrt
from re import L
import shutil
import sys
import time
import os
from datetime import datetime
import pyexcel
import xlwt

class Excel_1:
    

    def crear_xls(self,wb,i,tipoDocumento,dia,mes,ano,ciudad):
        if tipoDocumento == "ESTADOS": nombre_archivo = f"{i}.xls"
        elif tipoDocumento == "TRASLADOS": nombre_archivo = f"{i} TRASLADO.xls"
        elif tipoDocumento == "EDICTOS": nombre_archivo = f"{i} EDICTO.xls"
        elif tipoDocumento == "FIJACIONES": nombre_archivo = f"{i} FIJACION.xls"
        elif tipoDocumento == "SENTENCIAS": nombre_archivo = f"{i} SENTENCIA.xls"
        wb.add_sheet(tipoDocumento)
        if ciudad == "CONSEJO":
            wb.save(f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO\\{nombre_archivo}')
        elif ciudad == "SECCIONES":
            wb.save(f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\{nombre_archivo}')
        else:
            wb.save(f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\{nombre_archivo}')

    def escribir_xls(self,datos,i,tipoDocumento,dia,mes,ano,ciudad):
        if tipoDocumento == "ESTADOS": nombre_archivo = f"{i}.xls"
        elif tipoDocumento == "TRASLADOS": nombre_archivo = f"{i} TRASLADO.xls"
        elif tipoDocumento == "EDICTOS": nombre_archivo = f"{i} EDICTO.xls"
        elif tipoDocumento == "FIJACIONES": nombre_archivo = f"{i} FIJACION.xls"
        elif tipoDocumento == "SENTENCIAS": nombre_archivo = f"{i} SENTENCIA.xls"
        if ciudad == "CONSEJO":
            wb = pyexcel.get_book(file_name=f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO\\{nombre_archivo}')
        elif ciudad == "SECCIONES":
            wb = pyexcel.get_book(file_name=f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\{nombre_archivo}')
        else:
            wb = pyexcel.get_book(file_name=f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\{nombre_archivo}')
        for dats in datos:
            wb.sheet_by_name(tipoDocumento).row += dats
        if ciudad == "CONSEJO":
            wb.save_as(f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO\\{nombre_archivo}')
        elif ciudad == "SECCIONES":
            wb.save_as(f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\{nombre_archivo}')
        else:
            wb.save_as(f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\{nombre_archivo}')
    
    def crear_xls_revision(self,wb,dia,mes,ano,ciudad):

        data = {'ESTADOS': ['JUZGADO','FECHA ESTADO'],
                'TRASLADOS': ['JUZGADO','FECHA TRASLADO'],
                'ERRORES': ['JUZGADO','TIPO','FECHA']}
        for key, nomHoja in enumerate(data):
            ws = wb.add_sheet(nomHoja)
            for clave, valor in enumerate(data[nomHoja]):
                ws.write(0, clave, valor)
        if ciudad == "CONSEJO":
            wb.save(f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO\\REVISION.xls')
        elif ciudad == "SECCIONES":
            wb.save(f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\REVISION.xls')
        else:
            wb.save(f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\REVISION.xls')

    def escribir_xls_revision(self,estadosRevision,trasladosRevision,error,dia,mes,ano,ciudad):
        if ciudad == "CONSEJO":
            wb = pyexcel.get_book(file_name=f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO\\REVISION.xls')
        elif ciudad == "SECCIONES":
            wb = pyexcel.get_book(file_name=f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\REVISION.xls')
        else:
            wb = pyexcel.get_book(file_name=f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\REVISION.xls')
        
        
        for dats in estadosRevision:
            wb.sheet_by_name('ESTADOS').row += dats
        
        for dats in trasladosRevision:
            wb.sheet_by_name('TRASLADOS').row += dats
        
        for dats in error:
            wb.sheet_by_name('ERRORES').row += dats        

        if ciudad == "CONSEJO":
            wb.save_as(f'.\\{ano}\\{mes}\\{dia}\\CALI\\CONSEJO\\REVISION.xls')
        elif ciudad == "SECCIONES":
            wb.save_as(f'.\\{ano}\\{mes}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\REVISION.xls')
        else:
            wb.save_as(f'.\\{ano}\\{mes}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\REVISION.xls')
    
    
