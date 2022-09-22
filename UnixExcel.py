from Consultor import *
from Driver import *
from Excel import Excel_1
from Ruta import *
import msvcrt
import sys
import pyautogui

#Clase principal   
if __name__ == "__main__":

    #Se llaman las clases a trabajar
    browser =Driver_1()
    consulta =Consultor_1
    excel = Excel_1()
    ruta = rutaCiudad()
    
    dia = time.strftime("%d")
    mes = time.strftime("%m")
    ano = time.strftime("%Y")
    fecha_actual = f"{dia}/{mes}/{ano}"
    tipoDocumento = ["ESTADOS","TRASLADOS","EDICTOS","FIJACIONES","SENTENCIAS"]
    excelFile= ["CALI","BARRANQUILLA","BUCARAMANGA","CARTAGENA","MEDELLIN","MONTERIA"
                ,"NEIVA","PASTO","POPAYAN","QUIBDO","RIOHACHA","SAN ANDRES","SANTA MARTA","SINCELEJO"
                ,"TUNJA","TURBO","VALLEDUPAR","VILLAVICENCIO","CONSEJO","SECCIONES"]
    
    
    print("*"*30+"SAMAI"+"*"*30)
    
    tipoEstado = int(input("\n1.Cali\n2.Consejo\n3.Otros\n4.Secciones\n5.Salir\nIngrese: "))
    if tipoEstado == 1: excelFile = excelFile[0]
    elif tipoEstado == 2: excelFile = excelFile[-2]
    elif tipoEstado == 3: excelFile = excelFile[1:-2]
    elif tipoEstado == 4: excelFile = excelFile[-1]
    elif tipoEstado == 5: 
        print("\nPULSE UNA TECLA PARA CERRAR...")
        sys.exit()
    else: 
        print("\nOpcion incorrecta")
        print("\nPULSE UNA TECLA PARA CERRAR...")

    for ciudad in excelFile:
        if tipoEstado == 1 or tipoEstado == 2 or tipoEstado == 4: ciudad = excelFile
        estadosRevision = []
        trasladosRevision = []
        error = []
        ruta.crear_carpetas(dia=dia, mes=mes, ano=ano, ciudad=ciudad)
        my_array = pyexcel.get_array(file_name=f'.\\Ciudades\\{ciudad}.xlsx', start_row=0)
        
        print(f"\033[1;31;47m Ejecutando {ciudad}")  
        woutrevision = xlwt.Workbook()
        excel.crear_xls_revision(wb=woutrevision,dia=dia,mes=ruta.dife_fecha(),ano=ano,ciudad=ciudad)
        
        for tipo in tipoDocumento:
            print("\033[0;37m"+"*"*80)
            print(f"\033[1;32m Ejecutando {tipo} de {ciudad}")
            browser.load_page(tipo)
            consulta.ingresar_corporacion((browser.driver),my_array[0][0])
            contador = 1
            for i in my_array[1:]:
                try:
                    if i[0] == "": break
                    
                    juzgadoxls = []
                    print(f"\033[0;37m Ejecutando {i[0]} de {ciudad}")
                    consulta.ingresar_juzgado((browser.driver),i[0],contador,ciudad)
                    datos = []
                    fecha = consulta.verificar_fecha((browser.driver),fecha_actual)
                    if fecha == fecha_actual:
                        #Creamos la hoja de excel y la mandamos como parametro
                        wout = xlwt.Workbook()
                        excel.crear_xls(wb=wout,i=i[0],tipoDocumento=tipo,dia=dia,mes=ruta.dife_fecha(),ano=ano,ciudad=ciudad)
                        consulta.descargar((browser.driver),datos)
                        excel.escribir_xls(datos,i[0],tipo,dia,ruta.dife_fecha(),ano,ciudad)
                        
                    else:
                        
                        browser.driver.find_element(By.XPATH,'//*[@id="MainContent_LstUEstados"]').click()
                        if tipo == "ESTADOS":
                            juzgadoxls.append(i[0])
                            juzgadoxls.append(fecha)
                            estadosRevision.append(juzgadoxls)
                        if tipo == "TRASLADOS":
                            juzgadoxls.append(i[0])
                            juzgadoxls.append(fecha)
                            trasladosRevision.append(juzgadoxls)
                            
                        if ciudad == "CONSEJO":
                                browser.driver.save_screenshot(f'.\\{ano}\\{ruta.dife_fecha()}\\{dia}\\CALI\\CONSEJO\\NH\\{i[0]} {tipo}.png')
                        elif ciudad == "SECCIONES":
                                browser.driver.save_screenshot(f'.\\{ano}\\{ruta.dife_fecha()}\\{dia}\\BOGOTA\\SALAS Y SECCIONES\\NH\\{i[0]} {tipo}.png')
                        else:
                                browser.driver.save_screenshot(f'.\\{ano}\\{ruta.dife_fecha()}\\{dia}\\{ciudad}\\ADMINISTRATIVOS\\NH\\{i[0]} {tipo}.png')
                    contador+=1
                except:
                    juzgadoerror = []
                    print(f"ERROR EN {tipo} de {i[0]} - {ciudad} ")
                    juzgadoerror.append(i[0])
                    juzgadoerror.append(tipo)
                    juzgadoerror.append(fecha)
                    error.append(juzgadoerror)
            print(f" \033[1;32m Finalizó {tipo} de {ciudad}")
        excel.escribir_xls_revision(estadosRevision,trasladosRevision,error,dia,ruta.dife_fecha(),ano,ciudad)
                
        if ciudad == "CALI" or ciudad == "CONSEJO" or ciudad == "SECCIONES": break
        print("\033[0;37m"+"*"*80)
        print(f"\033[1;31;47m Finalizó {ciudad}")
        print("\033[0;37m"+"*"*80)
        
            

    
        
        
        

        
