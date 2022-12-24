import openpyxl as xl
import shutil
from datetime import date
from datetime import timedelta
import os
from os import remove
from os import replace

from pickle import TRUE

### Variables
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop/') 
documents = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents/Diarios/') 

today = date.today()
aaaammdd = today.strftime("%Y%m%d")
yesterday = today - timedelta(days = 1)
aaaammdd_ayer = yesterday.strftime("%Y%m%d")
tomorrow = today + timedelta(days = 1)
aaaammdd_tomorrow = yesterday.strftime("%Y%m%d")

#dir_drive='C:/Users/user/Desktop/'
#dir_drive= r"G:/Mi unidad/Diarios/" 

origen=aaaammdd_ayer+'.xlsx'
temporal='temporal.xlsx'
dir_origen=desktop
original = origen
dir_target=desktop
#target = aaaammdd+'.xlsx'
target = aaaammdd +'.xlsx'
print(original)
shutil.copyfile(dir_origen+original, temporal)

wb_data_only = xl.load_workbook(filename=temporal, data_only=TRUE)
wb = xl.load_workbook(filename=temporal)

### Sheet de Salidas ###
salidas = wb.worksheets[0]

tabla_salidas = salidas['B6':'I75']
print(type(tabla_salidas))
for fila in tabla_salidas:
    for celda in fila:
        if(celda.column in (2,3,7) and celda.value!=1):
           # print(str(celda.row) +' '+ str(celda.column))
            celda.value = 1
        if(celda.column == 4 and celda.value!=0):
          #  print(str(celda.row) +' '+ str(celda.column))
            celda.value = 0     
        if(celda.column in(5,9) and celda.value!=None):
           # print(str(celda.row) +' '+ str(celda.column))
            celda.value = ''

tabla_canchas = salidas['M5':'Q19']
for fila in tabla_canchas:
    for celda in fila:
        if(celda.column == 13 and celda.value!=0):
          #  print(str(celda.row) +' '+ str(celda.column))
            celda.value = 0     
        if(celda.column in (14,15,16,17) and celda.value!=None):
           # print(str(celda.row) +' '+ str(celda.column))
            celda.value = ''

tabla_gastos = salidas['N23':'O30']
for fila in tabla_gastos:
    for celda in fila:
        if(celda.column in (14,15) and celda.value!=None):
          #  print(str(celda.row) +' '+ str(celda.column))
            celda.value = ''


### Sheet de Bebidas ###
bebidas = wb.worksheets[1]
bebidas_data_only=wb_data_only.worksheets[1]
stock_inicial = bebidas['E2':'E22']
stock_actual = bebidas_data_only['H2':'H22']
for fila1, fila2 in zip(stock_inicial,stock_actual):
    for celda1, celda2 in zip(fila1,fila2):
        print(celda2.value)
        celda1.value = celda2.value

### Sheet de Entradas ###
entradas = wb.worksheets[2]

tabla_entradas = entradas['A2':'D24']
for fila in tabla_entradas:
    for celda in fila:
        if(celda.column in (3,4) and celda.value!=None):
          #  print(str(celda.row) +' '+ str(celda.column))
            celda.value = 0


wb.save(desktop+target)
wb.close()
#replace(dir_origen+origen,dir_drive+origen)
wb_data_only.close()
#shutil.copyfile(temporal,dir_drive+origen)
shutil.copyfile(temporal,documents+origen)
remove(temporal)
remove(desktop+origen)

os.system("pause")
