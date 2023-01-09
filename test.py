import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


nueva_carpeta = "nueva_carpeta"
rutaArchivo = "C:\\Users\\leandro.silva\\Desktop\\GLORIA_TRANSP"
nombreArchivo = "HOJA DE RUTA I.xlsx"
#Conocer la ruta
ruta = os.getcwd()

#Encontrar carpeta de archivo
def findfile(name, path):
    for dirpath, dirname, filename in os.walk(path):
        if name in filename:
            return os.path.join(dirpath, name)
filepath = findfile(nombreArchivo, rutaArchivo)
lastfilepath = filepath.replace("\\","/")
hojaLectura = 'DATA'

#Convertir a Dataframe
dfClientes = pd.read_excel(filepath,sheet_name=hojaLectura, usecols='F,G,H,I,J,K,L,S,T,U,V,W,X,Y', index_col=0)
dfClientesLimpio = dfClientes.drop_duplicates('NOMBRE CONTACTO',keep='last')
dfClientes2 = dfClientesLimpio.rename(columns={'NOMBRE CONTACTO':'CLIENTE'})
print(dfClientes2)

def crearExcel():
    wb = Workbook()
    filesheet = rutaArchivo+"/clientes.xlsx"
    ws= wb.active
    for r in dataframe_to_rows(dfClientesLimpio, index=True, header=True):
        ws.append(r)
    wb.save(filesheet)


# crearExcel()







#Separar tablas

# dfClientes = df.loc[:,['IDENTIFICADOR DE CONTACTO', 'NOMBRE CONTACTO', 'TELEFONO', 'EMAIL', 'CONTACTO', 'DIRECCION', 'LATITUD', 'LONGITUD']]
# print(dfClientes)
# dfDetalles = pd.DataFrame()
# dfGuias = pd.DataFrame()

#Convertir a json

























