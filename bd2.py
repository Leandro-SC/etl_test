import pandas as pd
import numpy as np
import pyodbc as dbc
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Declarar constantes
server = 'AUREN22\AUREN'
bd = 'Intranet'
usuario = 'eIntranet'
contrasena = '6rupoAuren'

# rutaArchivo = "C:\\Users\\leandro.silva\\Desktop\\TEST"

# Conectar con la base
def conectarBD():
    try:
        conexion = dbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER=' +
                               server+';DATABASE='+bd+';ENCRYPT=no;UID='+usuario+';PWD=' + contrasena)
        cursor = conexion.cursor()
        # query
        query_sql = f"SELECT idAutor as codigo,DAY(fecha) as dia,MONTH(fecha) as mes,consumo, operacion,CASE WHEN operacion = 'Postpago Migracion M4' THEN 'POST'WHEN operacion = 'Postpago Portabilidad Migración M4 (Or. Postpago)' THEN 'O_POST'WHEN operacion = 'Postpago Portabilidad ( Origen Postpago )' THEN 'O_POST'WHEN operacion = 'Postpago Renueva por Fidelización' THEN 'POST'WHEN operacion = 'Postpago Portabilidad Migración M4 (Or. Prepago)' THEN 'POST'WHEN operacion = 'Postpago Alta' THEN 'POST'WHEN operacion = 'Postpago Portabilidad ( Origen Prepago )' THEN 'POST'WHEN operacion = 'Postpago Renueva migracion de Pre a Post' THEN 'POST'WHEN operacion = 'Migracion de Prepago a Postpago' THEN 'POST'WHEN operacion = 'Postpago Renueva  por Fidelización' THEN 'POST' ELSE 'PRODUCTO DESCONOCIDO'END AS TipoVentaProducto,CASE WHEN LEN(imei) > 3 THEN 'EQUIPO'ELSE 'CHIP'END AS TipoVenta,COUNT(id) as Cantidad FROM [Intranet].[dbo].[controlnet] WHERE MONTH(fecha)>7 GROUP BY idAutor, consumo,operacion,DAY(fecha),MONTH(fecha),imei ORDER BY  MONTH(fecha);"
        # Convertir a dataframe ventas
        consulta = pd.read_sql_query(query_sql, conexion, index_col='codigo')
        df_renombrado = consulta.rename_axis('ID VENTA')
        df_ventas = df_renombrado.where((df_renombrado.operacion !='PRODUCTO DESCONOCIDO')).fillna(0)
        print("Conexión exitosa")
        return df_ventas
       
        # convertir a excel
        '''wb = Workbook()
        filesheet = rutaArchivo+"/VENTAS_DAM_New.xlsx"
        ws = wb.active
        for r in dataframe_to_rows(consulta, index=False, header=True):
                ws.append(r)
        wb.save(filesheet)
            # print(consulta)'''

        
    except:
        print("Falló la conexión")

# conectarBD()
