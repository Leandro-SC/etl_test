import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

scope = ['https://www.googleapis.com/auth/spreadsheets',
        "https://www.googleapis.com/auth/drive"]

credenciales = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json", scope)
cliente = gspread.authorize(credenciales)

# rutaArchivo = "C:\\Users\\leandro.silva\\Desktop\\TEST"

#sheet = cliente.create("Base desde Python")
#sheet.share('lsilvac4@gmail.com', perm_type='user', role='reader')
libro_lectura = cliente.open_by_key("1XS1ovDBG99Wnyai4BZjlxQogkxlOArnd4w2NbLirKAM")
#Informacion FFVV
hoja_lectura = libro_lectura.worksheet('ACTIVOS')
datos = hoja_lectura.get_all_values()
df_ffvv = pd.DataFrame(datos).set_index([11])[1:]
df_ffvv_renombrado = df_ffvv.rename(columns={0:'Zona',1:'Fecha_Ingreso', 2:'Tipo_Doc.', 3:'Numero_Doc', 4:'Nombres', 5:'Apellido_Paterno', 6:'Apellido_Materno', 7:'Telefono', 8:'Estado', 9:'Cargo', 10:'Bodega', 12:'Supervisor'}, inplace=True)
df_ffvv['Nombre_Completo'] = df_ffvv['Nombres'].astype(str) + " " + df_ffvv['Apellido_Paterno'].astype(str)+" "+ df_ffvv['Apellido_Materno'].astype(str)
df_ffvv_concatenado = df_ffvv.drop(['Nombres','Apellido_Paterno','Apellido_Materno'], axis=1)
#Resultado Final
df_ffvv_final = df_ffvv_concatenado.rename_axis('ID VENTA')


# wb = Workbook()
# filesheet = rutaArchivo+"/SELLERS_DAM.xlsx"
# ws= wb.active
# for r in dataframe_to_rows(df, index=False, header=False):
#     ws.append(r)
# wb.save(filesheet)
