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

libro_lectura = cliente.open_by_key("1XS1ovDBG99Wnyai4BZjlxQogkxlOArnd4w2NbLirKAM")

#Informacion PDV
hoja_lectura = libro_lectura.worksheet('BASE_PDV')

datos = hoja_lectura.get_all_values()
df_pdv = pd.DataFrame(datos).set_index([0])[2:]
df_pdv_renombrado = df_pdv.rename(columns={2:'Zona', 4:'Dni_Responsable', 5:'Nombre_Pdv', 7:'Direccion', 10:'Distrito', 24:'Estado', 40:'Supervisor'})
#Resultado Final
df_new_pdv = df_pdv_renombrado[['Zona','Dni_Responsable','Nombre_Pdv','Direccion','Distrito','Estado','Supervisor']]
