import bd2 as bd
import sheet_ffvv as ffvv
import sheet_pdv as pdv
import Class_Excel as xls
import share
from openpyxl.formatting.rule import ColorScale, FormatObject
from openpyxl.styles import Color
import sys, os


def generarReporteMapaDam():
    #Base de ventas
    df_base = bd.conectarBD()
    #Base de FFVV
    df_ffvv = ffvv.df_ffvv_final
    #Base de PDV
    df_pdv = pdv.df_pdv

    #Joins
    df_general = df_base.merge(df_ffvv, left_on='ID VENTA', right_on='ID VENTA')
    df_ventas = df_general.where((df_general.mes == 12) & (df_general.operacion != 'PRODUCTO DESCONOCIDO'))
    #Armar resumen
    tabla_resumen = df_ventas.pivot_table(index=['Nombre_Completo','Zona','Bodega','Supervisor'], columns=['dia'],values= 'Cantidad',aggfunc=sum,fill_value=0).sort_values(by=['Zona','Supervisor'])
    #convertir excel
    tabla_excel = tabla_resumen.to_excel('mapa.xlsx',sheet_name="ventas_dia")
    share.actualizarHoja('mapa.xlsx')

generarReporteMapaDam()












