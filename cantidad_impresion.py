import pandas as pd
import re
from datetime import datetime
from os import listdir
import os
import time
import xlwt
from xlwt.Workbook import *
from pandas import ExcelWriter
import xlsxwriter

def tipo_archivo(texto):
    texto = str(texto)
    texto = texto.lower()
    
    sinco = re.compile(r'^\d+$')
    
    pdf = re.compile(r'pdf')
    outlook = re.compile(r'outlook')
    powerpoint = re.compile(r'powerpoint')
    publisher = re.compile(r'in-store publisher')
    #Excel
    excel = re.compile(r'xls')
    csv = re.compile(r'csv')
    libro = re.compile(r'libro')
    hoja = re.compile(r'hoja')
    #Word
    word = re.compile(r'microsoft word')
    documento = re.compile(r'documento')
    #Web
    http = re.compile(r'http')
    html = re.compile(r'html')
    php = re.compile(r'php')
    aspx = re.compile(r'aspx')
    jsp = re.compile(r'jsp')
    #Correo Web
    correo = re.compile(r'correo')
    responder = re.compile(r're:')
    reenviar = re.compile(r'rv:')
    #Tipo Texto
    txt = re.compile(r'txt')
    prn = re.compile(r'prn')
    bloc = re.compile(r'bloc de notas')
    #Imagenes
    gif = re.compile(r'gif')
    oxps = re.compile(r'oxps')
    png = re.compile(r'png')
    jpg = re.compile(r'jpg')
    img = re.compile(r'img')
    #Recibos
    voucher = re.compile(r'voucher')
    recibo = re.compile(r'recibo')
    orden = re.compile(r'orden de')
    factura = re.compile(r'factura')
    eep = re.compile(r'eep')

    if pdf.search(texto):
        resultado = "PDF"
    elif outlook.search(texto):
        resultado = "Outlook"
    elif powerpoint.search(texto):
        resultado = "PowerPoint"
    elif publisher.search(texto):
        resultado = "Tabloide"
#PENDIENTE CLASIFICAR TIPO ARCHIVO
    elif eep.search(texto):
        resultado = "Archivo de Texto"
    elif excel.search(texto) or csv.search(texto) or libro.search(texto) or hoja.search(texto):
        resultado = "Excel"
    elif word.search(texto) or documento.search(texto):
        resultado = "Word"
    elif http.search(texto) or html.search(texto) or php.search(texto) or aspx.search(texto) or jsp.search(texto):
        resultado = "Pagina Web"
    elif correo.search(texto) or responder.search(texto) or reenviar.search(texto):
        resultado = "Correo Web"
    elif txt.search(texto) or prn.search(texto) or bloc.search(texto):
        resultado = "Archivo de Texto"
    elif gif.search(texto) or oxps.search(texto) or png.search(texto) or jpg.search(texto) or img.search(texto):
        resultado = "Imagen"
    elif voucher.search(texto) or recibo.search(texto) or orden.search(texto) or factura.search(texto):
        resultado = "Recibos"
    elif sinco.search(texto):
        resultado = "Sinco"
    else:
        resultado = "Otros"
    return resultado 

ruta = os.getcwd()+'\\data_export'
lista_export = listdir(ruta)

data = pd.concat([pd.read_csv('data_export/'+f,  delimiter=";", low_memory=False) for f in lista_export])

#crear cada una de las agrupaciones
data['clase_archivo'] = data['PRINTJOBNAME'].apply(lambda x: tipo_archivo(x))
data['mes_impresion'] = pd.DatetimeIndex(data['SUBMITDATE']).month
data['dia_impresion'] = pd.DatetimeIndex(data['SUBMITDATE']).day
data_modelo = data.groupby('RELEASEMODEL')['NUMPAGES'].sum()
data_dia = data.groupby('dia_impresion')['NUMPAGES'].sum()
data_archivo = data.groupby('clase_archivo')['NUMPAGES'].sum()
data_marca = data.groupby('SITE')['NUMPAGES'].sum()
data_usuario = pd.DataFrame(data.groupby('USERID')['NUMPAGES'].sum())
data_usuario = data_usuario.sort_values(by='NUMPAGES', ascending=False).head(100)
data_agrupada = pd.DataFrame(data.groupby(['dia_impresion', 'SITE', 'clase_archivo',  'RELEASEMODEL'])['NUMPAGES'].sum())
data_agrupada_usuario = pd.DataFrame(data.groupby(['dia_impresion', 'SITE', 'USERID', 'clase_archivo',  'RELEASEMODEL'])['NUMPAGES'].sum())

#grabar en cada excel
writer = pd.ExcelWriter('archivo_final.xlsx')
data_modelo.to_excel(writer, 'data_modelo')
data_dia.to_excel(writer, 'data_dia')
data_archivo.to_excel(writer, 'data_archivo')
data_marca.to_excel(writer, 'data_marca')
data_usuario.to_excel(writer, 'data_usuario')
data_agrupada.to_excel(writer, 'data_agrupada')
data_agrupada_usuario.to_excel(writer, 'data_agrupada_usuario')
writer.save()
