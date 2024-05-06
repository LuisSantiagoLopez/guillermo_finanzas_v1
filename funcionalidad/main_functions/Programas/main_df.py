# -*- coding: utf-8 -*-
"""
Created on Fri Jun 23 11:40:43 2023

@author: jpnarchi

Este programa se encarga de limpiar el archivo de Excel "Resumen_de_Mercado.xls" y generar un archivo Excel por cada sección del archivo original.
"""


#!/usr/bin/env python3
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys


def Borrar_Columnas_NaN(df):
    
    for k in range(df.shape[0]):
        if pd.isnull(df.iloc[k, 0]):
            df = df.iloc[:k]
            break
    # Itera sobre las columnas del DataFrame
    for columna in df.columns:
        # Verifica si todos los valores en la columna son NaN
        if df[columna].isnull().all():
            # Obtiene el índice de la columna y elimina todas las columnas a la derecha
            indice_columna = df.columns.get_loc(columna)
            df = df.iloc[:, :indice_columna]
            break  # Termina el bucle después de eliminar la primera columna NaN
    
    return df

def Generador_Foto(df, nombre):
    fig, ax = plt.subplots(figsize=(5, 6))
    tabla = plt.table(cellText=df.values,
                      colLabels=df.columns,
                      cellLoc='center',
                      loc='center',
                      bbox=None)  # No modificar el tamaño de las celdas
    
    tabla.auto_set_font_size(False)
    # Ajustar automáticamente el ancho de las columnas
    tabla.auto_set_column_width(col=list(range(len(df.columns))))
    tabla.set_fontsize(14)
    tabla.scale(1, 2)
    
    # Establecer la primera fila en negrita y con fondo azul
    for j, label in enumerate(df.columns):
        cell = tabla.get_celld()[(0, j)]
        cell.set_text_props(weight='bold')
        cell.set_facecolor("#4682B4")  # Azul claro
    
    # Establecer los colores de fondo alternados para las filas
    for i in range(1, len(df) + 1):
        if i % 2 != 0:
            color = 'lightgray'  # Fila gris claro
        else:
            color = 'white'  # Fila blanca
        for j, label in enumerate(df.columns):
            cell = tabla.get_celld()[(i, j)]
            cell.set_facecolor(color)
    
    ax.axis('off')
    ax.set_title(nombre, loc='center', size=16, fontweight='bold')  # Ajustar posición, tamaño y estilo del título
    fig.tight_layout(pad=0)
    plt.tight_layout()

    plt.savefig(f'Imagenes/{nombre}.png', format='png', bbox_inches='tight', dpi=300)
    plt.close(fig)

def Limpiador_DF(df):
    lista = df.iloc[0].tolist()
    df_copy = df.copy()
    
    if any("Unnamed:" in col for col in df.columns):
        df_copy.columns = lista
        # Eliminar la primera fila solo si hay columnas con "Unnamed:"
        df_copy = df_copy.iloc[1:]
        df_copy.reset_index(drop=True, inplace=True)
    
    df_copy.replace({np.nan: "-"}, inplace=True)
    return df_copy

ruta_archivo = "Data_a_Extraer/Resumen_de_Mercado.xls"

df_original = pd.read_excel(ruta_archivo)

### Buscamos la fecha actual
fecha = 0
for l in range(df_original.shape[0]):
    if "Tasa de Fondeo (Banxico)" in str(df_original['Unnamed: 1'].iloc[l]):
        fecha = str(df_original['Unnamed: 1'].iloc[l].split(" ")[-1])
        print(fecha)
        break

### Establecemos variables
buscar = ["CETE's",'Bonos de Tasa Fija (BONOS M)',"Tipos de Cambio","IRS TIIE",f"Tasa de Fondeo (Banxico) Jornada del {fecha}",
          "Basis Swap TIIE SOFR","Reportos Gubernamentales","TIIE Banxico","Puntos Forward","Tasas Reales (UDIBONO's)","Tasas Reales (CBIC's)","UDI-TIIE"]
df_cetes = 0
df_tipo_de_cambio = 0
df_reportos_gubernamentales = 0
df_irs_tiie = 0
df_tasa_fondeo = 0
df_basic_swap = 0
df_tiie_banxico = 0
df_puntos_forward = 0
df_tasas_reales_udibonos = 0
df_cbics = 0
df_udi_tiie = 0

for i,b in enumerate(buscar):
    indx = 0
    
    col = 'Unnamed: 1'
    cols_2 = [n for n in range(0,12)]
    if i > 8:
        col = 'Unnamed: 12'
        cols_2 = [n for n in range(12,19)]

        
    for n in range(df_original.shape[0]):
        check = str(df_original[col].iloc[n])
        if check.find(b) != -1:
    
           if pd.isnull(df_original[col].iloc[n+1]) == True:
               indx = n+3
           else:
               indx = n+2
               break
   
           
    df_check = pd.read_excel(ruta_archivo, skiprows=indx, usecols=cols_2)
    
    if i <= 8:
        df_check = df_check.drop(df_check.columns[0], axis=1)
        
    df_organizado = Borrar_Columnas_NaN(df_check)
    df_organizado = Limpiador_DF(df_organizado)
    
    if b == buscar[0]:
        df_cetes = df_organizado
        #Agregamos columna de Instrumento al df de CETES
        df_cetes.insert(0, 'Instrumento', ['CETES' for _ in range(df_cetes.shape[0])])
        df_organizado = df_cetes
    elif b == buscar[1]:
        df_bonos_tasa_fija = df_organizado
    elif b == buscar[2]:
        df_tipo_de_cambio = df_organizado
    elif b == buscar[3]:
        df_irs_tiie = df_organizado
    elif b == buscar[4]:
        df_tasa_fondeo = df_organizado
        df_organizado = df_organizado.iloc[:-1, :]
    elif b == buscar[5]:
        df_basic_swap = df_organizado
    elif b == buscar[6]:
        df_reportos_gubernamentales = df_organizado
    elif b == buscar[7]:
        df_tiie_banxico = df_organizado
    elif b == buscar[8]:
        df_organizado = df_organizado.iloc[1:, :]
        df_organizado = df_organizado.iloc[1:, :]
        df_organizado = df_organizado.iloc[:-1, :]
        df_puntos_forward = df_organizado
    elif b == buscar[9]:
        df_tasas_reales_udibonos = df_organizado
        print('HOLA')
    elif b == buscar[10]:
        df_cbics = df_organizado
        df_organizado = df_organizado.iloc[:, :-2]
    elif b == buscar[11]:
        df_organizado = df_organizado.iloc[1:, :]
        df_udi_tiie = df_organizado

    df_organizado.to_excel(f"Exceles/{b}.xlsx",index=False)
    #Generador_Foto(df_organizado, b)

    

import PEMEX_CFE_Resumen_Mercado
import main_writer
import main_op_indeval
import main_writer_indeval