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
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
main_functions_dir = os.path.dirname(script_dir)

excel_output_dir = os.path.join(main_functions_dir, 'Exceles')
image_output_dir = os.path.join(main_functions_dir, 'Imagenes')

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


ruta_archivo = os.path.join(main_functions_dir, 'Data_a_Extraer', 'Resumen_de_Mercado.xls')

df_original = pd.read_excel(ruta_archivo)

buscar2 = ["Tipos de Cambio","Tasa de Banxico","CETE's","Tasas Reales (UDIBONO's)","Tasas Reales (CBIC's)","Bonos de Tasa Fija (BONOS M)",
           "IRS TIIE","UDI-TIIE","Puntos Forward","Basis Swap TIIE SOFR"]

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

for i,b in enumerate(buscar2):
    indx = 10
    
    col = 'Unnamed: 1'
    cols_2 = [n for n in range(1,5)]
    if i > 1 and i < 3:
        col = 'Unnamed: 7'
        cols_2 = [n for n in range(7,11)]
    elif i > 2 and i < 6:
        col = 'Unnamed: 18'
        cols_2 = [n for n in range(18,25)]
    elif i >= 6:
        col = 'Unnamed: 37'
        cols_2 = [n for n in range(37,43)]
 
    
    for n in range(df_original.shape[0]):
        check = str(df_original[col].iloc[n])
        if check.find(b) != -1:
           if pd.isnull(df_original[col].iloc[n+1]) == True:
               indx = n+3
           else:
               indx = n+2
               break

    df_check = pd.read_excel(ruta_archivo, skiprows=indx, usecols=cols_2)
    
    #if i <= 8:
     #   df_check = df_check.drop(df_check.columns[0], axis=1)
        
    df_organizado = Borrar_Columnas_NaN(df_check)
    df_organizado = Limpiador_DF(df_organizado)
    
    
    if b == buscar2[0]:
        df_tipo_de_cambio = df_organizado

    elif b == buscar2[1]:
        df_tasa_fondeo = df_organizado

    elif b == buscar2[3]:     
        df_tasas_reales_udibonos = df_organizado
        df_tasas_reales_udibonos = df_tasas_reales_udibonos.drop(columns=df_tasas_reales_udibonos.columns[-1:])
        df_organizado = df_tasas_reales_udibonos
    elif b == buscar2[2]:     
        df_cetes = df_organizado
        #Agregamos columna de Instrumento al df de CETES
        df_cetes.insert(0, 'Instrumento', ['CETES' for _ in range(df_cetes.shape[0])])
        df_organizado = df_cetes   
        df_cetes = df_organizado
    elif b == buscar2[4]:     
        df_cbics = df_organizado
        df_cbics = df_cbics.drop(columns=df_cbics.columns[-2:])
        df_organizado = df_cbics
    elif b == buscar2[5]:     
        df_bonos_tasa_fija = df_organizado
        df_bonos_tasa_fija = df_bonos_tasa_fija.drop(columns=df_bonos_tasa_fija.columns[-2:])
        df_organizado = df_bonos_tasa_fija
    elif b == buscar2[6]:     
        df_irs_tiie = df_organizado
        print(df_irs_tiie)
    elif b == buscar2[7]:     
        df_udi_tiie = df_organizado
    elif b == buscar2[8]:     
        df_puntos_forward = df_organizado
    elif b == buscar2[9]:     
        df_basic_swap = df_organizado    

    df_organizado.to_excel(os.path.join(excel_output_dir, f"{b}.xlsx"), index=False)
    #Generador_Foto(df_organizado, b)

    
import PEMEX_CFE_Resumen_Mercado
import main_writer
import main_op_indeval
import main_writer_indeval
