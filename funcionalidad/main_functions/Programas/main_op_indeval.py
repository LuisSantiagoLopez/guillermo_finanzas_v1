"""Indeval Grande e Indeval Chico en un solo documento de Excel. El ejemplo está en REPORTE - REPORTE INDEVAL"""

import pandas as pd
import os
from datetime import datetime
import locale

script_dir = os.path.dirname(os.path.abspath(__file__))
data_dir = os.path.join(os.path.dirname(script_dir), 'Data_a_Extraer')
excel_output_dir = os.path.join(os.path.dirname(script_dir), 'Exceles')
os.makedirs(excel_output_dir, exist_ok=True)

# Leer los DataFrames
indeval_grande = pd.read_csv(os.path.join(data_dir, "Indeval_1830.csv"))
indeval_chico = pd.read_csv(os.path.join(data_dir, "Indeval_1415.csv"))

vector = pd.read_excel(os.path.join(data_dir, 'VectorAnalitico24h.xls'))
vector_viejo = pd.read_excel(os.path.join(data_dir, 'VectorViejo.xls'))

pip = pd.read_excel(os.path.join(data_dir, 'PIP.xls'), skiprows=1)
pip_viejo = pd.read_excel(os.path.join(data_dir, 'PIPViejo.xls'), skiprows=1)

hoy = datetime.today().date()

fechas_grande = indeval_grande['Fecha']
fechas_chico = indeval_chico['Fecha']

fechas_grande.to_excel(os.path.join(excel_output_dir, 'Fechas_Indeval_Grande.xlsx'), index=False)

fechas_chico.to_excel(os.path.join(excel_output_dir, 'Fechas_Indeval_Chico.xlsx'), index=False)


columnas_df = ['Instrumento','Subyacente','Monto','Dias x Vencer',
               'Tasa Operacion','Sobretasa Operacion', 'Tasa T',
               'Sobretasa T','Tasa T-1','Sobretasa T-1','Tasa T PiP',
               'Sobretasa T PiP','Tasa T-1 PiP','Sobretasa T-1 PiP']

df_final = pd.DataFrame(columns=columnas_df)

for i in range(indeval_grande.shape[0]):
    nombre_emisora = str(indeval_grande['Instrumento'].loc[i])
    
    emisora = str(indeval_grande['Instrumento'].loc[i]).split("_")[1]
    serie = str(indeval_grande['Instrumento'].loc[i]).split("_")[-1]
    
    
    monto = float(indeval_grande['Monto'].loc[i])
    
    ## EXTREAMOS DATOS DE PIP DIARIO
    emisora_pip = pip[pip['EMISORA'] == emisora]
    reducir_por_serie_pip = emisora_pip[emisora_pip['SERIE'] == serie]
    tasa_T_pip = reducir_por_serie_pip['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T_pip = reducir_por_serie_pip['SOBRETASA'].iloc[0]

    ## EXTREAMOS DATOS DE PIP VIEJO
    emisora_pip_viejo = pip_viejo[pip_viejo['EMISORA'] == emisora]
    reducir_por_serie_pip_viejo = emisora_pip_viejo[emisora_pip_viejo['SERIE'] == serie]
    tasa_T_pip_viejo = reducir_por_serie_pip_viejo['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T_pip_viejo = reducir_por_serie_pip_viejo['SOBRETASA'].iloc[0]
    
    ## EXTREAMOS DATOS DE VECTOR DIARIO
    emisora_vector = vector[vector['EMISORA'] == emisora]
    emisora_vector_viejo = vector_viejo[vector_viejo['EMISORA'] == emisora]
    tasa_T = reducir_por_serie_pip['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T = reducir_por_serie_pip['SOBRETASA'].iloc[0]
    
    
    ## EXTRAEMOS CALIFICACIONES Y FECHAS DE VCTO
    mdys = reducir_por_serie_pip['MDYS'].iloc[0].replace('.mx','')
    s_and_p = reducir_por_serie_pip['S&P'].iloc[0].replace('.mx','')
    fecha_vcto = reducir_por_serie_pip['FECHA VCTO'].iloc[0]
    fecha_vcto = fecha_vcto.date()

    
    ## RESTAMOS DÍAS PARA ENCONTRAR DIFERENCIA
    diferencia_dias = fecha_vcto - hoy
    diferencia_dias = diferencia_dias.days
    
    ## EXTREAMOS DATOS DE VECTOR VIEJO
    emisora_vector_viejo = vector_viejo[vector_viejo['EMISORA'] == emisora]
    reducir_por_serie_viejo = emisora_vector_viejo[emisora_vector_viejo['SERIE'] == serie]
    reducir_por_serie_pip = emisora_vector[emisora_vector['SERIE'] == serie]
    
    ## EXTRAEMOS DATOS DE ARCHIVO INDEVAL GRANDE
    tasa_operacion = indeval_grande['Rendimiento'].loc[i]
    sobretasa_operacion = indeval_grande['Sobretasa'].loc[i]
    tasa_T_1 = reducir_por_serie_viejo['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T_1 = reducir_por_serie_viejo['SOBRETASA'].iloc[0]
    
    if mdys == "AAA":
        calificacion = mdys
    else:
        calificacion = s_and_p
    
    fila = [nombre_emisora, calificacion, monto, diferencia_dias, tasa_operacion, 
            sobretasa_operacion, tasa_T, sobretasa_T, tasa_T_1, sobretasa_T_1,
            tasa_T_pip,sobretasa_T_pip,tasa_T_pip_viejo,sobretasa_T_pip_viejo]
    
    ## Añadimos fila a DF
    df_final.loc[i] = fila

df_final.to_excel(os.path.join(excel_output_dir, 'indeval_grande.xlsx'), index=False)

df_final_chico = pd.DataFrame(columns=columnas_df)

for i in range(indeval_chico.shape[0]):
    nombre_emisora = str(indeval_chico['Instrumento'].loc[i])
    
    emisora = str(indeval_chico['Instrumento'].loc[i]).split("_")[1]
    serie = str(indeval_chico['Instrumento'].loc[i]).split("_")[-1]
    
    
    monto = float(indeval_chico['Monto'].loc[i])
    
    ## EXTREAMOS DATOS DE PIP DIARIO
    emisora_pip = pip[pip['EMISORA'] == emisora]
    reducir_por_serie_pip = emisora_pip[emisora_pip['SERIE'] == serie]
    tasa_T_pip = reducir_por_serie_pip['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T_pip = reducir_por_serie_pip['SOBRETASA'].iloc[0]

    ## EXTREAMOS DATOS DE PIP VIEJO
    emisora_pip_viejo = pip_viejo[pip_viejo['EMISORA'] == emisora]
    reducir_por_serie_pip_viejo = emisora_pip_viejo[emisora_pip_viejo['SERIE'] == serie]
    tasa_T_pip_viejo = reducir_por_serie_pip_viejo['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T_pip_viejo = reducir_por_serie_pip_viejo['SOBRETASA'].iloc[0]
    
    ## EXTREAMOS DATOS DE VECTOR DIARIO
    emisora_vector = vector[vector['EMISORA'] == emisora]
    emisora_vector_viejo = vector_viejo[vector_viejo['EMISORA'] == emisora]
    tasa_T = reducir_por_serie_pip['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T = reducir_por_serie_pip['SOBRETASA'].iloc[0]
    
    ## EXTRAEMOS CALIFICACIONES Y FECHAS DE VCTO
    mdys = reducir_por_serie_pip['MDYS'].iloc[0].replace('.mx','')
    s_and_p = reducir_por_serie_pip['S&P'].iloc[0].replace('.mx','')
    fecha_vcto = reducir_por_serie_pip['FECHA VCTO'].iloc[0]
    fecha_vcto = fecha_vcto.date()
    
    ## RESTAMOS DÍAS PARA ENCONTRAR DIFERENCIA
    diferencia_dias = fecha_vcto - hoy
    diferencia_dias = diferencia_dias.days
    
    ## EXTREAMOS DATOS DE VECTOR VIEJO
    emisora_vector_viejo = vector_viejo[vector_viejo['EMISORA'] == emisora]
    reducir_por_serie_viejo = emisora_vector_viejo[emisora_vector_viejo['SERIE'] == serie]
    reducir_por_serie_pip = emisora_vector[emisora_vector['SERIE'] == serie]
    
    ## EXTRAEMOS DATOS DE ARCHIVO INDEVAL GRANDE
    tasa_operacion = indeval_chico['Rendimiento'].loc[i]
    sobretasa_operacion = indeval_chico['Sobretasa'].loc[i]
    tasa_T_1 = reducir_por_serie_viejo['TASA DE RENDIMIENTO'].iloc[0]
    sobretasa_T_1 = reducir_por_serie_viejo['SOBRETASA'].iloc[0]
    
    if mdys == "AAA":
        calificacion = mdys
    else:
        calificacion = s_and_p
    
    fila = [nombre_emisora, calificacion, monto, diferencia_dias, tasa_operacion, 
            sobretasa_operacion, tasa_T, sobretasa_T, tasa_T_1, sobretasa_T_1,
            tasa_T_pip,sobretasa_T_pip,tasa_T_pip_viejo,sobretasa_T_pip_viejo]
    
    ## Añadimos fila a DF
    df_final_chico.loc[i] = fila

df_final_chico.to_excel(os.path.join(excel_output_dir, 'indeval_chico.xlsx'), index=False)