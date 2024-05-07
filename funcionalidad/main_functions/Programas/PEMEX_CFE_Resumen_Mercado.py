import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
data_dir = os.path.join(os.path.dirname(script_dir), 'Data_a_Extraer')
excel_output_dir = os.path.join(os.path.dirname(script_dir), 'Exceles')
os.makedirs(excel_output_dir, exist_ok=True)

ruta_archivo = os.path.join(data_dir, 'VectorAnalitico24h.xls')
ruta_archivo_viejo = os.path.join(data_dir, 'VectorViejo.xls')

df_original = pd.read_excel(ruta_archivo)
df_viejo = pd.read_excel(ruta_archivo_viejo)

lista_tipo_valor = ['19','13-2','14-2','12U','14U','15U']


# Filtrar el DataFrame por los valores "CC", "CD" y "CP" en la columna "TIPO VALOR"
nuevo_df = df_original.loc[df_original['SERIE'].isin(lista_tipo_valor)]
nuevo_df = df_original.loc[df_original['EMISORA'].isin(['PEMEX'])]
nuevo_df = nuevo_df.loc[nuevo_df['TIPO VALOR'].isin(['95'])]

columnas = ['Serie','TASA','ST']
df_final_pemex = nuevo_df[['SERIE','TASA CUPON', 'SOBRETASA']]
df_final_pemex.columns = columnas

df_pemex_pip = pd.read_excel('Data_a_Extraer/PIP.xls',skiprows=1)
print(df_pemex_pip.columns)
df_pemex_pip = df_pemex_pip.loc[df_pemex_pip['SERIE'].isin(lista_tipo_valor)]
df_pemex_pip = df_pemex_pip.loc[df_pemex_pip['EMISORA'].isin(['PEMEX'])]
df_pemex_pip = df_pemex_pip.loc[df_pemex_pip['TIPO VALOR'].isin(['95'])]

df_pemex_pip = df_pemex_pip[['CUPON ACTUAL','SOBRETASA']]
df_pemex_pip.columns = columnas[1:]

tasas_pip = df_pemex_pip['TASA'].tolist()
sobretasas_pip = df_pemex_pip['ST'].tolist()

df_final_pemex.loc[:, 'TASA PIP'] = tasas_pip
df_final_pemex.loc[:, 'ST PIP'] = sobretasas_pip

df_final_pemex.to_excel(os.path.join(excel_output_dir, 'Pemex_Series.xlsx'), index=False)

#### EMPEZAMOS CON CFE

valores_cfe = ['14-2','17', '20-2', '21-5', '15U', '17U', '20U', '21U', '21-2U', '22-2S', '222UV', '22UV']

# Filtrar el DataFrame por los valores "CC", "CD" y "CP" en la columna "TIPO VALOR"
nuevo_df_cfe = df_original.loc[df_original['SERIE'].isin(valores_cfe)]
nuevo_df_cfe = nuevo_df_cfe.loc[df_original['EMISORA'].isin(['CFE'])]
nuevo_df_cfe = nuevo_df_cfe.loc[nuevo_df_cfe['TIPO VALOR'].isin(['95'])]

columnas = ['Serie','TASA','ST']
df_final_cfe = nuevo_df_cfe[['SERIE','TASA CUPON', 'SOBRETASA']]
df_final_cfe.columns = columnas

print(df_final_cfe)

df_cfe_pip = pd.read_excel('Data_a_Extraer/PIP.xls',skiprows=1)

df_cfe_pip = df_cfe_pip.loc[df_cfe_pip['EMISORA'].isin(['CFE'])]
df_cfe_pip = df_cfe_pip.loc[df_cfe_pip['TIPO VALOR'].isin(['95'])]
df_cfe_pip = df_cfe_pip.loc[df_cfe_pip['SERIE'].isin(valores_cfe)]

df_cfe_pip = df_cfe_pip[['CUPON ACTUAL','SOBRETASA']]
df_cfe_pip.columns = columnas[1:]

tasas_pip_cfe = df_cfe_pip['TASA'].tolist()
sobretasas_pip_cfe = df_cfe_pip['ST'].tolist()


df_final_cfe.loc[:, 'TASA PIP'] = tasas_pip_cfe
df_final_cfe.loc[:, 'ST PIP'] = sobretasas_pip_cfe

#################################################################
#### EMPEZAMOS CON PEMEX T-1 ####################################
#################################################################

pemex_t_1 = df_viejo.loc[df_viejo['SERIE'].isin(lista_tipo_valor)]
pemex_t_1 = pemex_t_1.loc[pemex_t_1['EMISORA'].isin(['PEMEX'])]
pemex_t_1 = pemex_t_1.loc[pemex_t_1['TIPO VALOR'].isin(['95'])]

columnas = ['Serie','TASA','V ST T-1']
pemex_t_1 = pemex_t_1[['SERIE','TASA CUPON', 'SOBRETASA']]
pemex_t_1.columns = columnas

pemex_valores_v_t_1 = pemex_t_1['V ST T-1'].tolist()

pemex_pip_t_1 = pd.read_excel('Data_a_Extraer/PIPViejo.xls',skiprows=1)
pemex_pip_t_1 = pemex_pip_t_1.loc[pemex_pip_t_1['SERIE'].isin(lista_tipo_valor)]
pemex_pip_t_1 = pemex_pip_t_1.loc[pemex_pip_t_1['EMISORA'].isin(['PEMEX'])]
df_pemex_pip = pemex_pip_t_1.loc[pemex_pip_t_1['TIPO VALOR'].isin(['95'])]

pemex_pip_t_1 = pemex_pip_t_1[['CUPON ACTUAL','SOBRETASA']]
pemex_pip_t_1.columns = ['CUPON ACTUAL','P ST T-1']

pemex_valores_p_t_1 = pemex_pip_t_1['P ST T-1'].tolist()

df_final_pemex.loc[:, 'V ST T-1'] = pemex_valores_v_t_1
df_final_pemex.loc[:, 'P ST T-1'] = pemex_valores_p_t_1

### EMPEZAMOS CFE T-1
cfe_t_1 = df_viejo.loc[df_viejo['SERIE'].isin(valores_cfe)]
cfe_t_1 = cfe_t_1.loc[df_viejo['EMISORA'].isin(['CFE'])]
cfe_t_1 = cfe_t_1.loc[cfe_t_1['TIPO VALOR'].isin(['95'])]

columnas = ['Serie','TASA','V ST T-1']
cfe_t_1 = cfe_t_1[['SERIE','TASA CUPON', 'SOBRETASA']]
cfe_t_1.columns = columnas


cfe_pip_t_1 = pd.read_excel('Data_a_Extraer/PIPViejo.xls',skiprows=1)

cfe_pip_t_1 = cfe_pip_t_1.loc[cfe_pip_t_1['EMISORA'].isin(['CFE'])]
cfe_pip_t_1 = cfe_pip_t_1.loc[cfe_pip_t_1['TIPO VALOR'].isin(['95'])]
cfe_pip_t_1 = cfe_pip_t_1.loc[cfe_pip_t_1['SERIE'].isin(valores_cfe)]

cfe_pip_t_1 = cfe_pip_t_1[['CUPON ACTUAL','SOBRETASA']]
df_cfe_pip.columns = ['CUPON ACTUAL','P ST T-1']

df_final_cfe.loc[:, 'V ST T-1'] = cfe_t_1['V ST T-1'].tolist()
df_final_cfe.loc[:, 'P ST T-1'] = df_cfe_pip['P ST T-1'].tolist()

df_final_pemex.to_excel(os.path.join(excel_output_dir, 'Pemex_Series.xlsx'), index=False)

df_final_cfe.to_excel(os.path.join(excel_output_dir, 'CFE_Series.xlsx'), index=False)
