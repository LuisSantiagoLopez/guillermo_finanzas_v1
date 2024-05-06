"""
FORMATTEA RESUMEN DE MERCADO MEMO 
"""

import pandas as pd

from datetime import datetime
import locale

### IMPORTAR FECHA 
# Intenta configurar el locale a español
try:
    locale.setlocale(locale.LC_TIME, 'es_ES')  # Para Windows usa 'spanish_Spain'
except locale.Error:
    print("Locale español no disponible, el resultado será en el idioma predeterminado.")

# Obtener la fecha actual
fecha_actual = datetime.now()

# Formatear la fecha en el formato deseado
texto_fecha_espanol = fecha_actual.strftime("%A, %d de %B de %Y")


# Leer los DataFrames
df_cetes = pd.read_excel("Exceles/CETE's.xlsx")
df_udibonos = pd.read_excel("Exceles/Tasas Reales (UDIBONO's).xlsx")
df_basic_swap = pd.read_excel("Exceles/Basis Swap TIIE SOFR.xlsx")
df_irs_tiie = pd.read_excel("Exceles/IRS TIIE.xlsx")
df_bonos_m = pd.read_excel("Exceles/Bonos de Tasa Fija (BONOS M).xlsx")
df_tipo_de_cambio = pd.read_excel("Exceles/Tipos de Cambio.xlsx")
df_cbics = pd.read_excel("Exceles/Tasas Reales (CBIC's).xlsx")
df_tiie_banxico = pd.read_excel("Exceles/TIIE Banxico.xlsx")
df_fondeo = pd.read_excel("Exceles/Tasa de Fondeo (Banxico) Jornada del 20240401.xlsx")
df_puntos_forward = pd.read_excel("Exceles/Puntos Forward.xlsx")
df_udi_tiie = pd.read_excel("Exceles/UDI-TIIE.xlsx")
df_PEMEX_series = pd.read_excel("Exceles/Pemex_Series.xlsx")
df_CFE_series = pd.read_excel("Exceles/CFE_Series.xlsx")

# Crear un Pandas Excel writer utilizando XlsxWriter como motor
### CREAS OBJETO DE WRITER CON ARCHIVO DE SALIDA
writer = pd.ExcelWriter("Output/Resumen_Mercado_Memo.xlsx", engine="xlsxwriter")

# INSERTAMOS DF'S A EXCEL
## DÓNDE QUIERES PONER EL DF 
df_cetes.to_excel(writer, sheet_name="Sheet1", startcol=1, startrow=4, index=False,header=False,)
df_udibonos.to_excel(writer, sheet_name="Sheet1", startcol=7, startrow=4, index=False, header=False)
df_irs_tiie.to_excel(writer, sheet_name="Sheet1", startcol=14, startrow=4, index=False,header=False)
df_bonos_m.to_excel(writer, sheet_name="Sheet1", startcol=1, startrow=int(df_cetes.shape[0]+6), index=False,header=False)
df_tipo_de_cambio.to_excel(writer, sheet_name="Sheet1", startcol=7, startrow=int(df_udibonos.shape[0]+5), index=False,header=False)
df_basic_swap.to_excel(writer, sheet_name="Sheet1", startcol=14, startrow=int(df_irs_tiie.shape[0]+6), index=False,header=False)
df_cbics.to_excel(writer, sheet_name="Sheet1", startcol=1, startrow=int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+3), index=False,header=False)
df_tiie_banxico.to_excel(writer, sheet_name="Sheet1", startcol=1, startrow=int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+3)+int(df_cbics.shape[0]+2), index=False,header=False)
df_fondeo.to_excel(writer, sheet_name="Sheet1", startcol=7, startrow=int(df_udibonos.shape[0]+5)+int(df_tipo_de_cambio.shape[0]+1), index=False,header=False)
df_puntos_forward.to_excel(writer, sheet_name="Sheet1", startcol=7, startrow=int(df_udibonos.shape[0]+5)+int(df_tipo_de_cambio.shape[0]+1)+int(df_fondeo.shape[0]+1), index=False,header=False)
df_udi_tiie.to_excel(writer, sheet_name="Sheet1", startcol=14, startrow=int(df_irs_tiie.shape[0]+6)+int(df_basic_swap.shape[0]), index=False,header=False)
df_PEMEX_series.to_excel(writer, sheet_name="Sheet1", startcol=21, startrow=6, index=False,header=False)
df_CFE_series.to_excel(writer, sheet_name="Sheet1", startcol=21, startrow=df_PEMEX_series.shape[0]+10, index=False,header=False)

# Obtener el objeto workbook y la hoja de trabajo
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

# ESTILOS
cell_format = workbook.add_format({
    'border': 0,  # Eliminar bordes
    'bg_color': '#FFFFFF',  # Fondo blanco
    'font_size': 8
})


center_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'border': 1,  # Establecer grosor del borde
    'border_color': '#000000',  # Establecer color del borde a negro
    'font_size': 8

})

#AJUSTE TAMAÑO COLUMNAS 
def adjust_column_width(worksheet, dataframe, startcol, header_format):
    for col_num, column_title in enumerate(dataframe.columns, start=startcol):
        # Calcular el ancho del encabezado
        if list(dataframe.columns)[0] == 'Contrato':
            column_width = max(len(column_title), max(dataframe[column_title].astype(str).map(len).max(), 7)) 
        else:
            column_width = max(len(column_title), max(dataframe[column_title].astype(str).map(len).max(), 12)) 
        # Ajustar el ancho de la columna
        worksheet.set_column(col_num, col_num, column_width, header_format)

#FORMATO BLANCO 
# Aplicar el formato a toda la hoja
worksheet.set_column('A:AG', None, cell_format)  # Aplica el formato desde la columna A hasta la Z

### AJUSTAMOS ANCHO DE COLUMNAS
adjust_column_width(worksheet, df_cetes, 1, cell_format)
adjust_column_width(worksheet, df_udibonos, 7, cell_format)
adjust_column_width(worksheet, df_irs_tiie, 14, cell_format)
adjust_column_width(worksheet, df_bonos_m, 1, cell_format)  # Asegúrate de ajustar las posiciones correctamente
adjust_column_width(worksheet, df_tipo_de_cambio, 7, cell_format)  # Asegúrate de ajustar las posiciones correctamente
adjust_column_width(worksheet, df_basic_swap, 14, cell_format)  # Asegúrate de ajustar las posiciones correctamente
adjust_column_width(worksheet, df_cbics, 1, cell_format)  # Asegúrate de ajustar las posiciones correctamente
adjust_column_width(worksheet, df_tiie_banxico, 1, cell_format)  # Asegúrate de ajustar las posiciones correctamente
adjust_column_width(worksheet, df_tipo_de_cambio, 7, cell_format)  # Asegúrate de ajustar las posiciones correctamente
adjust_column_width(worksheet, df_fondeo, 7, cell_format)  # Asegúrate de ajustar las posiciones correctamente
adjust_column_width(worksheet, df_puntos_forward, 7, cell_format)  # Asegúrate de ajustar las posiciones correctamente

# Función para aplicar formato de centrado a las celdas de los DataFrames
def apply_center_format(dataframe, worksheet, startrow, startcol, center_format):
    endrow = startrow + len(dataframe)
    endcol = startcol + len(dataframe.columns) - 1
    # Iterar sobre las celdas del DataFrame y aplicar el formato
    for row in range(startrow, endrow):
        for col in range(startcol, endcol + 1):
            # Aplicar el formato de centrado
            worksheet.write(row, col, dataframe.iloc[row-startrow, col-startcol], center_format)


# Aplicar el formato de centrado a las celdas de cada DataFrame insertado
apply_center_format(df_cetes, worksheet, 4, 1, center_format)
apply_center_format(df_udibonos, worksheet, 4, 7, center_format)
apply_center_format(df_irs_tiie, worksheet, 4, 14, center_format)
apply_center_format(df_bonos_m, worksheet, int(df_cetes.shape[0]+6), 1, center_format)
apply_center_format(df_tipo_de_cambio, worksheet, int(df_udibonos.shape[0]+5), 7, center_format)
apply_center_format(df_basic_swap, worksheet, int(df_irs_tiie.shape[0]+6), 14, center_format)
apply_center_format(df_cbics, worksheet, int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+3), 1, center_format)
apply_center_format(df_tiie_banxico, worksheet, int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+3)+int(df_cbics.shape[0]+2), 1, center_format)
apply_center_format(df_fondeo, worksheet, int(df_udibonos.shape[0]+5)+int(df_tipo_de_cambio.shape[0]+1), 7, center_format)
apply_center_format(df_puntos_forward, worksheet, int(df_udibonos.shape[0]+5)+int(df_tipo_de_cambio.shape[0]+1)+int(df_fondeo.shape[0]+1), 7, center_format)
apply_center_format(df_udi_tiie, worksheet, int(df_irs_tiie.shape[0]+6)+int(df_irs_tiie.shape[0]+1), 14, center_format)
apply_center_format(df_PEMEX_series, worksheet, 6, 21, center_format)
apply_center_format(df_CFE_series, worksheet, df_PEMEX_series.shape[0]+10, 21, center_format)


### ESTILOS ENCABEZADOS NEGROS 
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#000000',  # Color de fondo
    'border': 1,  # Establecer borde
    'align': 'center',  # Centrar el texto horizontalmente
    'valign': 'vcenter',  # Centrar el texto verticalmente
    'font_color':'#FFFFFF',
    'font_size': 8
})

#FORMATO ENCABEZADO AZULES 
# Configurar el formato para "Cetes"
TEXTOS = workbook.add_format({
    'bold': True,
    'font_color': '#0070c0',  # Color de texto azul claro
    'align': 'center',  # Centrar el texto horizontalmente
    'valign': 'vcenter',  # Centrar el texto verticalmente
    'font_size': 18,
})

#TEXTO CHIQUITO 
SUB_TEXTOS = workbook.add_format({
    'bold': True,
    'font_color': '#0070c0',  # Color de texto azul claro
    'align': 'center',  # Centrar el texto horizontalmente
    'valign': 'vcenter',  # Centrar el texto verticalmente
    'font_size': 12
})

 

# JUNTAMOS COLUMNAS PARA CIERTOS DF'S. PARA LAS QUE TIENEN QUE SER 2 COLUMNAS
valores_columna_L = df_udibonos['Ptos.'].values 
for row in range(4, 14):  
    valor = valores_columna_L[row-4]  
    worksheet.merge_range(row, 11, row, 12, valor, center_format)

valores_tiie_banxico = df_tiie_banxico['Dif.'].values  
for row in range(40,43): 
    worksheet.merge_range(row, 4, row, 5, valor, center_format)

# # AGREGAMOS ENCABEZADOS DE CADA TABLA AL WORKSHEET 
worksheet.merge_range(2, 1, 0, 5, "RESUMEN DE CIERRE DE MERCADO", TEXTOS)
worksheet.merge_range(2, 7, 0, 12, f"{texto_fecha_espanol}", TEXTOS)
worksheet.merge_range(3, 1, 3, 5, "CETES", SUB_TEXTOS)
worksheet.merge_range(3, 7, 3, 12, "UDIBONOS", SUB_TEXTOS)
worksheet.merge_range(2, 14, 0, 19, "RESUMEN DE CIERRE DE MERCADO", TEXTOS)
worksheet.merge_range(3, 14, 3, 19, "IRS TIIE", SUB_TEXTOS)
worksheet.merge_range(int(df_cetes.shape[0]+4), 1, int(df_cetes.shape[0]+4), 5, "BONOS M", SUB_TEXTOS)
worksheet.merge_range(int(df_udibonos.shape[0]+4), 7, int(df_udibonos.shape[0]+4), 12, "TIPO DE CAMBIO", SUB_TEXTOS)
worksheet.merge_range(int(df_irs_tiie.shape[0]+4), 14, int(df_irs_tiie.shape[0]+4), 19, "BASIS SWAP", SUB_TEXTOS)

### PONEMOS TITULOS PARA DF DE PEMEX 
worksheet.merge_range(2, 21, 0, 27, texto_fecha_espanol, TEXTOS)
worksheet.merge_range(3, 21, 3, 27, "PEMEX", SUB_TEXTOS)
worksheet.merge_range(4, 22, 4, 23, "VALMER", header_format)
worksheet.merge_range(4, 24, 4, 25, "PIP", header_format)
worksheet.merge_range(4, 26, 4, 27, "SOBRETASA T-1", header_format)

### PONEMOS TITULOS PARA DF DE CFE 
worksheet.merge_range(df_PEMEX_series.shape[0]+7, 21, df_PEMEX_series.shape[0]+7, 27, "CFE", SUB_TEXTOS)
worksheet.merge_range(df_PEMEX_series.shape[0]+8, 22, df_PEMEX_series.shape[0]+8, 23, "VALMER", header_format)
worksheet.merge_range(df_PEMEX_series.shape[0]+8, 24, df_PEMEX_series.shape[0]+8, 25, "PIP", header_format)
worksheet.merge_range(df_PEMEX_series.shape[0]+8, 26, df_PEMEX_series.shape[0]+8, 27, "SOBRETASA T-1", header_format)

worksheet.merge_range(int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+1), 1, int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+1), 5, "CBIC'S", SUB_TEXTOS)
worksheet.merge_range(int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+7), 1, int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+7), 5, "TIIE BANXICO", SUB_TEXTOS)
worksheet.merge_range(int(df_udibonos.shape[0]+4)+int(df_tipo_de_cambio.shape[0]), 7, int(df_udibonos.shape[0]+4)+int(df_tipo_de_cambio.shape[0]), 12, "FONDEO", SUB_TEXTOS)
worksheet.merge_range(int(df_udibonos.shape[0]+4)+int(df_tipo_de_cambio.shape[0])+int(df_fondeo.shape[0]+1), 7, int(df_udibonos.shape[0]+4)+int(df_tipo_de_cambio.shape[0])+int(df_fondeo.shape[0]+1), 12, "PUNTOS FORWARD", SUB_TEXTOS)
worksheet.merge_range(int(df_irs_tiie.shape[0]+4)+int(df_irs_tiie.shape[0]+1), 14, int(df_irs_tiie.shape[0]+4)+int(df_irs_tiie.shape[0]+1), 19, "UDI-TIIE", SUB_TEXTOS)

# NÚMEROS ROJOS 
# Aplicamos formato condicional a columnas seleccionadas
lista_columnas_condicional = ['F','L','M','T']
formato_negativo = workbook.add_format({'font_color': '#FF0000'})  # Rojo para negativos
for letra in lista_columnas_condicional:
    worksheet.conditional_format(f'{letra}5:{letra}104', {'type': 'cell',
                                            'criteria': '<',
                                            'value': 0,
                                            'format': formato_negativo})

#PARA CADA TABLA APLICAR FORMATO DE ENCABEZADO NEGRO Y BOLD 
# Aplicar el formato de encabezado a los encabezados de cada DataFrame
for col_num, value in enumerate(df_cetes.columns.values):
    worksheet.write(4, col_num + 1, value, header_format)
for col_num, value in enumerate(df_udibonos.columns.values):
    worksheet.write(4, col_num + 7, value, header_format)
for col_num, value in enumerate(df_irs_tiie.columns.values):
    worksheet.write(4, col_num + 14, value, header_format)
for col_num, value in enumerate(df_bonos_m.columns.values):
    worksheet.write(int(df_cetes.shape[0]+5), col_num + 1, value, header_format)
for col_num, value in enumerate(df_tipo_de_cambio.columns.values):
    worksheet.write(int(df_udibonos.shape[0]+5), col_num + 7, value, header_format)
for col_num, value in enumerate(df_basic_swap.columns.values):
    worksheet.write(int(df_irs_tiie.shape[0]+5), col_num + 14, value, header_format)
for col_num, value in enumerate(df_cbics.columns.values):
    worksheet.write(int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+2), col_num + 1, value, header_format)
for col_num, value in enumerate(df_tiie_banxico.columns.values):
    worksheet.write(int(df_cetes.shape[0]+6)+int(df_bonos_m.shape[0]+3)+int(df_cbics.shape[0]+1), col_num + 1, value, header_format)
for col_num, value in enumerate(df_tipo_de_cambio.columns.values):
    worksheet.write(int(df_udibonos.shape[0]+5)+int(df_tipo_de_cambio.shape[0]), col_num + 7, value, header_format)
for col_num, value in enumerate(df_tipo_de_cambio.columns.values):
    worksheet.write(int(df_udibonos.shape[0]+5)+int(df_tipo_de_cambio.shape[0])+int(df_fondeo.shape[0]+1), col_num + 7, value, header_format)
for col_num, value in enumerate(df_udi_tiie.columns.values):
    worksheet.write(int(df_irs_tiie.shape[0]+5)+int(df_irs_tiie.shape[0]+1), col_num + 14, value, header_format)
for col_num, value in enumerate(df_PEMEX_series.columns.values):
    worksheet.write(5, col_num + 21, value, header_format)
for col_num, value in enumerate(df_CFE_series.columns.values):
    worksheet.write(df_PEMEX_series.shape[0]+9, col_num + 21, value, header_format)


# HACEMOS COLUMNAS MAS PEQUEÑAS Y LAS PINTAMOS
white_fill_format = workbook.add_format({'bg_color': '#FFFFFF'})
worksheet.set_column('A:A', 1, white_fill_format)
worksheet.set_column('G:G', 1, white_fill_format)
worksheet.set_column('N:N', 1, white_fill_format)
worksheet.set_column('U:U', 1, white_fill_format)

# Cerrar el writer
writer.close()
