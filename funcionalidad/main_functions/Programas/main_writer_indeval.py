import pandas as pd
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
excel_dir = os.path.join(os.path.dirname(script_dir), 'Exceles')
output_dir = os.path.join(os.path.dirname(script_dir), 'Output')
os.makedirs(output_dir, exist_ok=True)

## CARGAR DATOS 
df_grande = pd.read_excel(os.path.join(excel_dir, 'indeval_grande.xlsx'))
df_chico = pd.read_excel(os.path.join(excel_dir, 'indeval_chico.xlsx'))

df_fechas_grande = pd.read_excel('Exceles/Fechas_Indeval_Grande.xlsx')
df_fechas_chico = pd.read_excel('Exceles/Fechas_Indeval_Chico.xlsx')

## OBJECTO XLSXWRITER
writer = pd.ExcelWriter(os.path.join(output_dir, 'Styled_Indeval_Report.xlsx'), engine='xlsxwriter')
workbook = writer.book

## FORMATOS CUSTOMIZADOS PARA CADA SECCIÓN 

### FORMATO PARA ENCABEZADO VERDE 
header_green_format = workbook.add_format({
    'bg_color': '#A9D08E', # Verde claro 
    'color': 'white',
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 14,
    'bold': True,
    'border': 0
})

### FORMATO PARA ENCABEZADO AZUL CLARO DE LAS COLUMNAS DE LAS TABLAS, Y DEL TOTAL 
light_blue_column_header_format = workbook.add_format({
    'bg_color': '#00BFFF',  # Azul claro 
    'color': 'black',
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 11,
    'bold': True,
    'border': 1
})

### FORMATO PARA CELDAS AZUL OSCURO DE VALMER 
dark_blue_valmer_format = workbook.add_format({
    'bg_color': '#305496',  # Azul oscuro 
    'color': 'white',
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 11,
    'bold': True,
    'border': 1
})

### FORMATO PARA LAS CELDAS PIP 
dark_blue_pip_format = workbook.add_format({
    'bg_color': '#1F4E78',  # Azul oscuro (más oscuro)
    'color': 'white',
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 11,
    'bold': True,
    'border': 1
})

### FORMATO PARA LAS CELDAS PLAIN 
plain_white_format = workbook.add_format({
    'bg_color': '#FFFFFF',  # White color
    'color': 'black',
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 11,
    'border': 1
})

## FORMATO PARA ENCABEZADO "VALUACION"
valuacion_header_format = workbook.add_format({
    'color': '#1F4E78',
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Calibri',
    'font_size': 14,  
    'bold': True,
    'border': 0,
})

### FORMATO PARA EL BACKGROUND 
white_background_format = workbook.add_format({
    'bg_color': '#FFFFFF',  # Color blanco para el background 
    'border': 0,           # Sin bordes 
    'align': 'center',      # Center align
    'valign': 'vcenter',     # Center vertically
})

#### DATAFRAME GRANDE ####

## DEFINICIÓN DE POSICIÓN DE DATAFRAMES GRANDE Y CHICO EN EL EXCEL
df_grande_startrow = 8
df_chico_startrow = 33  # ESPACIO ENTRE TABLAS 
both_df_startcol = 3

## TRADUCCIÓN DE LOS DATAFRAMES GRANDES Y CHICOS AL EXCEL 
df_grande.to_excel(writer, sheet_name='Sheet1', startrow=df_grande_startrow, startcol=both_df_startcol, index=False, header=False)
df_chico.to_excel(writer, sheet_name='Sheet1', startrow=df_chico_startrow, startcol=both_df_startcol, index=False, header=False)

## DEFINICION DE POSICION DE DATAFRAMES DE FECHAS EN EL EXCEL
both_fechas_startcol = 1

## TRADUCCIÓN DE LOS DATAFRAMES DE FECHAS AL EXCEL
df_fechas_grande.to_excel(writer, sheet_name='Sheet1', startrow=df_grande_startrow, startcol=both_fechas_startcol, index=False, header=False)
df_fechas_chico.to_excel(writer, sheet_name='Sheet1', startrow=df_chico_startrow, startcol=both_fechas_startcol, index=False, header=False)

## OBTENCIÓN DE LA HOJA DE TRABAJO
worksheet = writer.sheets['Sheet1']

## HACER TODAS LAS CELDAS BLANCAS SIN BORDES ANTES DE CARGAR DF O APLICAR ESTILOS Y ESTABLECER ANCHO DE COLUMNAS
width_all_columns = 20
worksheet.set_column('A:XFD', width_all_columns, white_background_format)

## TÍTULO VERDE HASTA ARRIBA 
header_green_text = """TIPO DE VALOR "I" BANCA DE DESARROLLO"""													
worksheet.merge_range("D6:Q6", header_green_text, header_green_format)

## TÍTULO V 
dark_blue_valmer_title = "V"													
worksheet.merge_range("J7:M7", dark_blue_valmer_title, dark_blue_valmer_format)

## TÍTULO PIP
dark_blue_pip_title = "P"
worksheet.merge_range("N7:Q7", dark_blue_pip_title, dark_blue_pip_format)

## APLICACIÓN DE FORMATO AZUL PARA LAS COLUMNAS DE INDEVAL GRANDE E INDEVAL CHICO
# Convierte los headers de df_grande y df_chico en listas
headers_grande = df_grande.columns.tolist()
chico_headers = df_chico.columns.tolist()

# Aplica el color azul solamente a los headers de las primeras 7 columnas de df_grande
for col_num, header in enumerate(headers_grande):
    if col_num < 6:  # Solamente 7 columnas 
        # Escribe el header con formato, tomando en cuenta el indexado de Excel
        worksheet.write(df_grande_startrow - 1, both_df_startcol + col_num, header, light_blue_column_header_format)

# Aplica el color azul solamente a los headers de las primeras 7 columnas de df_chico
for col_num, header in enumerate(chico_headers):
    if col_num < 6:  # Solamente 7 columnas
        # Escribe el header con formato, tomando en cuenta el indexado de Excel
        worksheet.write(df_chico_startrow - 1, both_df_startcol + col_num, header, light_blue_column_header_format)

## APLICACIÓN DE FORMATO AZUL OSCURO PARA LAS COLUMNAS DE VALMER
# Define the start row for the headers and the range for the data cells
valmer_header_startrow = df_grande_startrow - 1
valmer_data_startrow = df_grande_startrow
valmer_data_endrow = valmer_data_startrow + len(df_grande) - 1

# Loop to apply dark blue format to headers from J8 to M8
for col_num in range(9, 13):  # Columns J to M
    worksheet.write(valmer_header_startrow, col_num, df_grande.columns[col_num - both_df_startcol], dark_blue_valmer_format)

# Loop to apply dark blue format to cells
for row in range(valmer_data_startrow, valmer_data_endrow + 1):  # Adjusted for dynamic data size
    for col in range(9, 13):  # Columns J to M
        data_value = df_grande.iloc[row - df_grande_startrow, col - both_df_startcol]
        worksheet.write(row, col, data_value, dark_blue_valmer_format)

## APLICACIÓN DE FORMATO PIP PARA LAS COLUMNAS BAJO EL TÍTULO "P"
pip_header_startrow = df_grande_startrow - 1
pip_data_startrow = df_grande_startrow
pip_data_endrow = pip_data_startrow + len(df_grande) - 1

# Loop to apply dark blue PIP format to headers from N8 to Q8
for col_num in range(13, 17):  # Columns N to Q
    worksheet.write(pip_header_startrow, col_num, df_grande.columns[col_num - both_df_startcol], dark_blue_pip_format)

# Loop to apply dark blue PIP format to cells
for row in range(pip_data_startrow, pip_data_endrow + 1):
    for col in range(13, 17):  # Columns N to Q
        data_value = df_grande.iloc[row - df_grande_startrow, col - both_df_startcol]
        worksheet.write(row, col, data_value, dark_blue_pip_format)

## APLICACIÓN DE FORMATO PARA CELDAS BLANCAS EN DF GRANDE
# Loop to apply plain white format dynamically based on the actual data bounds
for row in range(df_grande_startrow, df_grande_startrow + len(df_grande)):
    for col in range(3, 9):  # From column D to I
        data_value = df_grande.iloc[row - df_grande_startrow, col - both_df_startcol]
        worksheet.write(row, col, data_value, plain_white_format)

## APLICACIÓN DE CELDA TOTAL DF GRANDE
# Add "TOTAL" to the last row of column E
worksheet.write(df_grande_startrow + len(df_grande), 4, "TOTAL", light_blue_column_header_format)

# Add the sum of the column F to the same row
total_grande = df_grande.iloc[:, 2].sum()
worksheet.write(df_grande_startrow + len(df_grande), 5, total_grande, light_blue_column_header_format)

#### DATAFRAME CHICO ####

## TÍTULO "VALUACION" PARA EL DATAFRAME CHICO
valuacion_header_text = """VALUACION"""
worksheet.merge_range("J31:Q31", valuacion_header_text, valuacion_header_format)

## TÍTULO VALMER
dark_blue_valmer_title = "VALMER"
worksheet.merge_range("J32:M32", dark_blue_valmer_title, dark_blue_valmer_format)

## TÍTULO PIP
dark_blue_pip_title_chico = "PIP"
worksheet.merge_range("N32:Q32", dark_blue_pip_title_chico, dark_blue_pip_format)

## APLICACIÓN DE FORMATO AZUL OSCURO PARA LAS COLUMNAS DE VALMER EN DF CHICO
valmer_header_row_chico = df_chico_startrow - 1
valmer_data_start_row_chico = df_chico_startrow
valmer_data_end_row_chico = valmer_data_start_row_chico + len(df_chico) - 1

# Apply format to headers from J33 to M33
for col_num in range(9, 13):  # Columns J to M
    worksheet.write(valmer_header_row_chico, col_num, chico_headers[col_num - both_df_startcol], dark_blue_valmer_format)

# Apply format to cells from J34 onwards
for row in range(valmer_data_start_row_chico, valmer_data_end_row_chico + 1):
    for col in range(9, 13):
        data_value = df_chico.iloc[row - df_chico_startrow, col - both_df_startcol]
        worksheet.write(row, col, data_value, dark_blue_valmer_format)

## APLICACIÓN DE FORMATO AZUL OSCURO PARA LAS COLUMNAS DE PIP EN DF CHICO
pip_header_row_chico = df_chico_startrow - 1
pip_data_start_row_chico = df_chico_startrow
pip_data_end_row_chico = pip_data_start_row_chico + len(df_chico) - 1

# Apply format to headers from N33 to Q33
for col_num in range(13, 17):  # Columns N to Q
    worksheet.write(pip_header_row_chico, col_num, chico_headers[col_num - both_df_startcol], dark_blue_pip_format)

# Apply format to cells within df_chico's size from N34 onward
for row in range(pip_data_start_row_chico, pip_data_end_row_chico + 1):
    for col in range(13, 17):
        data_value = df_chico.iloc[row - df_chico_startrow, col - both_df_startcol]
        worksheet.write(row, col, data_value, dark_blue_pip_format)

## APLICACIÓN DE FORMATO PARA CELDAS BLANCAS EN DF CHICO
# Loop to apply plain white format to cells from D34 to I40
for row in range(df_chico_startrow, df_chico_startrow + 7):  # From D34 to D40
    for col in range(3, 9):  # From column D to I
        # Check if the cell index is within the actual data bounds of df_chico
        if row - df_chico_startrow < df_chico.shape[0]:
            data_value = df_chico.iloc[row - df_chico_startrow, col - both_df_startcol]
            worksheet.write(row, col, data_value, plain_white_format)
        else:
            worksheet.write(row, col, "", plain_white_format)  # Write empty string if out of data bounds

## APLICACIÓN DE CELDA TOTAL DF CHICO
# Add "TOTAL" to the last row of column E
worksheet.write(df_chico_startrow + len(df_chico), 4, "TOTAL", light_blue_column_header_format)

# Add the sum of the column F to the same row
total_chico = df_chico.iloc[:, 2].sum()
worksheet.write(df_chico_startrow + len(df_chico), 5, total_chico, light_blue_column_header_format)

### FECHAS ###

# Headers for `df_fechas_grande`
fechas_grande_header_start_row = df_grande_startrow - 1  # Move the header up by one row
worksheet.write(fechas_grande_header_start_row, both_fechas_startcol, df_fechas_grande.columns[0], light_blue_column_header_format)

# Apply white format to data in `df_fechas_grande`
for row in range(fechas_grande_header_start_row + 1, fechas_grande_header_start_row + len(df_fechas_grande) + 1):
    data_value = df_fechas_grande.iloc[row - fechas_grande_header_start_row - 1, 0]
    worksheet.write(row, both_fechas_startcol, data_value, plain_white_format)

# Headers for `df_fechas_chico`
fechas_chico_header_start_row = df_chico_startrow - 1  # Move the header up by one row
worksheet.write(fechas_chico_header_start_row, both_fechas_startcol, df_fechas_chico.columns[0], light_blue_column_header_format)

# Apply white format to data in `df_fechas_chico`
for row in range(fechas_chico_header_start_row + 1, fechas_chico_header_start_row + len(df_fechas_chico) + 1):
    data_value = df_fechas_chico.iloc[row - fechas_chico_header_start_row - 1, 0]
    worksheet.write(row, both_fechas_startcol, data_value, plain_white_format)

writer.close()