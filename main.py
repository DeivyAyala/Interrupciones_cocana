import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
import matplotlib.pyplot as plt
import numpy as np

# Configuración de la API de Google Sheets
credentials_path = 'credentials.json'  # Ruta a tu archivo de credenciales
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
credentials = service_account.Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
service = build('sheets', 'v4', credentials=credentials)

# ID del archivo de Google Sheets
file_id = '1zftCB8k2NtfkhkjazBc8eoIUlh1YU0XMzdj06JlgUKo'
range_name = 'Respuestas de formulario 1'

# Función para obtener los datos de Google Sheets
def get_sheet_data(file_id, range_name):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=file_id, range=range_name).execute()
    values = result.get('values', [])
    return values

# Obtener los datos completos del Google Sheets
sheet_data = get_sheet_data(file_id, range_name)

# Asegurarse de que todas las filas tengan el mismo número de columnas que el encabezado
max_columns = len(sheet_data[0])
for row in sheet_data:
    if len(row) < max_columns:
        row.extend([None] * (max_columns - len(row)))

# Convertir los datos a un DataFrame de pandas y filtrar filas vacías
df = pd.DataFrame(sheet_data[1:], columns=sheet_data[0])

# Convertir las columnas de fecha y hora a datetime sin especificar el formato
df['FECHA INICIO'] = pd.to_datetime(df['FECHA INICIO'] + ' ' + df['HORA INICIO'], errors='coerce')
df['FECHA FINAL'] = pd.to_datetime(df['FECHA FINAL'] + ' ' + df['HORA FIN'], errors='coerce')

# Calcular la duración de las interrupciones en minutos
df['DURACIÓN (MINUTOS)'] = (df['FECHA FINAL'] - df['FECHA INICIO']).dt.total_seconds() / 60

# Convertir la duración de cada interrupción a formato hh:mm:ss
df['DURACIÓN TOTAL (HH:MM:SS)'] = pd.to_timedelta(df['DURACIÓN (MINUTOS)'], unit='m').apply(lambda x: str(x) if pd.notnull(x) else '00:00:00')

# Crear la tabla de causas
causas = [
    'DESCARGAS ATMOSFÉRICAS',
    'EQUIPOS SUBESTACIÓN',
    'MANIOBRAS ENERGIZACIÓN',
    'MANTENIMIENTO',
    'VEGETACIÓN',
    'ORDEN PÚBLICO',
    'SOBRECORRIENTE'
]

# Contar frecuencias y duraciones totales por causa
causa_freq = df['CAUSA INTERRUPCIÓN '].value_counts().reindex(causas, fill_value=0)
causa_durations = df.groupby('CAUSA INTERRUPCIÓN ')['DURACIÓN (MINUTOS)'].sum().reindex(causas, fill_value=0)
causa_durations_hms = pd.to_timedelta(causa_durations, unit='m').apply(lambda x: str(x) if pd.notnull(x) else '00:00:00')

causas_df = pd.DataFrame({
    'Causa': causas,
    'Frecuencia': causa_freq.values,
    'Duración Total (HH:MM:SS)': causa_durations_hms.values
})

# Crear la tabla de fallas por circuito
circuitos = ['01GU03SB', '01GU02OH', '01GU', '02OH']

# Contar frecuencias y duraciones totales por circuito
circuito_freq = df['CÓDIGO DEL CIRCUITO'].value_counts().reindex(circuitos, fill_value=0)
circuito_durations = df.groupby('CÓDIGO DEL CIRCUITO')['DURACIÓN (MINUTOS)'].sum().reindex(circuitos, fill_value=0)
circuito_durations_hms = pd.to_timedelta(circuito_durations, unit='m').apply(lambda x: str(x) if pd.notnull(x) else '00:00:00')

circuitos_df = pd.DataFrame({
    'Código del Circuito': circuitos,
    'Frecuencia': circuito_freq.values,
    'Duración Total (HH:MM:SS)': circuito_durations_hms.values
})

# Convertir la palabra 'days' por 'días' solo para la presentación final
causas_df['Duración Total (HH:MM:SS)'] = causas_df['Duración Total (HH:MM:SS)'].str.replace('days', 'días')
circuitos_df['Duración Total (HH:MM:SS)'] = circuitos_df['Duración Total (HH:MM:SS)'].str.replace('days', 'días')

# Crear un archivo Excel con las tablas y la duración de las interrupciones
output_file_path = 'resultados/fallas_115kV.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    df.to_excel(writer, sheet_name='Interrupciones', index=False)
    causas_df.to_excel(writer, sheet_name='CAUSAS', index=False)
    circuitos_df.to_excel(writer, sheet_name='FALLAS POR CIRCUITO', index=False)

# Función para calcular el color basado en la duración
def calcular_color(duracion):
    if pd.isna(duracion):
        return '#ffffff'  # blanco
    minutos = duracion.total_seconds() / 60
    if minutos <= 60:
        return '#00ff00'  # verde
    elif minutos <= 180:
        return '#ffff00'  # amarillo
    elif minutos <= 360:
        return '#ffa500'  # naranja
    else:
        return '#ff0000'  # rojo

# Graficar frecuencia de interrupciones por causa
plt.figure(figsize=(10, 6))
colores = [calcular_color(pd.to_timedelta(d.replace('días', 'days'))) for d in causas_df['Duración Total (HH:MM:SS)']]
plt.bar(causas_df['Causa'], causas_df['Frecuencia'], color=colores)
plt.xlabel('Causa')
plt.ylabel('Frecuencia')
plt.title('Frecuencia de Interrupciones por Causa')
plt.xticks(rotation=45, ha='right')

# Añadir la duración encima de cada barra
for i in range(len(causas_df)):
    plt.text(i, causas_df['Frecuencia'][i], causas_df['Duración Total (HH:MM:SS)'][i], ha='center', va='bottom')

plt.tight_layout()
plt.savefig('resultados/causas_frecuencia_con_duracion.png')
plt.show()

# Graficar frecuencia de interrupciones por circuito
plt.figure(figsize=(10, 6))
colores = [calcular_color(pd.to_timedelta(d.replace('días', 'days'))) for d in circuitos_df['Duración Total (HH:MM:SS)']]
plt.bar(circuitos_df['Código del Circuito'], circuitos_df['Frecuencia'], color=colores)
plt.xlabel('Código del Circuito')
plt.ylabel('Frecuencia')
plt.title('Frecuencia de Interrupciones por Circuito')
plt.xticks(rotation=45, ha='right')

# Añadir la duración encima de cada barra
for i in range(len(circuitos_df)):
    plt.text(i, circuitos_df['Frecuencia'][i], circuitos_df['Duración Total (HH:MM:SS)'][i], ha='center', va='bottom')

plt.tight_layout()
plt.savefig('resultados/circuitos_frecuencia_con_duracion.png')
plt.show()

print(f"Archivo Excel '{output_file_path}' creado con éxito.")
print("Gráficos creados con éxito.")

