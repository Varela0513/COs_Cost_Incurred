import os
import pandas as pd
from tkinter import Tk, filedialog

# Función para obtener el directorio de datos usando Tkinter
def get_data_directory():
    root = Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Abrir el cuadro de diálogo para seleccionar el directorio
    data_directory = filedialog.askdirectory(title="Seleccionar directorio de datos")
    return data_directory

# Función para obtener el archivo de códigos de trabajo usando Tkinter
def get_job_codes_file():
    root = Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Abrir el cuadro de diálogo para seleccionar el archivo de códigos de trabajo
    job_codes_file = (
        filedialog.askopenfilename(title="Seleccionar archivo de códigos de trabajo", filetypes=[("Archivos de Excel", "*.xlsx")]))
    return job_codes_file

# Función para obtener la ruta del archivo de salida usando Tkinter
def get_output_file_path():
    root = Tk()
    root.withdraw()  # Ocultar la ventana principal

    # Abrir el cuadro de diálogo para seleccionar la ruta del archivo de salida
    local_output_file_path = filedialog.asksaveasfilename(
        title="Guardar resultado como", filetypes=[("Archivos de Excel", "*.xlsx")], defaultextension=".xlsx")
    return local_output_file_path

# Obtener directorio de datos, archivo de códigos de trabajo y ruta del archivo de salida
data_path = get_data_directory()
job_codes_file_path = get_job_codes_file()
output_file_path = get_output_file_path()

# Lista para almacenar DataFrames para cada archivo
dataframes = []

# Lista de columnas a eliminar
columns_to_remove = [
    "Estimated cost",
    "Projected cost",
    "Last cost",
    "% complete - cost"
]

# Iterar sobre cada archivo en el directorio de datos
for file in os.listdir(data_path):
    if file.endswith(".xlsx"):
        # Ruta completa del archivo
        file_path = os.path.join(data_path, file)

        # Leer el archivo de Excel y eliminar columnas no deseadas
        df = pd.read_excel(file_path)
        df = df.drop(columns=columns_to_remove, errors='ignore')

        # Filtrar filas que comienzan con 'C' en la columna 'Cost type'
        df = df[df['Cost type'].str.startswith('C', na=False)]

        # Ordenar el DataFrame por la columna 'Cost type'
        df = df.sort_values(by='Cost type')

        # Eliminar duplicados basados en todas las columnas
        df = df.drop_duplicates()

        # Agregar el DataFrame a la lista
        dataframes.append(df)

# Concatenar todos los DataFrames en uno solo, ordenando por las columnas 'Job' y 'Open commitments'
result = pd.concat(dataframes, axis=0, ignore_index=True, sort=False)

# Convertir la columna 'Job' a formato numérico
result['Job'] = pd.to_numeric(result['Job'], errors='coerce')

# Ordenar el DataFrame por las columnas 'Job' y 'Open commitments'
result = result.sort_values(by=['Job', 'Open commitments'])

# Leer el archivo de códigos de trabajo
job_codes_df = pd.read_excel(job_codes_file_path)

# Fusionar los DataFrames basándose en la columna 'Job'
result = pd.merge(result, job_codes_df, on='Job', how='left')

# Crear la nueva columna 'Cost Incurred' y agregar una marca de verificación si la suma es mayor que 0
result['Cost Incurred'] = (result['JTD cost'] + result['Open commitments'] > 0)

# Eliminar filas donde 'Cost Incurred' es Falso
result = result[result['Cost Incurred']]

# Reorganizar las columnas
column_order = (['Job', 'Project Name', 'Cost Incurred'] +
                [col for col in result.columns if col not in ['Job', 'Project Name', 'Cost Incurred']])
result = result[column_order]

# Guardar cada trabajo en una hoja separada en el archivo de Excel
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    for job, df_job in result.groupby('Job'):
        df_job.to_excel(writer, sheet_name=f'Job_{int(job)}', index=False)

        # Ajustar automáticamente el ancho de las columnas
        for i, col in enumerate(df_job.columns):
            max_len = max(df_job[col].astype(str).apply(len).max(), len(col))
            writer.sheets[f'Job_{int(job)}'].set_column(i, i, max_len + 2)
