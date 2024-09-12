import pandas as pd

# Ruta al archivo de Excel
file_path = '/juancise/Documentos/PYTHON_1/Origen_excel/Reporte de Asistencias2.xls'

# Lee el archivo de Excel
df = pd.read_excel(file_path)

# Elimina las primeras 11 filas y filas y columnas completamente vacías
df = df.drop(index=range(11)).dropna(how='all').dropna(axis=1, how='all')

# Renombra las columnas para mayor claridad
df.columns = ['Dia', 'Fecha', 'Nombre', 'Horario', 'Marca1', 'Marca2', 'Marca3', 'Marca4', 'DNI', 'Marca5']

# Formatea la columna 'Fecha' para mostrar solo la fecha
df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%d')

# Elimina las columnas no necesarias
df = df.drop(columns=['Dia', 'Nombre', 'Horario', 'Marca5'])

# Agrega las columnas vacías 'IdMarca' y 'IdHorario'
df['IdMarca'] = ''
df['IdHorario'] = ''

# Reorganiza las columnas en el orden deseado
df = df[['IdMarca', 'Marca1', 'Marca2', 'Marca3', 'Marca4', 'DNI', 'IdHorario', 'Fecha']]

# Crear y rellenar la columna 'Id_Dni'
df['Id_Dni'] = df['DNI'].where(df['DNI'].str.startswith('DNI:')).ffill().str.replace('DNI: ', '')

# Elimina la columna original 'DNI'
df = df.drop(columns=['DNI'])

# Mueve la columna 'Id_Dni' a la izquierda de 'IdHorario'
cols = df.columns.tolist()
cols.insert(cols.index('IdHorario'), cols.pop(cols.index('Id_Dni')))
df = df[cols]

# Elimina filas donde todas las columnas de marcas de horario están vacías
df = df.dropna(subset=['Marca1', 'Marca2', 'Marca3', 'Marca4'], how='all')

# Muestra los datos limpios
print(df)

# Guarda el DataFrame limpio en un nuevo archivo de Excel o CSV
output_path = '/juancise/Documentos/PYTHON_1/Resultados_excel/cleaned_schedule14.csv'
df.to_csv(output_path, index=False, sep=';', encoding='utf-8')

print(f"Datos guardados en: {output_path}")