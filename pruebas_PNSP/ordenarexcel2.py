import pandas as pd

# Ruta al archivo de Excel
file_path = '/juancise/Documentos/PYTHON_1/Origen_excel/Reporte de Asistencias2.xls'

# Lee el archivo de Excel
df = pd.read_excel(file_path)


# Elimina las primeras 11 filas
df = df.drop(index=range(11))

# Elimina filas completamente vacías
df.dropna(how='all', inplace=True)

# Elimina columnas completamente vacías
df.dropna(axis=1, how='all', inplace=True)

# Muestra los datos después de eliminar filas y columnas vacías
print(df)

# Renombra las columnas para mayor claridad, si es necesario
# Ajusta el nombre de las columnas según corresponda
df.columns = ['Dia', 'Fecha', 'Nombre', 'Horario', 'Marca1','Marca2','Marca3','Marca4','DNI','Marca5']

# Formatea la columna 'Date' para mostrar solo la fecha
df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%d')

# Muestra los datos limpios
print(df)

# Eliminar las columnas especificadas
df = df.drop(columns=['Dia', 'Nombre', 'Horario', 'Marca5'])

# Agregar las columnas vacías
df['IdMarca'] = ''  # Columna vacía para IdMarca
df['IdHorario'] = ''  # Columna vacía para IdHorario

# Reorganizar las columnas en el orden deseado
df = df[['IdMarca', 'Marca1', 'Marca2', 'Marca3', 'Marca4', 'DNI', 'IdHorario', 'Fecha']]

# Muestra los datos limpios
print(df)

# Crear una nueva columna Id_Dni y rellenarla con el DNI correspondiente
df['Id_Dni'] = ""

current_dni = ""
for index, row in df.iterrows():
    if isinstance(row['DNI'], str) and row['DNI'].startswith("DNI:"):
        current_dni = row['DNI'].replace("DNI: ", "")
    df.at[index, 'Id_Dni'] = current_dni

# Eliminar la columna original DNI
df.drop(columns=['DNI'], inplace=True)

# Mover la columna Id_Dni a la izquierda de IdHorario
cols = df.columns.tolist()
cols.insert(cols.index('IdHorario'), cols.pop(cols.index('Id_Dni')))
df = df[cols]

# Eliminar filas donde todas las columnas de marcas de horario están vacías
df = df.dropna(subset=['Marca1', 'Marca2', 'Marca3', 'Marca4'], how='all')

print(df)

# Guarda el DataFrame limpio en un nuevo archivo de Excel o CSV
output_path = '/juancise/Documentos/PYTHON_1/Resultados_excel/cleaned_schedule13.csv'
df.to_csv(output_path, index=False, sep=';', encoding='utf-8')

print(f"Datos guardados en: {output_path}")