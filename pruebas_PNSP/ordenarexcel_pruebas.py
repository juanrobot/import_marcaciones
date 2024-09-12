import pandas as pd

# Ruta al archivo de Excel
file_path = '/Users/kivef/Downloads/Reporte de Asistencias (4).xls'

# Lee el archivo de Excel
df = pd.read_excel(file_path, skiprows=11)

print(df)
# Elimina las primeras 11 filas
# df = df.drop(index=range(11))

# Guarda el DataFrame limpio en un nuevo archivo de Excel o CSV
#output_path = 'F:\Otros ordenadores\Mi PC\PYTHON_1\Resultados_excel\Reporte_prueba88241319.xlsx'
#df.to_excel(output_path, index=False)

#print(f"Datos guardados en: {output_path}")

"""# Elimina filas completamente vacías
df.dropna(how='all', inplace=True)

# Elimina columnas completamente vacías
df.dropna(axis=1, how='all', inplace=True)

# Muestra los datos después de eliminar filas y columnas vacías
#print(df)

# Renombra las columnas para mayor claridad, si es necesario
# Ajusta el nombre de las columnas según corresponda
df.columns = ['Dia', 'Fecha', 'Nombre', 'Horario', 'Marca1','Marca2','Marca3','Marca4','Marca5','Marca6','DNI','Marca7','FISCALIZADO']

# Formatea la columna 'Date' para mostrar solo la fecha
df['Fecha'] = pd.to_datetime(df['Fecha']).dt.strftime('%Y-%m-%d')

# Eliminar las columnas especificadas
df = df.drop(columns=['Dia', 'Nombre', 'Horario', 'FISCALIZADO'])

# Agregar las columnas vacías
df['IdMarca'] = ''  # Columna vacía para IdMarca
df['IdHorario'] = ''  # Columna vacía para IdHorario

# Reorganizar las columnas en el orden deseado
df = df[['IdMarca', 'Marca1', 'Marca2', 'Marca3', 'Marca4', 'Marca5','Marca6','Marca7', 'DNI', 'IdHorario', 'Fecha']]

# Muestra los datos limpios
print(df)

# Guarda el DataFrame limpio en un nuevo archivo de Excel o CSV
output_path = '/Users/kivef/Documents/PYTHON_1/Resultados_excel/Reporte_prueba3.xlsx'
df.to_excel(output_path, index=False)

print(f"Datos guardados en: {output_path}")"""