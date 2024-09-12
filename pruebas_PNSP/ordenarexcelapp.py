import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def process_file(input_path, output_path):
    try:
        # Lee el archivo de Excel
        df = pd.read_excel(input_path)

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

        # Guarda el DataFrame limpio en un nuevo archivo CSV
        df.to_csv(output_path, index=False, sep=';', encoding='utf-8')
        messagebox.showinfo("Éxito", f"Datos guardados en: {output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def select_file():
    input_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xls;*.xlsx")],
        title="Selecciona el archivo de Excel"
    )
    if input_path:
        output_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Selecciona la ubicación para guardar el archivo CSV"
        )
        if output_path:
            process_file(input_path, output_path)

# Crear la ventana principal
root = tk.Tk()
root.title("Procesador de Excel a CSV")

# Mensaje de autor
author_label = tk.Label(root, text="Desarrollado por: JCED", font=("Arial", 10))
author_label.pack(pady=10)

# Botón para seleccionar y procesar el archivo
btn = tk.Button(root, text="Seleccionar archivo y procesar", command=select_file)
btn.pack(pady=20)

# Ejecutar la interfaz gráfica
root.mainloop()