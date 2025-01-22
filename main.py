import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import unicodedata
from openpyxl import load_workbook

# Función para normalizar texto
def normalizar_texto(texto):
    texto = str(texto).strip()
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('utf-8')
    texto = " ".join(texto.split())
    return texto.lower()

# Función para cambiar al menú principal
def cambiar_a_menu():
    for frame in [frame_anexo]:
        frame.pack_forget()
    frame_menu.pack(fill="both", expand=True)

# Función para cargar el archivo Excel
def archivo_anexo():
    global data  # Declarar la variable global
    # Seleccionar archivo Excel
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        # Leer el archivo
        data = pd.read_excel(file_path, skiprows=4)

        # Limpiar nombres de columnas
        data.columns = data.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # Validar columnas obligatorias
        columnas_requeridas = [
            "Estado MMS", "SISTEMA", "TIPO", "NOMBRE DEL EQUIPO",
            "OTROS SISTEMAS", "EMPLAZAMIENTO", "LINEA", "TRATAMIENTO"
        ]
        columnas_faltantes = [col for col in columnas_requeridas if col not in data.columns]
        if columnas_faltantes:
            messagebox.showerror("Error", f"Columnas faltantes en el archivo: {', '.join(columnas_faltantes)}")
            return

                # Convertir la columna "Fecha" a formato datetime
        if "Fecha" in data.columns:
            data["Fecha"] = pd.to_datetime(data["Fecha"], errors="coerce")  # Convertir a datetime
            data = data.dropna(subset=["Fecha"])  # Eliminar filas con fechas inválidas

            # Extraer los años únicos
            anios_disponibles = data["Fecha"].dt.year.dropna().astype(int).unique().tolist()
            anios_disponibles.sort()

            # Actualizar el menú desplegable para los años
            if "Fecha" not in menus:
                label_fecha = tk.Label(frame_campos, text="Fecha (Año)", bg="#2b2b2b", fg="white", font=("Arial", 10))
                label_fecha.grid(row=0, column=0, padx=10, pady=5, sticky="e")
                menus["Fecha"] = tk.OptionMenu(frame_campos, variables["Fecha"], *anios_disponibles, command=lambda v: actualizar_seleccion("Fecha", v))
                menus["Fecha"].grid(row=0, column=1, padx=10, pady=5, sticky="w")
            else:
                menus["Fecha"].children["menu"].delete(0, "end")
                for anio in anios_disponibles:
                    menus["Fecha"].children["menu"].add_command(
                        label=anio,
                        command=lambda v=anio: actualizar_seleccion("Fecha", v)
                    )

    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")

        # Rellenar valores nulos con un texto predeterminado
        for columna in columnas_requeridas:
            data[columna] = data[columna].fillna("Desconocido").astype(str).str.strip()

        print("Archivo cargado y validado correctamente.")
        print(f"Total de filas: {len(data)}")

    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")

    # Actualizar semanas disponibles (si aplicable)
    if "Fecha" in data.columns:
        data["Semana"] = data["Fecha"].dt.isocalendar().week
        semanas_disponibles = data["Semana"].dropna().astype(int).unique().tolist()
        semanas_disponibles.sort()

        if "Semana" not in menus:
            label_semana = tk.Label(frame_campos, text="Semana", bg="#2b2b2b", fg="white", font=("Arial", 10))
            label_semana.grid(row=1, column=0, padx=10, pady=5, sticky="e")
            menus["Semana"] = tk.OptionMenu(frame_campos, variables["Semana"], *semanas_disponibles, command=lambda v: actualizar_seleccion("Semana", v))
            menus["Semana"].grid(row=1, column=1, padx=10, pady=5, sticky="w")
        else:
            menus["Semana"].children["menu"].delete(0, "end")
            for semana in semanas_disponibles:
                menus["Semana"].children["menu"].add_command(
                    label=semana,
                    command=lambda v=semana: actualizar_seleccion("Semana", v)
                )

    # Detectar y normalizar valores únicos de otras columnas
    columnas_procesar = [
        "TIPO SICE", "ESTADO SICE", "TIPO MMS", "Estado MMS",
        "SISTEMA", "TIPO", "NOMBRE DEL EQUIPO",
        "OTROS SISTEMAS", "EMPLAZAMIENTO", "LINEA", "TRATAMIENTO"
    ]

    for columna in columnas_procesar:
        if columna in data.columns:
            valores_originales = data[columna].dropna().unique().tolist()
            normalizados = {normalizar_texto(valor): valor for valor in valores_originales}
            valores_unicos = list(normalizados.values())
            valores_unicos.sort()

            if columna not in variables:
                variables[columna] = tk.StringVar(value="Seleccione")

            if columna not in menus:
                label = tk.Label(frame_campos, text=columna, bg="#2b2b2b", fg="white", font=("Arial", 10))
                label.grid(row=2, column=0, padx=10, pady=5, sticky="e")
                menus[columna] = tk.OptionMenu(frame_campos, variables[columna], *valores_unicos)
                menus[columna].grid(row=2, column=1, padx=10, pady=5, sticky="w")
            else:
                menus[columna].children["menu"].delete(0, "end")
                for valor in valores_unicos:
                    menus[columna].children["menu"].add_command(
                        label=valor,
                        command=lambda v=valor, col=columna: actualizar_seleccion(col, v)
                    )


# Función para actualizar las selecciones y previsualizarlas
def actualizar_seleccion(campo, valor):
    variables[campo].set(valor)

def exportar_seleccion():
    if data is None or data.empty:
        messagebox.showwarning("Advertencia", "No hay datos cargados para filtrar.")
        return

    # Recoger las selecciones de los campos
    seleccion = {campo: variables[campo].get() for campo in variables if variables[campo].get() != "Seleccione"}

    # Recoger palabras clave de los campos de descripción
    palabras_clave_descripcion = {campo: variables_independientes[campo].get().strip().lower()
                                   for campo in campos_independientes if variables_independientes[campo].get().strip()}

    # Crear una copia del DataFrame original
    datos_filtrados = data.copy()

    # Aplicar los filtros por selecciones generales
    for columna, valor in seleccion.items():
        if columna in datos_filtrados.columns:
            if columna == "Fecha":  # Filtro especial para el año
                try:
                    valor = int(valor)  # Convertir el año seleccionado a entero
                    fecha_inicio = f"{valor}-01-01"  # Inicio del año
                    fecha_fin = f"{valor}-12-31"  # Fin del año

                    # Filtrar datos dentro del rango del año
                    datos_filtrados = datos_filtrados[
                        (datos_filtrados["Fecha"] >= pd.to_datetime(fecha_inicio)) &
                        (datos_filtrados["Fecha"] <= pd.to_datetime(fecha_fin))
                    ]
                except ValueError:
                    messagebox.showerror("Error", "El valor del año seleccionado no es válido.")
                    return
            else:  # Filtro estándar para otras columnas
                datos_filtrados = datos_filtrados[
                    datos_filtrados[columna].astype(str).str.strip().str.lower() == valor.strip().lower()
                ]

    # Aplicar filtros basados en palabras clave de los campos de descripción
    for columna, palabra_clave in palabras_clave_descripcion.items():
        if columna in datos_filtrados.columns:
            datos_filtrados = datos_filtrados[
                datos_filtrados[columna].astype(str).str.lower().str.contains(palabra_clave, na=False)
            ]

    # Verificar si hay datos después de aplicar los filtros
    if datos_filtrados.empty:
        messagebox.showinfo("Sin resultados", "No hay datos que coincidan con los filtros seleccionados.")
        return

    # Cambiar el nombre de 'Unnamed: 0' a 'N°' y enumerar las filas
    if "Unnamed: 0" in datos_filtrados.columns:
        datos_filtrados = datos_filtrados.rename(columns={"Unnamed: 0": "N°"})
    datos_filtrados["N°"] = range(1, len(datos_filtrados) + 1)  # Enumerar las filas

    # Asegurarse de que la columna Fecha solo muestre la fecha (sin hora)
    if "Fecha" in datos_filtrados.columns:
        datos_filtrados["Fecha"] = datos_filtrados["Fecha"].dt.date  # Extraer solo la fecha

    # Guardar los datos filtrados en un archivo Excel
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        datos_filtrados.to_excel(save_path, index=False)  # Exportar usando pandas

        # Ajustar el ancho de columnas con openpyxl
        try:
            wb = load_workbook(save_path)  # Cargar el archivo guardado
            ws = wb.active  # Obtener la hoja activa

            # Ajustar el ancho de cada columna según su contenido
            for column_cells in ws.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter  # Letra de la columna (A, B, C, ...)
                for cell in column_cells:
                    if cell.value:  # Si la celda tiene un valor
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = max_length + 2  # Ajustar el ancho con un pequeño margen
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(save_path)  # Guardar los cambios
            wb.close()
            messagebox.showinfo("Éxito", f"Archivo exportado correctamente en:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo ajustar el tamaño de las columnas: {e}")



# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("SICE INFORMES 2025")
ventana.geometry("1000x600")
ventana.configure(bg="#2b2b2b")

# Frame del menú principal
frame_menu = tk.Frame(ventana, bg="#2b2b2b")
label_menu = tk.Label(frame_menu, text="SICE 2025\nMenú Principal", bg="#2b2b2b", fg="white", font=("Arial", 20, "bold"))
label_menu.pack(pady=20)

boton_anexo = tk.Button(
    frame_menu, 
    text="Ir a Informe anexos", 
    command=lambda: [frame_menu.pack_forget(), frame_anexo.pack(fill="both", expand=True)],
    font=("Arial", 10, "bold")
)
boton_anexo.pack(pady=10)

# Frame para la opción "Informe semanal"
frame_anexo = tk.Frame(ventana, bg="#2b2b2b")

# Configuración para centrar todo el contenido en el frame_anexo
frame_anexo.grid_rowconfigure(0, weight=1)  # Espacio arriba
frame_anexo.grid_rowconfigure(1, weight=1)  # Título
frame_anexo.grid_rowconfigure(2, weight=1)  # Botón insertar
frame_anexo.grid_rowconfigure(3, weight=2)  # Campos principales
frame_anexo.grid_rowconfigure(4, weight=1)  # Botón exportar
frame_anexo.grid_rowconfigure(5, weight=1)  # Botón volver
frame_anexo.grid_columnconfigure(0, weight=1)  # Centrar horizontalmente

# Título
label_anexo = tk.Label(
    frame_anexo,
    text="Gestión de Informes Semanales",
    bg="#2b2b2b",
    fg="white",
    font=("Arial", 20, "bold")
)
label_anexo.grid(row=1, column=0, pady=10, sticky="n")

# Etiqueta de información adicional
label_informacion = tk.Label(
    frame_anexo,
    text="Rellene los campos correspondientes.",
    bg="#1e1e1e",
    fg="#a9a9a9",
    font=("Arial", 10, "italic")
)
label_informacion.grid(row=2, column=0, pady=5, sticky="n")

# Botón para cargar archivo Excel
boton_insertar = tk.Button(
    frame_anexo,
    text="Cargar Archivo Excel",
    command=archivo_anexo,
    font=("Arial", 10, "bold"),
    activebackground="#4c70ba",
    activeforeground="white",
    bd=2,
    relief="raised"
)
boton_insertar.grid(row=3, column=0, pady=15, sticky="n")

# Frame para organizar las entradas en columnas
frame_campos = tk.Frame(frame_anexo, bg="#2b2b2b")
frame_campos.grid(row=4, column=0, pady=10, sticky="nsew")

# Configuración del frame_campos para centrar sus elementos
frame_campos.grid_rowconfigure(0, weight=1)
frame_campos.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)  # Espaciado uniforme

# Variables para los campos generales
campos = [
    "Fecha", "Semana", "TIPO SICE", "ESTADO SICE", "TIPO MMS", "Estado MMS",
    "SISTEMA", "TIPO", "NOMBRE DEL EQUIPO", "OTROS SISTEMAS", 
    "EMPLAZAMIENTO", "LINEA", "TRATAMIENTO"
]
variables = {campo: tk.StringVar(value="Seleccione") for campo in campos}
menus = {}

# Crear las columnas principales y centrarlas
num_columnas = 3
for i, campo in enumerate(campos):
    columna = i % num_columnas
    fila = i // num_columnas

    # Crear etiquetas y menús desplegables alineados al centro
    label = tk.Label(
        frame_campos,
        text=campo,
        bg="#2b2b2b",
        fg="white",
        font=("Arial", 10)
    )
    label.grid(row=fila, column=columna * 2, padx=10, pady=5, sticky="e")

    menu = tk.OptionMenu(frame_campos, variables[campo], "Seleccione", command=lambda v, c=campo: actualizar_seleccion(c, v))
    menu.config(width=15, bg="white", fg="#2b2b2b")
    menu.grid(row=fila, column=columna * 2 + 1, padx=5, pady=5, sticky="w")
    menus[campo] = menu

# Crear un marco adicional para los campos de descripción independientes
frame_independiente = tk.Frame(frame_campos, bg="#2b2b2b")
frame_independiente.grid(row=fila + 1, column=0, columnspan=num_columnas * 2, pady=20, sticky="nsew")

# Variables para los campos independientes
campos_independientes = ["DESCRIPCION", "DESCRIPCIÓN DE LA FALLA"]
variables_independientes = {campo: tk.StringVar(value="") for campo in campos_independientes}

# Centrando los campos de descripción
for i, campo in enumerate(campos_independientes):
    label = tk.Label(
        frame_independiente,
        text=campo,
        bg="#2b2b2b",
        fg="white",
        font=("Arial", 10)
    )
    label.grid(row=i, column=0, padx=10, pady=5, sticky="e")

    entry = tk.Entry(
        frame_independiente,
        textvariable=variables_independientes[campo],
        width=40,
        bg="white",
        fg="#2b2b2b",
        font=("Arial", 10)
    )
    entry.grid(row=i, column=1, padx=10, pady=5, sticky="w")

# Botón para exportar datos seleccionados
boton_exportar = tk.Button(
    frame_anexo,
    text="Exportar Datos Seleccionados",
    command=exportar_seleccion,
    font=("Arial", 10, "bold"),
    activebackground="#4c70ba",
    activeforeground="white",
    bd=2,
    relief="raised"
)
boton_exportar.grid(row=5, column=0, pady=10, sticky="n")

# Botón para volver al menú principal
boton_volver = tk.Button(
    frame_anexo,
    text="← Volver al menú",
    command=cambiar_a_menu,
    font=("Arial", 10, "bold")
)
boton_volver.grid(row=6, column=0, pady=20, sticky="s")

# Mostrar el frame principal
frame_menu.pack(fill="both", expand=True)

# Iniciar la aplicación
ventana.mainloop()
