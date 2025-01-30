import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import Workbook, load_workbook  # Para Excel
from docx import Document  # Para Word
import pandas as pd
from datetime import datetime
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.enum.table import WD_ALIGN_VERTICAL
from tkcalendar import Calendar
import matplotlib.pyplot as plt
from docx.oxml.ns import qn
import unicodedata

def normalizar_texto(texto):
    texto = str(texto).strip()
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('utf-8')
    texto = texto.lower()  # Convertir a minúsculas
    return texto

# Función para cambiar al menú principal
def cambiar_a_menu():
    for frame in [frame_anexo]:
        frame.pack_forget()
    frame_menu.pack(fill="both", expand=True)

def frame_semanal():
    frame_menu.pack_forget()
    frame_semanal.pack(fill="both", expand=True)

def cambiar_frame(actual, siguiente):
    actual.pack_forget()
    siguiente.pack(fill="both", expand=True)

# Funciones para cambiar el estilo del botón cuando el mouse pasa sobre él
def on_enter(event):
    event.widget.config(bg="#4c4c4c", fg="white")

def on_leave(event):
    event.widget.config(bg="#dbdcdc", fg="black")    

# Función para cargar el archivo Excel
def archivo_anexo():
    global data # Declarar la variable global
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
                label_fecha = tk.Label(frame_campos, text="Fecha (Año)", bg="#2b2b2b", fg="white", font=("arial", 10))
                label_fecha.grid(row=0, column=0, padx=10, pady=5, sticky="e")
                menus["Fecha"] = tk.OptionMenu(frame_campos, variables["Fecha"], *anios_disponibles, command=lambda v: actualizar_seleccion("Fecha", v))
                menus["Fecha"].grid(row=0, column=1, padx=10, pady=5, sticky="w")
            else:
                menus["Fecha"].children["menu"].delete(0, "end")
                for anio in anios_disponibles:
                    menus["Fecha"].children["menu"].add_command(label=anio, command=lambda v=anio: actualizar_seleccion("Fecha", v))

    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")


        # Rellenar valores nulos con un texto predeterminado
        for columna in columnas_requeridas:
            data[columna] = data[columna].fillna("Desconocido").astype(str).str.strip()

    except Exception as e:
        messagebox.showerror("Error", f"Error al cargar el archivo: {e}")

    # Actualizar semanas disponibles (si aplicable)
    if "Fecha" in data.columns:
        data["Semana"] = data["Fecha"].dt.isocalendar().week
        semanas_disponibles = data["Semana"].dropna().astype(int).unique().tolist()
        semanas_disponibles.sort()

        if "Semana" not in menus:
            label_semana = tk.Label(frame_campos, text="Semana", bg="#2b2b2b", fg="white", font=("arial", 10))
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
                label = tk.Label(frame_campos, text=columna, bg="#2b2b2b", fg="white", font=("arial", 10))
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

# Modificar la función exportar_seleccion
def exportar_seleccion():
    if data is None or data.empty:
        messagebox.showwarning("Advertencia", "No hay datos cargados para filtrar.")
        return

    # Recoger las selecciones de los campos generales
    seleccion = {campo: variables[campo].get() for campo in variables if variables[campo].get() != "Seleccione"}

    # Recoger palabras clave de los campos de descripción
    palabras_clave_descripcion = {campo: variables_independientes[campo].get().strip().lower()
                                   for campo in campos_independientes if variables_independientes[campo].get().strip()}

    # Crear una copia del DataFrame original
    datos_filtrados = data.copy()

    # Normalizar las columnas relevantes
    for columna in datos_filtrados.columns:
        if datos_filtrados[columna].dtype == object:  # Normalizar solo columnas de tipo texto
            datos_filtrados[columna] = datos_filtrados[columna].fillna("").apply(normalizar_texto)

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
                    datos_filtrados[columna] == normalizar_texto(valor)
                ]

    # Aplicar filtros basados en palabras clave de los campos de descripción
    for columna, palabra_clave in palabras_clave_descripcion.items():
        if columna in datos_filtrados.columns:  # Verificar que la columna exista
            if palabra_clave:
                # Dividir las palabras clave ingresadas por el usuario en una lista
                palabras = [normalizar_texto(p) for p in palabra_clave.strip().split()]

                # Lógica de filtro: Y u O
                if logica_filtro.get() == "Y":  # Lógica AND (todas las palabras deben coincidir)
                    for palabra in palabras:
                        datos_filtrados = datos_filtrados[
                            datos_filtrados[columna].str.contains(palabra, na=False)
                        ]
                elif logica_filtro.get() == "O":  # Lógica OR (al menos una palabra debe coincidir)
                    datos_filtrados = datos_filtrados[
                        datos_filtrados[columna].apply(lambda x: any(palabra in x for palabra in palabras))
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
label_menu = tk.Label(frame_menu, text="SICE 2025\nMenú Principal", bg="#2b2b2b", fg="white", font=("arial", 20, "bold"))
label_menu.pack(pady=20)

boton_anexo = tk.Button(
    frame_menu, 
    text="Ir a Informe anexos", 
    command=lambda: [frame_menu.pack_forget(), frame_anexo.pack(fill="both", expand=True)],
    font=("arial", 10, "bold")
)
boton_anexo.pack(pady=10)

boton_nuevo = tk.Button(
    frame_menu, 
    text="Ir a Nuevo Frame", 
    command=frame_semanal,
    font=("arial", 10, "bold")
)
boton_nuevo.pack(pady=10)

# Frame para la opción "Informe semanal"
frame_anexo = tk.Frame(ventana, bg="#2b2b2b")
label_anexo = tk.Label(
    frame_anexo,
    text="Gestión de Informes Semanales",
    bg="#2b2b2b",
    fg="white",
    font=("arial", 20, "bold")
)
label_anexo.pack(pady=20)


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
    font=("arial", 20, "bold")
)
label_anexo.grid(row=1, column=0, pady=10, sticky="n")

# Etiqueta de información adicional
label_informacion = tk.Label(
    frame_anexo,
    text="Rellene los campos correspondientes.",
    bg="#1e1e1e",
    fg="#a9a9a9",
    font=("arial", 10, "italic")
)
label_informacion.grid(row=2, column=0, pady=5, sticky="n")

# Botón para cargar archivo Excel
boton_insertar = tk.Button(
    frame_anexo,
    text="Cargar Archivo Excel",
    command=archivo_anexo,
    font=("arial", 10, "bold"),
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
frame_campos.grid_rowconfigure(0, weight=0)
frame_campos.grid_columnconfigure((0, 1, 2, 3, 4), weight=0)  # Espaciado uniforme

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
total_columnas = num_columnas * 2 + 2  # Añadir dos columnas vacías (izquierda y derecha)

# Configurar las columnas para centrar el contenido
for col in range(total_columnas):
    if col == 0 or col == total_columnas - 1:  # Columnas vacías (izquierda y derecha)
        frame_campos.grid_columnconfigure(col, weight=1)  # Espacio expansible
    else:  # Columnas con contenido
        frame_campos.grid_columnconfigure(col, weight=0)  # Fijas

# Crear las columnas principales centradas
for i, campo in enumerate(campos):
    columna = (i % num_columnas) * 2 + 1  # Ajustar posición dejando espacio a los lados
    fila = i // num_columnas

    # Crear etiquetas alineadas a la derecha
    label = tk.Label(
        frame_campos,
        text=campo,
        bg="#2b2b2b",
        fg="white",
        font=("arial", 10)
    )
    label.grid(row=fila, column=columna, padx=10, pady=5, sticky="e")  # Alinear a la derecha

    # Crear menús desplegables alineados a la izquierda
    menu = tk.OptionMenu(frame_campos, variables[campo], "Seleccione", command=lambda v, c=campo: actualizar_seleccion(c, v))
    menu.config(width=15, bg="white", fg="#2b2b2b")
    menu.grid(row=fila, column=columna + 1, padx=10, pady=5, sticky="w")  # Alinear a la izquierda
    menus[campo] = menu

# Crear un marco adicional para los campos de descripción independientes
frame_independiente = tk.Frame(frame_campos, bg="#2b2b2b")
frame_independiente.grid(row=fila + 1, column=0, columnspan=num_columnas * 2, pady=20, sticky="nsew")

# Variables para los campos independientes
campos_independientes = ["DESCRIPCION", "DESCRIPCIÓN DE LA FALLA"]
variables_independientes = {campo: tk.StringVar(value="") for campo in campos_independientes}

# Configuración del frame_independiente para centrar los campos de descripción
total_columnas_descripcion = 4  # Dos columnas vacías a los lados y dos para contenido

# Configurar las columnas para centrar el contenido
for col in range(total_columnas_descripcion):
    if col == 0 or col == total_columnas_descripcion - 1:  # Columnas vacías (izquierda y derecha)
        frame_independiente.grid_columnconfigure(col, weight=1)  # Espacio expansible
    else:  # Columnas con contenido
        frame_independiente.grid_columnconfigure(col, weight=0)  # Fijas

# Centrando los campos de descripción
for i, campo in enumerate(campos_independientes):
    # Etiquetas alineadas a la derecha
    label = tk.Label(
        frame_independiente,
        text=campo,
        bg="#2b2b2b",
        fg="white",
        font=("arial", 10)
    )
    label.grid(row=i, column=1, padx=10, pady=5, sticky="e")  # Columna 1: etiquetas

    # Entradas alineadas a la izquierda
    entry = tk.Entry(
        frame_independiente,
        textvariable=variables_independientes[campo],
        width=40,
        bg="white",
        fg="#2b2b2b",
        font=("arial", 10)
    )
    entry.grid(row=i, column=2, padx=10, pady=5, sticky="w")  # Columna 2: entradas

def actualizar_logica(valor):
    logica_filtro.set(valor)


def abrir_archivo_excel():
    try:
        # Abrir el archivo Excel
        ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta_archivo:
            # Leer la hoja L63 del archivo Excel
            df = pd.read_excel(ruta_archivo, sheet_name="L63")

            # Renombrar las columnas basándose en la primera fila del archivo
            df.columns = df.iloc[0]
            df = df[1:]

            # Verificar si la columna FE existe
            if "FE" not in df.columns:
                messagebox.showerror("Error", "El archivo no contiene una columna 'FE'.")
                return

            # Convertir la columna de FE a tipo datetime
            df["FE"] = pd.to_datetime(df["FE"], errors="coerce")

            # Verificar si hay fechas válidas
            if df["FE"].isnull().all():
                messagebox.showerror("Error", "La columna 'FE' no contiene fechas válidas.")
                return

            # Obtener la fecha mínima y máxima de la columna FE
            fecha_minima = df["FE"].min()
            fecha_maxima = df["FE"].max()

            # Validar rango de fechas
            if pd.isnull(fecha_minima) or pd.isnull(fecha_maxima):
                messagebox.showerror("Error", "No se encontraron fechas válidas para establecer el rango.")
                return

            # Crear una nueva ventana para seleccionar el rango de fechas
            ventana_fechas = tk.Toplevel()
            ventana_fechas.title("Filtrar por Fecha")
            ventana_fechas.geometry("800x700")

            # Etiquetas y calendarios para las fechas
            tk.Label(ventana_fechas, text="Fecha de inicio:").pack(pady=5)
            calendario_inicio = Calendar(
                ventana_fechas, selectmode="day", date_pattern="yyyy-mm-dd",
                mindate=min(fecha_minima, datetime.now() - pd.Timedelta(days=365)),
                maxdate=max(fecha_maxima, datetime.now() + pd.Timedelta(days=365))
            )
            calendario_inicio.pack(pady=10)

            tk.Label(ventana_fechas, text="Fecha de fin:").pack(pady=5)
            calendario_fin = Calendar(
                ventana_fechas, selectmode="day", date_pattern="yyyy-mm-dd",
                mindate=min(fecha_minima, datetime.now() - pd.Timedelta(days=365)),
                maxdate=max(fecha_maxima, datetime.now() + pd.Timedelta(days=365))
            )
            calendario_fin.pack(pady=10)

            # Función para aplicar el filtro y mostrar los resultados
            def aplicar_filtro():
                ventana_fechas.destroy()
                try:
                    fecha_inicio = datetime.strptime(calendario_inicio.get_date(), "%Y-%m-%d")
                    fecha_fin = datetime.strptime(calendario_fin.get_date(), "%Y-%m-%d")

                    # Validar que la fecha de inicio no sea mayor que la de fin
                    if fecha_inicio > fecha_fin:
                        messagebox.showerror("Error", "La fecha de inicio no puede ser mayor que la fecha de fin.")
                        return

                    # Filtrar los datos por el rango de fechas
                    df_filtrado = df[(df["FE"] >= fecha_inicio) & (df["FE"] <= fecha_fin)]

                    # Verificar si hay datos filtrados
                    if df_filtrado.empty:
                        messagebox.showinfo("Sin datos", "No se encontraron datos para el rango de fechas seleccionado.")
                        return

                    # Seleccionar columnas específicas para mostrar
                    df_filtrado = df_filtrado[["Descripcion OT", "Equipo", "DEP", "EST", "SIST", "F.LIBERACIÓN", "FP", "Número de OT", "FE", "CAT"]]
                    df_filtrado["FE"] = df_filtrado["FE"].dt.strftime('%Y-%m-%d')

                    crear_word(df_filtrado)

                except ValueError:
                    messagebox.showerror("Error", "Error al procesar las fechas seleccionadas.")

            # Botón para aplicar el filtro
            boton_aplicar = tk.Button(ventana_fechas, text="Aplicar Filtro", command=aplicar_filtro)
            boton_aplicar.pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo Excel:\n{e}")

import tkinter as tk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def crear_word(df_filtrado):
    # Crear un nuevo documento de Word
    doc = Document()

    doc.add_paragraph('\n'*2)

    def semanas_transcurridas():
        # Obtener el año actual
        año = datetime.now().year
        # Crear el objeto de fecha para el primer día del año
        fecha_inicial_obj = datetime.strptime(str(año)+("/01/01"), "%Y/%m/%d")

        # Obtener la fecha actual
        fecha_final_obj = datetime.now()  

        # Calcular la diferencia en días entre la fecha final y la inicial
        diferencia_dias = (fecha_final_obj - fecha_inicial_obj).days + 1  # Se suma 1 para contar el primer día como parte de la semana

        # Calcular las semanas transcurridas (redondeando hacia arriba si es necesario)
        semanas_transcurridas = diferencia_dias // 7
        if diferencia_dias % 7 != 0:
            semanas_transcurridas += 1  # Si no es exacto, contar la semana adicional
        return str(semanas_transcurridas)

    # Para centrar la imagen, puedes agregar un parágrafo vacío antes y después de la imagen
    # y luego utilizar el método `alignment` para centrar el parágrafo
    imagen_portada = doc.add_paragraph()
    imagen_portada.alignment = 1  # 1 es el valor para centrar
    imagen_portada.add_run().add_picture("imagenes/logo_metro.png", width=Inches(3))
    imagen_portada = doc.add_paragraph()

    portada = doc.add_paragraph('\nMANTENIMIENTO DEL SISTEMA DE COMUNICACIONES \nLINEAS 6 Y 3\nMETRO DE SANTIAGO \n\nREPORTE SEMANAL ' + str(datetime.now().year) + '\nSEMANA ' + semanas_transcurridas() + '\nCONTRATO N° MN-236-2014-G')
    portada.alignment = 1
    portada.runs[0].font.name = 'Calibri'
    portada.runs[0].font.size = Pt(15.5)
    portada.runs[0].font.bold = True

    # Acceder al pie de página de la primera sección
    footer = doc.sections[0].footer

    # Añadir un párrafo al pie de página
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT  

    # Crear un campo para el número de página
    page_field = OxmlElement('w:fldSimple')
    page_field.set(qn('w:instr'), 'PAGE')

    # Añadir el campo al párrafo
    run = p.add_run()
    run._r.append(page_field)

    # Opcional: Personalizar el estilo del texto
    run.font.name = 'Calibri'
    run.font.size = Pt(9.5)

    # Acceder a la sección del encabezado
    section = doc.sections[0]
    header = section.header

    # Crear una tabla en el encabezado con 3 celdas (izquierda, centro, derecha)
    table = header.add_table(rows=1, cols=3, width=Inches(6))
    table.autofit = False

    # Configurar el ancho de las columnas
    for idx, width in enumerate([Inches(2), Inches(2), Inches(2)]):
        table.columns[idx].width = width

    # Obtener la primera fila de la tabla
    row = table.rows[0].cells

    # Insertar la imagen en la celda izquierda
    celda_izquierda = row[0].paragraphs[0]
    try:
        celda_izquierda.add_run().add_picture("imagenes/logo_metro.png", width=Inches(1.2))
        celda_izquierda.paragraph_format.alignment = 0
    except Exception as e:
        print(f"Error al cargar la imagen izquierda: {e}")

    # Insertar el texto en la celda central
    celda_centro = row[1].paragraphs[0]
    celda_centro.text = "METRO DE SANTIAGO LÍNEAS 6 Y 3 \nINFORME SEMANA " + semanas_transcurridas() + "\n Contrato N° MN-236-2014-G" 
    celda_centro.runs[0].font.name = 'Calibri'
    celda_centro.runs[0].font.size = Pt(8.5)
    celda_centro.paragraph_format.alignment = 1  # Centrar el texto (1 = CENTER)

    # Insertar la imagen en la celda derecha
    celda_derecha = row[2].paragraphs[0]
    try:
        celda_derecha.add_run().add_picture("imagenes/logo_sice.png", width=Inches(1))
        celda_derecha.paragraph_format.alignment = 2
    except Exception as e:
        print(f"Error al cargar la imagen derecha: {e}")
    
    doc.add_paragraph("\n" * 9)
    # Añadir una tabla con el número de filas y columnas necesarias
    tabla_portada = doc.add_table(rows=4, cols=7)
    tabla_portada.style = 'Table Grid'

    # Modificar el ancho de las celdas (opcional)
    for row in tabla_portada.rows:
        for cell in row.cells:
            cell.width = Pt(70)

    # Contenido de las celdas
    header_row = ['\n'+str(Rev_documento.get())+'\n', '09-12-2024', 'REVISIÓN', 'HSANMARTIN', 'RZAMBRANO', 'MCOLLAO', 'LLÓPEZ']
    sub_header_row = ['REV.', 'FECHA', 'EMITIDO PARA', 'PREPARÓ', 'REVISÓ', 'APROBÓ', 'APROBÓ METRO']
    footer_row = ['Nº DOCUMENTO', '', 'P63-MA-0632-INT-000-CO-00347', '', '\nPágina 1 de 20', '', 'REV. ' + str(Rev_documento.get())]

    # Función para dar formato a las celdas con estilo de letra
    def format_cell(cell, text, font_name="Calibri", font_size=9.5, font_bold= False):
        cell.text = text
        paragraph = cell.paragraphs[0]
        run = paragraph.runs[0]
        run.font.name = font_name  # Cambiar la fuente
        run.font.size = Pt(font_size)  # Cambiar el tamaño de letra
        run.font.bold = font_bold  # Negrita
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Centrar horizontalmente
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    # Añadir datos y aplicar estilo a cada celda
    for i, text in enumerate(header_row):
        format_cell(tabla_portada.rows[0].cells[i], text, font_name="Calibri", font_size=9.5, font_bold=False)
    for i, text in enumerate(sub_header_row):
        format_cell(tabla_portada.rows[1].cells[i], text, font_name="Calibri", font_size=9.5, font_bold =True)
    for i, text in enumerate(footer_row):
        if i == 0 or i == 6:
            format_cell(tabla_portada.rows[2].cells[i], text, font_name="Calibri", font_size=9.5, font_bold=True)
        else:
            format_cell(tabla_portada.rows[2].cells[i], text, font_name="Calibri", font_size=9.5, font_bold=False)

    logo_tabla = tabla_portada.rows[3].cells[6].paragraphs[0]
    logo_tabla.add_run().add_picture("imagenes/logo_sice.png", width=Inches(1))
    logo_tabla.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Unir celdas 
    tabla_portada.rows[2].cells[4].merge(tabla_portada.rows[2].cells[5])
    tabla_portada.rows[2].cells[1].merge(tabla_portada.rows[2].cells[3])
    tabla_portada.rows[3].cells[0].merge(tabla_portada.rows[3].cells[6])
    
    doc.add_page_break()   

    # PAGINA 2

    Control_cambios = doc.add_paragraph('\nCONTROL DE CAMBIOS')
    Control_cambios.alignment = 1
    Control_cambios.runs[0].font.name = 'Calibri'
    Control_cambios.runs[0].font.size = Pt(10)
    Control_cambios.runs[0].font.bold = True

    tabla_2 = doc.add_table(rows=2, cols=7)
    tabla_2.style = 'Table Grid'

    Rev_titulo = ['REVISIÓN', 'Fecha', 'PÁGINA', 'PÁRRAFO', '', 'MODIFICACIÓN REALIZADA', '']
    Rev_contenido= [str(Rev_documento.get()), str(datetime.today().date()), 'N/A', 'N/A', '', 'PRIMERA VERSIÓN DEL DOCUMENTO', '']

    for i, text in enumerate(Rev_titulo):
        format_cell(tabla_2.rows[0].cells[i], text, font_name="Calibri", font_size=9.5, font_bold=False)
    for i, text in enumerate(Rev_contenido):
        format_cell(tabla_2.rows[1].cells[i], text, font_name="Calibri", font_size=9.5, font_bold =False)
    tabla_2.rows[0].cells[4].merge(tabla_portada.rows[0].cells[6])
    tabla_2.rows[1].cells[4].merge(tabla_portada.rows[1].cells[6])
    doc.add_page_break()
    # indice

    titulo_indice = doc.add_heading('Indice')
    titulo_indice.runs[0].font.name = 'Calibri'
    titulo_indice.runs[0].font.size = Pt(10)
    titulo_indice.runs[0].font.bold = True
    titulo_indice.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro
    titulo_indice.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    doc.add_page_break()    

    # Objetivo
    titulo_objetivo = doc.add_heading('1. Objetivo', 1)
    titulo_objetivo.runs[0].font.name = 'Calibri'
    titulo_objetivo.runs[0].font.size = Pt(10)
    titulo_objetivo.runs[0].font.bold = True
    titulo_objetivo.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro
    objeto = doc.add_paragraph('\n        Reportar el desarrollo de las actividades de Mantenimiento Preventivo, Correctivo, eventos y Hallazgos asociados a la gestión del contrato de Mantenimiento MN-236-2014-G correspondiente a la Semana '+ semanas_transcurridas()+ ' del año ' +str(datetime.now().year)+ ' del sistema de Comunicaciones de las líneas 6 y 3 del Metro de Santiago.')
    objeto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    objeto.runs[0].font.name = 'Calibri'
    objeto.runs[0].font.size = Pt(9.5)
    objeto.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro

    doc.add_page_break()   

    # Mantenimiento
    titulo_mantenimiento = doc.add_heading('2. Mantenimiento Preventivo', 1)
    titulo_mantenimiento.runs[0].font.name = 'Calibri'
    titulo_mantenimiento.runs[0].font.size = Pt(10)
    titulo_mantenimiento.runs[0].font.bold = True
    titulo_mantenimiento.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro

    titulo_mantenimiento2_1 = doc.add_heading('2.1.   Programación Semanal del Mantenimiento Preventivo ', 2)
    titulo_mantenimiento2_1.runs[0].font.name = 'Calibri'
    titulo_mantenimiento2_1.runs[0].font.size = Pt(10)
    titulo_mantenimiento2_1.runs[0].font.bold = True
    titulo_mantenimiento2_1.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro

    mantenimiento2_1 = doc.add_paragraph('A continuación, se presentan los mantenimientos preventivos según su categoría que fueron planificados para la Semana ' + semanas_transcurridas()+ ' del año ' +str(datetime.now().year))
    mantenimiento2_1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    mantenimiento2_1.runs[0].font.name = 'Calibri'
    mantenimiento2_1.runs[0].font.size = Pt(9.5)

    titulo_mantenimiento_CAT1 = doc.add_paragraph('\nMantenimientos Preventivos CAT 1')
    titulo_mantenimiento_CAT1.runs[0].font.name = 'Calibri'
    titulo_mantenimiento_CAT1.runs[0].font.size = Pt(10)
    titulo_mantenimiento_CAT1.runs[0].font.bold = True
    titulo_mantenimiento_CAT1.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro
    titulo_mantenimiento_CAT1.alignment = 1

    mantenimiento_CAT1 = doc.add_paragraph('\n      Para la semana en estudio se registran los siguientes mantenimientos preventivos de categoría 1.')
    mantenimiento_CAT1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    mantenimiento_CAT1.runs[0].font.name = 'Calibri'
    mantenimiento_CAT1.runs[0].font.size = Pt(9.5)

    try:
        if not df_filtrado.empty:
            # Filtrar los valores de CAT1
            if 'CAT' in df_filtrado.columns:
                df_filtro1 = df_filtrado[df_filtrado['CAT'] == 'CAT 1']

            # Eliminar las columnas no requeridas del DataFrame si existen
            columnas_requeridas = ['Descripcion OT', 'Equipo', 'DEP', 'EST', 'SIST', 'FP', 'F.LIBERACIÓN', 'Número de OT']
            columnas_a_eliminar = [col for col in df_filtro1.columns if col not in columnas_requeridas]
            df_filtro1 = df_filtro1.drop(columns=columnas_a_eliminar, errors='ignore')

            # Formatear las columnas 'F.LIBERACIÓN' y 'FP' para que solo contengan la fecha
            for col in ['F.LIBERACIÓN', 'FP']:
                if col in df_filtro1.columns:
                    df_filtro1[col] = pd.to_datetime(df_filtro1[col], errors='coerce').dt.date

            # Ordenar las columnas según el formato deseado
            df_filtro1 = df_filtro1[['Descripcion OT', 'Equipo', 'DEP', 'EST', 'SIST', 'FP', 'F.LIBERACIÓN', 'Número de OT']]

            # Crear la tabla con las columnas requeridas
            columnas = ["N°", "DESCRIPCIÓN", "Equipo", "Línea", "Estación", "Sistema", "Plan Matriz", "Ejecutado", "OT"]
            num_filas = df_filtro1.shape[0]  # Número de filas de datos
            num_columnas = len(columnas)  # Número de columnas predefinidas
            tablaCAT1 = doc.add_table(rows=num_filas + 1, cols=num_columnas)  # Crear tabla con encabezados y filas de datos
            tablaCAT1.style = 'Table Grid'

            # Agregar encabezados
            for j, column_name in enumerate(columnas):
                cell = tablaCAT1.cell(0, j)
                cell.text = column_name
                paragraph = cell.paragraphs[0]
                run = paragraph.runs[0]
                run.font.name = 'Calibri'
                run.font.size = Pt(6)
                run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # Agregar los datos a la tabla
            for i, row in enumerate(df_filtro1.itertuples(index=False), start=1):
                # Agregar numeración en la primera columna
                cell = tablaCAT1.cell(i, 0)
                cell.text = str(i)
                paragraph = cell.paragraphs[0]
                run = paragraph.runs[0]
                run.font.name = 'Calibri'
                run.font.size = Pt(6)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # Agregar el resto de los datos
                for j, value in enumerate(row, start=1):
                    cell = tablaCAT1.cell(i, j)
                    cell.text = str(value)
                    paragraph = cell.paragraphs[0]
                    run = paragraph.runs[0]
                    run.font.name = 'Calibri'
                    run.font.size = Pt(6)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    # Ajustar altura de la fila para la columna de "DESCRIPCIÓN"
                    if columnas[j] == "DESCRIPCIÓN":
                        cell.width = Inches(3)  # Duplicar el espacio en la celda de descripción

                    # Ajustar espacio para la columna "Plan Matriz"
                    if columnas[j] == "Plan Matriz":
                        cell.width = Inches(1.1)  

                # Centrar texto en todas las celdas
                for row in tablaCAT1.rows:
                    for cell in row.cells:
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        else:
            doc.add_paragraph('\nNo se encontraron datos para mostrar.')
    except Exception as e:
        print(f"Error en la creación del Word: {e}")
    
    mantenimiento2_1 = doc.add_paragraph('Tabla N°3 - Mantenimientos Categoría 3. \n')
    mantenimiento2_1.alignment = 1
    mantenimiento2_1.runs[0].font.name = 'Calibri'
    mantenimiento2_1.runs[0].font.size = Pt(9.5)
    
    doc.add_page_break()  
    
    # 2.2.2 Adherencia al plan de mantenimiento
    
    titulo_adherencia = doc.add_heading('2.2.2	 Adherencia Plan de Mantenimiento', 1)
    titulo_adherencia.runs[0].font.name = 'Calibri'
    titulo_adherencia.runs[0].font.size = Pt(10)
    titulo_adherencia.runs[0].font.bold = True
    titulo_adherencia.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro
    
    # Tabla para Desviación al Plan de Mantenimiento
    try:
        if not df_filtrado.empty:
            # Eliminar las columnas no requeridas del DataFrame si existen
            columnas_requeridas = ['Descripcion OT', 'DEP', 'EST', 'CAT', 'Número de OT', 'F.LIBERACIÓN']
            columnas_a_eliminar = [col for col in df_filtrado.columns if col not in columnas_requeridas]
            df_filtro2 = df_filtrado.drop(columns=columnas_a_eliminar, errors='ignore')
            
            # Formatear las columnas 'F.LIBERACIÓN' y 'FP' para que solo contengan la fecha
            for col in ['F.LIBERACIÓN', 'FP']:
                if col in df_filtro2.columns:
                    df_filtro2[col] = pd.to_datetime(df_filtro2[col], errors='coerce').dt.date
            
            # Ordenar las columnas según el formato deseado
            df_filtro2 = df_filtro2[['Descripcion OT', 'DEP', 'EST', 'CAT', 'Número de OT', 'F.LIBERACIÓN']]
            
            # Crear la tabla con las columnas requeridas
            columnas = ["N°", "DESCRIPCIÓN", "Línea", "Estación", "CAT", "OT", "Plan Matriz", "Motivo Desviación"]
            num_filas = df_filtro2.shape[0]  # Número de filas de datos
            num_columnas = len(columnas)  # Número de columnas predefinidas
            tablaCAT3 = doc.add_table(rows=num_filas + 1, cols=num_columnas)  # Crear tabla con encabezados y filas de datos
            tablaCAT3.style = 'Table Grid'
            
            # Agregar encabezados
            for j, column_name in enumerate(columnas):
                cell = tablaCAT3.cell(0, j)
                cell.text = column_name
                paragraph = cell.paragraphs[0]
                run = paragraph.runs[0]
                run.font.name = 'Calibri'
                run.font.size = Pt(6)
                run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Agregar los datos a la tabla
            for i, row in enumerate(df_filtro2.itertuples(index=False), start=1):
                # Agregar numeración en la primera columna
                cell = tablaCAT3.cell(i, 0)
                cell.text = str(i)
                paragraph = cell.paragraphs[0]
                run = paragraph.runs[0]
                run.font.name = 'Calibri'
                run.font.size = Pt(6)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Agregar el resto de los datos
                for j, value in enumerate(row, start=1):
                    cell = tablaCAT3.cell(i, j)
                    cell.text = str(value)
                    paragraph = cell.paragraphs[0]
                    run = paragraph.runs[0]
                    run.font.name = 'Calibri'
                    run.font.size = Pt(6)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Ajustar altura de la fila para la columna de "DESCRIPCIÓN"
                    if columnas[j] == "DESCRIPCIÓN":
                        cell.width = Inches(3)  # Duplicar el espacio en la celda de descripción
                    
                    # Ajustar espacio para la columna "Plan Matriz"
                    if columnas[j] == "Plan Matriz":
                        cell.width = Inches(1.1)
            
            # Centrar texto en todas las celdas
            for row in tablaCAT3.rows:
                for cell in row.cells:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        else:
            doc.add_paragraph('\nNo se encontraron datos para mostrar.')
    except Exception as e:
        print(f"Error en la creación del Word: {e}")

#
    titulo_adherencia2_3 = doc.add_heading('2.2.3 	Adherencia al Plan de Mantenimiento últimas 12 Semanas \n', 3)
    titulo_adherencia2_3.runs[0].font.name = 'Calibri'
    titulo_adherencia2_3.runs[0].font.size = Pt(10)
    titulo_adherencia2_3.runs[0].font.bold = True
    titulo_adherencia2_3.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro

    tabla_4 = doc.add_table(rows=6, cols=7)
    tabla_4.style = 'Table Grid'    

    # Fusión inicial de celdas
    tabla_4.rows[0].cells[0].merge(tabla_4.rows[0].cells[6])
    tabla_4.rows[1].cells[5].merge(tabla_4.rows[1].cells[6])
    tabla_4.rows[1].cells[3].merge(tabla_4.rows[1].cells[4])
    tabla_4.rows[2].cells[3].merge(tabla_4.rows[2].cells[4])
    tabla_4.rows[2].cells[5].merge(tabla_4.rows[2].cells[6])

    tabla_4.rows[3].cells[3].merge(tabla_4.rows[3].cells[4])
    tabla_4.rows[3].cells[5].merge(tabla_4.rows[3].cells[6])
    tabla_4.rows[4].cells[3].merge(tabla_4.rows[4].cells[4])
    tabla_4.rows[4].cells[5].merge(tabla_4.rows[4].cells[6])
    # Fusión de columnas en la misma sección
    tabla_4.rows[2].cells[6].merge(tabla_4.rows[4].cells[6])
    tabla_4.rows[2].cells[3].merge(tabla_4.rows[4].cells[3])
    
    adherencia_tabla = ['\nMANTENIMIENTO PREVENTIVO\n']
    adherencia_1 = ['', '\nProgramado\n', 'Ejecutado', 'Programado Total', '','Ejecutado Total']
    adherencia_2 = ['Categoría1 (C1)', 'numero', 'numero', 'numero', '','numero']
    adherencia_3 = ['Categoría1 (C2)', 'numero', 'numero', 'numero', '','numero']
    adherencia_4 = ['Categoría1 (C3)', 'numero', 'numero', 'numero', '','numero']
    adherencia_5 = ['Cumplimiento Semana '+ semanas_transcurridas(), '', '', '', '\nporcentaje\n', '','']

    # Aplicar formato y contenido
    for i, text in enumerate(adherencia_tabla):
        format_cell(tabla_4.rows[0].cells[i], text, font_name="Calibri", font_size=9.5, font_bold=False)

    # Agregar datos dinámicamente con formato
    adherencias = [adherencia_1, adherencia_2, adherencia_3, adherencia_4, adherencia_5]
    for row_idx, adherencia in enumerate(adherencias, start=1):
        for col_idx, text in enumerate(adherencia):
            if text.strip():  # Agregar solo texto no vacío
                format_cell(tabla_4.rows[row_idx].cells[col_idx], text, font_name="Calibri", font_size=9.5, font_bold=False)

    tabla_4.rows[5].cells[0].merge(tabla_4.rows[5].cells[3])
    tabla_4.rows[5].cells[4].merge(tabla_4.rows[5].cells[6])

    doc.add_page_break()

    tabla_adherencia = doc.add_paragraph('Tabla N°6 – Resumen Mantenimientos Preventivos Semana ' + semanas_transcurridas())
    tabla_adherencia.runs[0].font.name = 'Calibri'
    tabla_adherencia.runs[0].font.size = Pt(9.5)
    tabla_adherencia.runs[0].font.bold = False
    tabla_adherencia.alignment = 1
    tabla_adherencia.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro

    # Datos
    categorias = ["RMS", "TETRA", "RBA", "CCTV", "SIP", "SAP", "TEL", "CRO", "SCA", "GRA", "MR SIP/SAP"]
    programada = [12, 12, 12, 0, 0, 0, 0, 0, 0, 0, 12]
    ejecutada = [12, 12, 12, 0, 0, 0, 0, 0, 0, 0, 12]

    # Crear el gráfico
    x = range(len(categorias))  # Índices para las barras
    plt.bar(x, programada, width=0.4, label='PROGRAMADA', color='blue', align='center')
    plt.bar(x, ejecutada, width=0.4, label='EJECUTADA', color='orange', align='edge')

    # Personalización del gráfico
    plt.title("ADHERENCIA AL PLAN DE MANTENIMIENTO")
    plt.ylabel("ORDENES DE TRABAJO\nMANTENIMIENTO PREVENTIVO")
    plt.xticks(x, categorias, rotation=45)
    plt.legend()

    # Guardar el gráfico como una imagen temporal
    grafico2 = "grafico_temporal.png"
    plt.tight_layout()
    plt.savefig(grafico2)
    plt.close()

    # Cambiar el estilo de fuente predeterminado del documento
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(9.5)

    # Agregar el gráfico al documento
    imagen_parrafo = doc.add_paragraph()
    imagen_parrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    imagen_parrafo.add_run().add_picture(grafico2, width=Inches(4))
    graf = doc.add_paragraph("Gráfico N°2 – Adherencia Plan de Mantenimiento Preventivo Semana " + semanas_transcurridas() + ".")
    graf.name = 'Calibri'
    graf.size = Pt(9.5)
    graf.alignment = 1

    # Crear la tabla con los datos
    tabla_grafico1 = doc.add_table(rows=3, cols=len(categorias) + 1)
    tabla_grafico1.style = 'Table Grid'

    # Encabezado
    format_cell(tabla_grafico1.cell(0, 0), "Categoría", font_name="Calibri", font_size=9.5, font_bold =False)
    for i, categoria in enumerate(categorias):
        format_cell(tabla_grafico1.cell(0, i + 1), categoria)

    # Fila PROGRAMADA
    format_cell(tabla_grafico1.cell(1, 0), "PROGRAMADA", font_name="Calibri", font_size=9.5, font_bold =False)
    for i, valor in enumerate(programada):
        format_cell(tabla_grafico1.cell(1, i + 1), str(valor))

    # Fila EJECUTADA
    format_cell(tabla_grafico1.cell(2, 0), "EJECUTADA", font_name="Calibri", font_size=9.5, font_bold =False)
    for i, valor in enumerate(ejecutada):
        format_cell(tabla_grafico1.cell(2, i + 1), str(valor))

    graf = doc.add_paragraph("Tabla N°6 – Adherencia Plan de Mantenimiento Preventivo Semana " + semanas_transcurridas() + ".")
    graf.name = 'Calibri'
    graf.size = Pt(9.5)
    graf.alignment = 1

    doc.add_page_break()

    plan_mantenimiento = doc.add_heading("2.2.4	Proyección Mantenimiento Próximas 12 Semanas", 3)
    plan_mantenimiento.name = 'Calibri'
    plan_mantenimiento.size = Pt(10)
    plan_mantenimiento.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Negro

    mantenimiento3 = doc.add_paragraph("       \nA continuación, se declara proyección del Mantenimiento Preventivo para las próximas 12 semanas.")
    mantenimiento3.name = 'Calibri'
    mantenimiento3.size = Pt(9.5)



    doc.add_page_break()


################################# 
        # Guardar el documento con el nombre especificado
    doc.save("prueba.docx")
    print(f"Portada agregada correctamente. Archivo guardado")

# Variable global para la lógica (Y u O)
logica_filtro = tk.StringVar(value="Y")  # Valor inicial: "Y"

# Configurar el frame_independiente para tener espacio adicional para el botón "Y/O"
total_columnas_descripcion = 5  # Aumentar a 5 para incluir espacio adicional para el botón "Y/O"

# Configurar las columnas para centrar el contenido
for col in range(total_columnas_descripcion):
    if col == 0 or col == total_columnas_descripcion - 1:  # Columnas vacías (izquierda y derecha)
        frame_independiente.grid_columnconfigure(col, weight=1)  # Espacio expansible
    else:  # Columnas con contenido
        frame_independiente.grid_columnconfigure(col, weight=0)  # Fijas

# Mover el botón "Y/O" a la derecha del frame_independiente
label_logica = tk.Label(
    frame_independiente,
    text="Lógica (Y/O)",
    bg="#2b2b2b",
    fg="white",
    font=("arial", 10)
)
label_logica.grid(row=0, column=3, padx=10, pady=5, sticky="e")  # Columna 3 para el texto "Lógica (Y/O)"

menu_logica = tk.OptionMenu(
    frame_independiente,
    logica_filtro,
    "Y", "O",
    command=actualizar_logica
)
menu_logica.config(width=10, bg="white", fg="#2b2b2b")
menu_logica.grid(row=0, column=4, padx=10, pady=5, sticky="w")  # Columna 4 para el menú desplegable

# Botón para exportar datos seleccionados
boton_exportar = tk.Button(
    frame_anexo,
    text="Exportar Datos Seleccionados",
    command=exportar_seleccion,
    font=("arial", 10, "bold"),
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
    font=("arial", 10, "bold")
)
boton_volver.grid(row=6, column=0, pady=20, sticky="s")

# Crear el frame_semanal (nuevo frame)
frame_semanal = tk.Frame(ventana, bg="#2b2b2b")
label_semanal = tk.Label(frame_semanal, text="Informe Semanal", bg="#2b2b2b", fg="white", font=("arial", 20))
label_semanal.pack(pady=20)

# Botón para abrir archivo Excel
boton_abrir_excel = tk.Button(
    frame_semanal, text="Abrir Archivo Excel",
    command=abrir_archivo_excel,
    font=("arial", 10)
)
boton_abrir_excel.pack(pady=10)

# Aquí se agrega el menú desplegable al frame_semanal
opciones = ["A", "B", "0"]
Rev_documento = tk.StringVar(value="REV. documento")  # Variable para almacenar la selección

menu_opciones = tk.OptionMenu(
    frame_semanal,
    Rev_documento,
    *opciones
)
menu_opciones.config(
    font=("arial", 8),
    activebackground="#4b4b4b",
    activeforeground="white"
)
menu_opciones["menu"].config(bg="#1e1e1e", fg="white")  # Estilo del menú desplegable
menu_opciones.pack(pady=10)

# Botón para volver al menú principal
boton_volver_nuevo = tk.Button(
    frame_semanal, text="← Volver al Menú",
    command=lambda: cambiar_frame(frame_semanal, frame_menu),
    font=("arial", 10)
)
boton_volver_nuevo.pack(pady=10)
# Mostrar el frame principal
frame_menu.pack(fill="both", expand=True)

# Botones del frame_menu
boton_anexo.bind("<Enter>", on_enter)
boton_anexo.bind("<Leave>", on_leave)

boton_nuevo.bind("<Enter>", on_enter)
boton_nuevo.bind("<Leave>", on_leave)

# Botones del frame_anexo
boton_insertar.bind("<Enter>", on_enter)
boton_insertar.bind("<Leave>", on_leave)

boton_exportar.bind("<Enter>", on_enter)
boton_exportar.bind("<Leave>", on_leave)

boton_volver.bind("<Enter>", on_enter)
boton_volver.bind("<Leave>", on_leave)

# Botones del frame_semanal
boton_abrir_excel.bind("<Enter>", on_enter)
boton_abrir_excel.bind("<Leave>", on_leave)

boton_volver_nuevo.bind("<Enter>", on_enter)
boton_volver_nuevo.bind("<Leave>", on_leave)


# Iniciar la aplicación
ventana.mainloop()
