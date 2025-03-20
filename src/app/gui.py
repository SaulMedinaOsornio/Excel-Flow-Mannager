"""
Importaciones y configuración inicial de la aplicación.

Este bloque de código importa las bibliotecas necesarias y define las configuraciones iniciales
para la aplicación, incluyendo las rutas de archivos, parámetros predeterminados y dimensiones
de la ventana.

Importaciones:
- `os`: Módulo estándar para interactuar con el sistema operativo, en particular para gestionar
  rutas de archivos.
- `Path` de `pathlib`: Para trabajar con rutas de archivos y directorios de forma más conveniente
  y moderna.
- `filedialog`, `Tk`, `Canvas`, `Entry`, `Button`, `PhotoImage` de `tkinter`: Componentes y
  funcionalidades de Tkinter utilizados para la interfaz gráfica de usuario.
- `Image` y `ImageTk` de `PIL`: Para manejar imágenes dentro de la aplicación.
- `messagebox` de `tkinter`: Para mostrar mensajes emergentes al usuario.
- `json`: Para manejar la lectura y escritura de archivos JSON.
- `generator`, `excelFormat` desde `src.tools`: Herramientas personalizadas para procesar archivos
  y generar contenido específico de la aplicación.

Definición de rutas y parámetros iniciales:
- `OUTPUT_PATH`: Ruta del directorio donde se encuentra el archivo actual.
- `json_filename`: Ruta absoluta del archivo JSON que contiene los datos de salida.
- `ASSETS_PATH`: Ruta del directorio de recursos (imágenes) relacionados con la interfaz.
- `WINDOW_LENGTH` y `WINDOW_HIGH`: Dimensiones iniciales de la ventana de la aplicación (1100x700 píxeles).
- `file_path` y `output_file`: Variables para almacenar las rutas de archivo y salida seleccionadas por el usuario.
- `default_params`: Lista de parámetros predeterminados que se utilizarán en algún proceso dentro
  de la aplicación. Cada parámetro es una lista con dos valores (ID, fila, columna).
"""

# Importaciones
import os
from pathlib import Path
from tkinter import filedialog
from tkinter import Tk, Canvas, Entry, Button, PhotoImage
from PIL import Image, ImageTk
from src.tools import generator
from src.tools import excelFormat
from tkinter import messagebox
import tkinter as tk
import json

# Definición de rutas y parámetros
OUTPUT_PATH = Path(__file__).parent  # Ruta del directorio donde está el script actual
json_filename = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output", "data.json"))  # Ruta del archivo JSON de salida
ASSETS_PATH = OUTPUT_PATH / Path(r"../images")  # Ruta de los recursos (imágenes)
WINDOW_LENGTH = 1100  # Ancho de la ventana
WINDOW_HIGH = 700  # Alto de la ventana
file_path = ""  # Variable para almacenar la ruta del archivo seleccionado por el usuario
output_file = ""  # Variable para almacenar la ruta del directorio de salida

# Parámetros predeterminados (ID, fila, columna)
default_params = [
    [10, 2],  # Primer parámetro: ID = 10, fila = 2
    [12, 2],  # Segundo parámetro: ID = 12, fila = 2
    [14, 2],  # Tercer parámetro: ID = 14, fila = 2
    [17, 23],  # Cuarto parámetro: ID = 17, fila = 23
    [20, 2],  # Quinto parámetro: ID = 20, fila = 2
    [0, 0]  # Sexto parámetro: ID = 0, fila = 0
]



def ex():
    """
    Configura e inicializa la ventana principal de la aplicación "Excel Flow Manager".

    Esta función es responsable de crear la ventana principal de la interfaz gráfica de usuario
    (GUI) utilizando Tkinter. Realiza la configuración de la ventana, establece su título, icono,
    tamaño y color de fondo, y crea un lienzo (canvas) dentro de la ventana para agregar elementos
    gráficos.

    Pasos:
    1. Crea la ventana principal de la aplicación.
    2. Establece el título de la ventana como "Excel Flow Manager".
    3. Establece el ícono de la ventana desde un archivo de imagen.
    4. Define el tamaño de la ventana como 1100x700 píxeles.
    5. Establece el color de fondo de la ventana como blanco.
    6. Crea un lienzo (canvas) donde se podrán dibujar o colocar otros elementos gráficos.
    """

    # Crear la ventana principal de la aplicación
    window = Tk()

    # Establecer el título de la ventana
    window.title("Excel Flow Manager")

    # Establecer el ícono de la ventana desde un archivo de imagen
    icono = PhotoImage(file="../images/icono.png")
    window.iconphoto(True, icono)

    # Definir el tamaño de la ventana
    window.geometry("1100x700")

    # Establecer el color de fondo de la ventana
    window.configure(bg="#FFFFFF")

    # Crear un lienzo dentro de la ventana para agregar elementos gráficos
    canvas = Canvas(
        window,
        bg="#FFFFFF",  # Fondo blanco para el lienzo
        height=700,  # Altura del lienzo
        width=1100,  # Ancho del lienzo
        bd=0,  # Sin borde alrededor del lienzo
        highlightthickness=0,  # Sin grosor de resaltado
        relief="ridge"  # Estilo de borde (sin relieve)
    )

    """
    Boton Reload:
    Este boton ejecuta la primera y las posteriores recargas
    para ir ajustando los valores de entrada y de salida.
    """
    canvas.place(x=0, y=0)
    button_image_1 = ImageTk.PhotoImage(Image.open(relative_to_assets("button_1.png")))

    button_1 = Button(
        image=button_image_1,
        borderwidth=0,
        highlightthickness=0,
        command=lambda: safe_run(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6, entry_7, entry_8),
        relief="flat"
    )
    button_1.place(
        x=399.0,
        y=593.0,
        width=121.0,
        height=37.0
    )
    """
    Boton Load:
    Este boton ejecuta la carga final del archivo una vez 
    que se haya verificado que todos los datos y parametros 
    hayan sido introducidos de manera correcta
    """
    button_image_2 = ImageTk.PhotoImage(Image.open(relative_to_assets("button_2.png")))

    button_2 = Button(
        image=button_image_2,
        borderwidth=0,
        highlightthickness=0,
        command=lambda: generator.generate(output_file, entry_9.get()),
        relief="flat"
    )
    button_2.place(
        x=570.0,
        y=593.0,
        width=121.0,
        height=37.0
    )
    """
    Boton UPLOAD:
    Este boton carga el archivo Excel asi como su ruta.
    """
    button_image_3 = ImageTk.PhotoImage(Image.open(relative_to_assets("button_3.png")))

    button_3 = Button(
        image=button_image_3,
        borderwidth=0,
        highlightthickness=0,
        command=open_file_explorer,  # Llama a la función cuando se hace clic
        relief="flat"
    )

    button_3.place(
        x=54.0,
        y=274.0,
        width=131.0,
        height=127.0
    )
    """
    Boton DOWNLOAD:
    Este boton descarga el archivo Excel en su respectiva ruta.
    """
    button_image_4 = ImageTk.PhotoImage(Image.open(relative_to_assets("button_4.png")))

    button_4 = Button(
        image=button_image_4,
        borderwidth=0,
        highlightthickness=0,
        command=get_output_file,  # Llama a la función cuando se hace clic
        relief="flat"
    )

    button_4.place(
        x=915.0,
        y=279.0,
        width=131.0,
        height=127.0
    )
    """
    Se colocan los nombres de las descripciones
    de cada caja de texto y sus funcionalidades
    Func: EtiquetaTex -> TextSpace 
            -> EtiquetaRow -> TextSpace 
                -> EtiquetaCol -> TextSpace
    """
    canvas.create_text(
        376.0,
        70.0,
        anchor="nw",
        text="Carrera",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_1 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_1.place(
        x=376.0,
        y=88.0,
        width=491.0,
        height=33.0
    )
    canvas.create_text(
        233.0,
        70.0,
        anchor="nw",
        text="Row",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_10 = Entry(
        bd=0,
        bg="#FFFFFF",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_10.place(
        x=233.0,
        y=88.0,
        width=56.0,
        height=33.0
    )
    canvas.create_text(
        304.0,
        70.0,
        anchor="nw",
        text="Column",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_16 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_16.place(
        x=303.0,
        y=88.0,
        width=56.0,
        height=33.0
    )

    canvas.create_text(
        378.0,
        140.0,
        anchor="nw",
        text="Cuatrimestre",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_2 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_2.place(
        x=376.0,
        y=158.0,
        width=128.0,
        height=33.0
    )
    entry_11 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_11.place(
        x=233.0,
        y=158.0,
        width=56.0,
        height=33.0
    )
    entry_17 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_17.place(
        x=304.0,
        y=158.0,
        width=56.0,
        height=33.0
    )

    canvas.create_text(
        378.0,
        210.0,
        anchor="nw",
        text="Grupo",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_3 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_3.place(
        x=377.0,
        y=368.0,
        width=127.0,
        height=33.0
    )
    entry_12 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_12.place(
        x=233.0,
        y=228.0,
        width=56.0,
        height=33.0
    )
    entry_18 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_18.place(
        x=303.0,
        y=228.0,
        width=56.0,
        height=33.0
    )

    canvas.create_text(
        377.0,
        280.0,
        anchor="nw",
        text="No. Materias",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_4 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_4.place(
        x=377.0,
        y=439.0,
        width=127.0,
        height=33.0
    )
    entry_13 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_13.place(
        x=233.0,
        y=299.0,
        width=56.0,
        height=33.0
    )
    entry_19 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_19.place(
        x=304.0,
        y=298.0,
        width=56.0,
        height=33.0
    )
    canvas.create_text(
        378.0,
        349.0,
        anchor="nw",
        text="Primera Matricula",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_5 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_5.place(
        x=376.0,
        y=228.0,
        width=128.0,
        height=33.0
    )
    entry_14 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_14.place(
        x=235.0,
        y=369.0,
        width=56.0,
        height=33.0,
    )
    entry_20 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_20.place(
        x=305.0,
        y=369.0,
        width=56.0,
        height=33.0
    )

    canvas.create_text(
        378.0,
        422.0,
        anchor="nw",
        text="Ultima Matricula",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_6 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_6.place(
        x=376.0,
        y=298.0,
        width=128.0,
        height=33.0
    )
    entry_15 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_15.place(
        x=235.0,
        y=440.0,
        width=56.0,
        height=33.0
    )
    entry_21 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_21.place(
        x=306.0,
        y=439.0,
        width=56.0,
        height=33.0
    )
    canvas.create_text(
        233.0,
        506.0,
        anchor="nw",
        text="Catalogo Materias",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_7 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_7.place(
        x=233.0,
        y=524.0,
        width=131.0,
        height=33.0
    )
    canvas.create_text(
        378.0,
        506.0,
        anchor="nw",
        text="Fecha Asignacion",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_8 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_8.place(
        x=378.0,
        y=524.0,
        width=131.0,
        height=33.0
    )
    canvas.create_text(
        523.0,
        506.0,
        anchor="nw",
        text="Nombre del archivo",
        fill="#000000",
        font=("Inter Medium", 15 * -1)
    )
    entry_9 = Entry(
        bd=0,
        bg="#FFFFFF",
        fg="#000716",
        highlightthickness=2,
        highlightbackground="#606060"
    )
    entry_9.place(
        x=523.0,
        y=524.0,
        width=344.0,
        height=33.0
    )

    """
    Funciones básicas en la carga de la aplicación.

    Este bloque de código se encarga de configurar la ventana principal de la aplicación
    al momento de iniciar. Realiza tres acciones fundamentales:

    1. Desactiva la posibilidad de redimensionar la ventana (establece las dimensiones fijas).
    2. Centra la ventana en la pantalla utilizando la función `centrar_ventana`.
    3. Inicia el bucle principal de la interfaz gráfica de usuario (GUI) con `window.mainloop()`.

    Este conjunto de funciones es comúnmente utilizado para configurar la ventana principal 
    de una aplicación de escritorio con Tkinter.
    """

    # Desactiva la opción de redimensionar la ventana
    window.resizable(False, False)

    # Centra la ventana en la pantalla con las dimensiones proporcionadas
    centrar_ventana(window, WINDOW_LENGTH, WINDOW_HIGH)

    # Inicia el bucle de eventos principal de la ventana
    window.mainloop()


def relative_to_assets(path: str) -> Path:
    """
    Esta función toma una ruta de archivo relativa y la convierte en una ruta absoluta
    dentro del directorio de activos (ASSETS_PATH).

    Parámetros:
    path (str): La ruta relativa del archivo que se desea convertir a una ruta absoluta.

    Retorna:
    Path: La ruta absoluta correspondiente a la ruta de archivo proporcionada, concatenada
    con el directorio de activos.
    """
    return ASSETS_PATH / Path(path)


def open_file_explorer():
    """
    Esta función abre el explorador de archivos para que el usuario seleccione un archivo.
    La ruta del archivo seleccionado se guarda en la variable global `file_path`.

    Esta función utiliza `filedialog.askopenfilename()` para permitir al usuario seleccionar
    un archivo en su sistema.
    """
    global file_path  # Variable global para almacenar la ruta del archivo seleccionado
    file_path = filedialog.askopenfilename()  # Abre el explorador de archivos y guarda la ruta


def get_output_file():
    """
    Esta función abre un explorador de directorios para que el usuario seleccione una carpeta.
    La ruta del directorio seleccionado se guarda en la variable global `output_file`.

    Esta función utiliza `filedialog.askdirectory()` para permitir al usuario seleccionar
    un directorio en su sistema.
    """
    global output_file  # Variable global para almacenar la ruta del directorio seleccionado
    output_file = filedialog.askdirectory()  # Abre el explorador de directorios y guarda la ruta


def centrar_ventana(ventana, ancho, alto):
    """
    Esta función centra una ventana en la pantalla en función del ancho y alto proporcionados.

    Parámetros:
    ventana (tk.Tk): La ventana de la que se desea cambiar la posición en la pantalla.
    ancho (int): El ancho de la ventana.
    alto (int): El alto de la ventana.

    Esta función calcula las coordenadas x e y necesarias para centrar la ventana en la
    pantalla y ajusta la geometría de la ventana en consecuencia.
    """
    # Obtener el ancho y alto de la pantalla
    pantalla_ancho = ventana.winfo_screenwidth()
    pantalla_alto = ventana.winfo_screenheight()

    # Calcular las coordenadas x e y para centrar la ventana
    x = (pantalla_ancho // 2) - (ancho // 2)
    y = (pantalla_alto // 2) - (alto // 2)

    # Establecer la geometría de la ventana
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")  # Aplica la geometría para centrarla


"""
Error al ejecutar la funcion
Call in -> safe_run
#if get_params: get_params()
#excelFormat.set_params(default_params)
"""


def get_params():
    global default_params
    modificado = False

    col_1 = 10  # Valor inicial para la primera columna (entry_10)
    col_2 = 16  # Valor inicial para la segunda columna (entry_16)

    for fila in default_params:
        # Verificar que la fila tenga al menos dos elementos
        if len(fila) < 2:
            print(f"Error: La fila {fila} no tiene suficientes elementos.")
            continue  # Saltar esta fila

        entry_1_key = f"entry_{col_1}"  # entry_10, entry_11, etc.
        entry_2_key = f"entry_{col_2}"  # entry_16, entry_17, etc.

        # Verificar si las entradas existen en globals()
        if entry_1_key not in globals() or entry_2_key not in globals():
            print(f"Error: No se encontraron las entradas {entry_1_key} o {entry_2_key}.")
            col_1 += 1
            col_2 += 1
            continue  # Saltar esta fila

        entry_1 = globals()[entry_1_key]
        entry_2 = globals()[entry_2_key]

        nuevo_valor_1 = entry_1.get()
        nuevo_valor_2 = entry_2.get()

        # Solo actualizar si los valores no están vacíos
        if nuevo_valor_1 != "" and nuevo_valor_1 != str(fila[0]):
            try:
                fila[0] = int(nuevo_valor_1)  # Convertir a entero
                modificado = True
            except ValueError:
                print(f"Error: El valor {nuevo_valor_1} no es un número válido.")

        if nuevo_valor_2 != "" and nuevo_valor_2 != str(fila[1]):
            try:
                fila[1] = int(nuevo_valor_2)  # Convertir a entero
                modificado = True
            except ValueError:
                print(f"Error: El valor {nuevo_valor_2} no es un número válido.")

        col_1 += 1  # Incrementar para la siguiente entrada (entry_11, entry_12, etc.)
        col_2 += 1  # Incrementar para la siguiente entrada (entry_17, entry_18, etc.)

    return modificado


def safe_run(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6, entry_7, entry_8):
    """
    Esta función intenta ejecutar un proceso relacionado con la apertura de un archivo Excel y la
    visualización de información. Si el proceso tiene éxito, se muestra un mensaje de éxito.
    Si falla, se muestra un mensaje de error.

    Parámetros:
    entry_1, entry_2, entry_3, entry_4, entry_5, entry_6 (tk.Entry): Campos de entrada donde
    se mostrarán los datos del archivo procesado.
    entry_7, entry_8 (tk.Entry): Campos de entrada que contienen valores para la ejecución
    del proceso.

    Si la ruta del archivo `file_path` es válida y no está vacía, se ejecuta `excelFormat.run()`
    para procesar el archivo. Después se llama a `show_info()` para mostrar la información
    extraída del archivo. En caso de éxito, se muestra un mensaje informando que el archivo
    fue abierto correctamente.

    Si la ruta del archivo no está definida o si ocurre un error, se muestra un mensaje
    indicando que el archivo no fue abierto.
    """
    try:
        if file_path.strip():  # Verifica si la ruta del archivo no está vacía
            excelFormat.run(file_path, entry_8.get(), entry_7.get())  # Ejecuta el proceso de apertura del archivo
            show_info(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6)  # Muestra la información en las entradas
            messagebox.showinfo("Done", "¡The file was successfully opened!")  # Mensaje de éxito
        else:
            messagebox.showinfo("Fail", "¡The file was no Updated!")  # Mensaje si no se encuentra la ruta del archivo

    except Exception as e:
        # En caso de error, se muestra un mensaje con el detalle del error
        print("Error al ejecutar la función:", e)
        messagebox.showinfo("Fail", "¡The file was no opened!")  # Mensaje de error


def show_info(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6):
    """
    Esta función carga datos de un archivo JSON y los muestra en campos de entrada (entry).
    Los campos se limpian antes de insertar los nuevos valores.

    Parámetros:
    entry_1, entry_2, entry_3, entry_4, entry_5, entry_6 (tk.Entry): Campos de entrada donde
    se mostrarán los datos del archivo JSON.

    Este proceso lee un archivo JSON con información de un estudiante y luego inserta la
    información relevante en los campos de entrada correspondientes:
    - Carrera
    - Cuatrimestre
    - Matrícula (primera y segunda)
    - Grupo
    - Número de materias
    """
    with open(json_filename, "r", encoding="utf-8") as archivo:
        # Abre el archivo JSON y carga su contenido
        datos_json = json.load(archivo)

    # Limpia los campos de entrada antes de mostrar los nuevos datos
    entry_1.delete(0, tk.END)
    entry_2.delete(0, tk.END)
    entry_3.delete(0, tk.END)
    entry_4.delete(0, tk.END)
    entry_5.delete(0, tk.END)
    entry_6.delete(0, tk.END)

    # Ingresa los nuevos datos en los campos de entrada correspondientes
    entry_1.insert(0, datos_json["general"]["CARRERA"])  # Carrera
    entry_2.insert(0, datos_json["general"]["CUATRIMESTRE"])  # Cuatrimestre
    entry_3.insert(0, datos_json["ESTUDIANTES"][0]["MATRICULA"])  # Primera Matrícula
    entry_4.insert(0, datos_json["ESTUDIANTES"][-1]["MATRICULA"])  # Segunda Matrícula
    entry_5.insert(0, datos_json["general"]["GRUPO"])  # Grupo
    entry_6.insert(0, datos_json["general"]["NO_MATERIAS"])  # No. Materias
