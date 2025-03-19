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

OUTPUT_PATH = Path(__file__).parent
json_filename = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output", "data.json"))
ASSETS_PATH = OUTPUT_PATH / Path(r"../images")
WINDOW_LENGTH = 1100
WINDOW_HIGH = 700
file_path = ""

# ID | Fila | Column
default_params = [
    [10, 2],
    [12, 2],
    [14, 2],
    [17, 23],
    [20, 2],
    [0, 0]
]


def ex():
    window = Tk()
    window.geometry("1100x700")
    window.configure(bg="#FFFFFF")

    canvas = Canvas(
        window,
        bg="#FFFFFF",
        height=700,
        width=1100,
        bd=0,
        highlightthickness=0,
        relief="ridge"
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
        command=lambda: safe_run(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6, entry_7, entry_8, entry_9),
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
        command=lambda: generator.generate(),
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
        command=open_file_explorer,  # Llama a la función cuando se hace clic
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
    Funciones basicas en la carga de la app
    """
    window.resizable(False, False)
    centrar_ventana(window, WINDOW_LENGTH, WINDOW_HIGH)
    window.mainloop()


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)


def open_file_explorer():
    global file_path
    file_path = filedialog.askopenfilename()


def centrar_ventana(ventana, ancho, alto):
    # Obtener el ancho y alto de la pantalla
    pantalla_ancho = ventana.winfo_screenwidth()
    pantalla_alto = ventana.winfo_screenheight()

    # Calcular las coordenadas x e y para centrar la ventana
    x = (pantalla_ancho // 2) - (ancho // 2)
    y = (pantalla_alto // 2) - (alto // 2)

    # Establecer la geometría de la ventana
    ventana.geometry(f"{ancho}x{alto}+{x}+{y}")

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


def safe_run(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6, entry_7, entry_8, entry_9):
    try:

        if file_path.strip():
            excelFormat.run(file_path, entry_9.get(), entry_8.get(), entry_7.get())
            show_info(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6)
            messagebox.showinfo("Done", "¡The file was successfully opened!")
        else:
            messagebox.showinfo("Fail", "¡The file was no Updated!")

    except Exception as e:
        print("Error al ejecutar la función:", e)
        messagebox.showinfo("Fail", "¡The file was no opened!")


def show_info(entry_1, entry_2, entry_3, entry_4, entry_5, entry_6):
    with open(json_filename, "r", encoding="utf-8") as archivo:
        datos_json = json.load(archivo)

    """
    Limpia los espacios para los nuevos datos
    """
    entry_1.delete(0, tk.END)
    entry_2.delete(0, tk.END)
    entry_3.delete(0, tk.END)
    entry_4.delete(0, tk.END)
    entry_5.delete(0, tk.END)
    entry_6.delete(0, tk.END)

    """
    Ingresa los nuevos datos en sus posiciones
    """
    entry_1.insert(0, datos_json["general"]["CARRERA"])  # Carrera
    entry_2.insert(0, datos_json["general"]["CUATRIMESTRE"])  # Cuatrimestre
    entry_3.insert(0, datos_json["ESTUDIANTES"][0]["MATRICULA"])  # Primera Matricula
    entry_4.insert(0, datos_json["ESTUDIANTES"][-1]["MATRICULA"])  # Segunda Matricula
    entry_5.insert(0, datos_json["general"]["GRUPO"])  # Grupo
    entry_6.insert(0, datos_json["general"]["NO_MATERIAS"])  # No. Materias
