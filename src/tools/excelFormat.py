import os
import pandas as pd
import json

# Datos finales, no modificables
TIPO_EVALUACION = "Ordinaria"
PROFESOR_NUMERO_CONTROL = ""

# Entradas de datos
root = r""
json_filename = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output", "data.json"))

# Datos dinamicos
num_materias = 0
fecha_asignacion = ""
clave_materia = ""
filename = ""
set_rc = []


def load_excel_file():
    """
    Carga un archivo Excel y lo devuelve como un DataFrame.
    Determina el motor (openpyxl o xlrd) dependiendo de la extensión del archivo.
    """
    try:
        file_extension = os.path.splitext(root)[1].lower()  # Obtener la extensión del archivo

        if file_extension == '.xls':  # Usar xlrd para archivos .xls
            df = pd.read_excel(root, engine='xlrd', header=None)
        elif file_extension == '.xlsx':  # Usar openpyxl para archivos .xlsx
            df = pd.read_excel(root, engine='openpyxl', header=None)
        else:
            print(f"Formato de archivo no soportado: {file_extension}")
            return None

        return df
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return None

def extract_materias(dataFrame):
    """
    Extrae las materias desde el archivo Excel.
    Las materias están en la fila 18, comenzando en la columna F (índice 5).
    """
    global num_materias

    if dataFrame is None:
        print("No se pudo cargar el archivo.")
        num_materias = 0
        return None

    try:
        materias = []
        columna_actual = 5  # Columna F (índice 5)

        while columna_actual < dataFrame.shape[1]:  # Evitar índice fuera de rango
            materia = dataFrame.iloc[17, columna_actual]  # Fila 18 (índice 17)

            if pd.isna(materia):  # Si la celda está vacía, detenerse
                break

            materias.append(str(materia))
            columna_actual += 3  # Saltar 3 columnas (F, I, L, O, R, U, etc.)

        num_materias = len(materias)  # Actualizar la variable global
        return materias  # Retorna la lista de materias encontradas

    except Exception as e:
        print(f"Error al extraer las materias: {e}")
        num_materias = 0
        return None

def extract_cuatrimestre_cursado(dataFrame):
    codigo_periodo_escolar = generate_codigo_periodo_escolar(dataFrame)
    parts = codigo_periodo_escolar.split("-")
    return int(parts[1])

def extract_nivel_educativo(dataFrame):
    cuatrimestre = extract_cuatrimestre(dataFrame, 13, 3)
    return int(cuatrimestre[(len(cuatrimestre)) - 1:])

def extract_matriculas_and_calificaciones(dataFrame, start_row=20, matricula_column='B'):
    """
    Extrae las matrículas y sus calificaciones asociadas desde el archivo Excel.
    """
    if dataFrame is None:
        print("No se pudo cargar el archivo.")
        return None, None, None, None

    # Convertir la letra de la columna en un índice numérico para las matrículas
    matricula_col_index = ord(matricula_column.upper()) - ord('A')

    # Extraer las matrículas desde la fila `start_row` hasta la última con datos
    matriculas = dataFrame.iloc[start_row - 1:, matricula_col_index].dropna().astype(str).tolist()

    # Extraer las calificaciones para cada materia (PROM, EGO, CDEF)
    cdef = []  # Calificación Definitiva (CDEF)

    for i in range(start_row - 1, len(dataFrame)):
        row = dataFrame.iloc[i]
        cdef_estudiante = []

        columna_actual = 5  # Comenzar en la columna F (índice 5)
        while True:
            if columna_actual + 2 >= len(row):  # Si no hay más columnas, terminar
                break

            # Extraer CDEF, reemplazando NaN por 0
            cdef_value = row[columna_actual + 2] if not pd.isna(row[columna_actual + 2]) else 0

            cdef_estudiante.append(cdef_value)

            columna_actual += 3  # Saltar 3 columnas (F, G, H -> I, J, K -> etc.)

        cdef.append(cdef_estudiante)

    return matriculas, cdef

def extract_cuatrimestre(dataFrame, fila, columna):
    cuatrimestre = dataFrame.iloc[fila, columna]
    return cuatrimestre

def extract_grupo(dataFrame, fila, columna):
    grupo = dataFrame.iloc[fila, columna]
    return grupo

def generate_codigo_periodo_escolar(dataFrame):
    cuatrimestre = extract_cuatrimestre(dataFrame, 13, 3)
    grupo = extract_grupo(dataFrame, 15, 3)

    anioCurso = cuatrimestre[0:4]
    periodo = int(grupo[4:5])
    if (periodo >= 1 and periodo <= 6):
        nivel_educativo = "lic"
    else:
        nivel_educativo = "ing"
    codigo_periodo_escolar = anioCurso + "-" + str(periodo) + "-" + nivel_educativo
    return codigo_periodo_escolar

def extract_carrera(dataFrame, fila, columna):
    carrera = dataFrame.iloc[fila, columna]
    return carrera

def create_json(matriculas, cdef, materias, dataFrame):
    """
    Crea un objeto JSON
    """
    data = {
        "general": {
            "CUATRIMESTRE": extract_cuatrimestre(dataFrame, 13, 3),
            "GRUPO": extract_grupo(dataFrame, 13, 3),
            "CARRERA": extract_carrera(dataFrame, 11, 2),
            "CODIGO_PERIODO_ESCOLAR": generate_codigo_periodo_escolar(dataFrame),
            "FECHA_ASIGNACION": fecha_asignacion,
            "CUATRIMESTRE_CURSADO": extract_cuatrimestre_cursado(dataFrame),
            "NIVEL_EDUCATIVO": extract_nivel_educativo(dataFrame),
            "PROFESOR_NUMERO_CONTROL": PROFESOR_NUMERO_CONTROL,
            "TIPO_EVALUACION": TIPO_EVALUACION,
            "NO_MATERIAS": num_materias,
        },
        "ESTUDIANTES": []
    }

    for i in range(len(matriculas)):
        estudiante = {
            "MATRICULA": matriculas[i],
            "MATERIAS": []
        }

        for j in range(len(materias)):
            estudiante["MATERIAS"].append({
                "MATERIA": materias[j],
                "CDEF": cdef[i][j],  # Calificación Definitiva
                "CLAVE_MATERIA": clave_materia
            })

        data["ESTUDIANTES"].append(estudiante)

    return data

def save_json_to_file(json_data, filename):
    """
    Guarda el objeto JSON en un archivo.
    """
    try:
        # Crear el directorio si no existe
        os.makedirs(os.path.dirname(filename), exist_ok=True)

        # Guardar el archivo JSON
        with open(filename, 'w', encoding='utf-8') as json_file:
            json.dump(json_data, json_file, indent=4, ensure_ascii=False)
        print(f"Archivo guardado exitosamente en: {filename}")
    except Exception as e:
        print(f"Error al guardar el archivo JSON: {e}")

def getInfo(url, fileName, asignationDate, subjectId):
    global root
    global filename
    global fecha_asignacion
    global clave_materia
    root = url
    filename = fileName
    fecha_asignacion = asignationDate
    clave_materia = subjectId

def set_params(rc_params):
    global set_rc
    try:
        set_rc = rc_params
        print(set_rc)
    except Exception as e:
        print(f"Error en el parametro de set_params {e}")

def run(url, fileName, asignationDate, subjectId):
    """
    Inicia el programa obteniendo el archivo Excel
    """
    getInfo(url, fileName, asignationDate, subjectId)
    dataFrame = load_excel_file()
    materias = extract_materias(dataFrame)
    matriculas, cdef = extract_matriculas_and_calificaciones(dataFrame)
    json_data = create_json(matriculas, cdef, materias, dataFrame)
    save_json_to_file(json_data, json_filename)
