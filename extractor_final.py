import os
import glob
from docx import Document
from docx.shared import RGBColor
import pandas as pd

def extraer():
    def extraer_texto_rojo(doc_path):
        doc = Document(doc_path)
        red_texts = []
        current_text = ""

        for para in doc.paragraphs:
            for run in para.runs:
                if run.font.color and run.font.color.rgb == RGBColor(255, 0, 0):
                    current_text += run.text
                else:
                    if current_text:
                        red_texts.append(current_text)
                        current_text = ""
            if current_text:
                red_texts.append(current_text)
                current_text = ""

        return red_texts

    def procesar_documentos_en_carpeta(input_folder):
        all_red_texts = []

        # Obtener todos los archivos .docx en la carpeta
        for doc_file in glob.glob(os.path.join(input_folder, '*.docx')):
            red_texts = extraer_texto_rojo(doc_file)
            all_red_texts.extend(red_texts)
            print(f"Textos rojos en {doc_file}: {red_texts}")

        return all_red_texts

    # Carpeta de entrada
    input_folder = 'INPUT'  # Ajusta esta ruta según tu configuración

    # Ejecutar el proceso
    red_texts = procesar_documentos_en_carpeta(input_folder)

    # Crear un DataFrame con los textos extraídos
    df = pd.DataFrame(red_texts, columns=['PROBLEMA'])

    # Exportar el DataFrame a un archivo Excel
    output_path = 'extraccion.xlsx'
    df.to_excel(output_path, index=False)
