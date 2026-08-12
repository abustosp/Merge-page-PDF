import pdfplumber
import os
import sys
from tkinter.filedialog import askdirectory
import re
from io import StringIO

patron_cuit = r'(\d{11})'

directorio = askdirectory(title="Selecciona una carpeta con PDFs")

# Listar todos los archivos del directorio y subdirectorios y devolver una lista de los archivos

archivos = []

for root, dirs, files in os.walk(directorio):
    for file in files:
        if file.endswith(".pdf"):
            archivos.append(os.path.join(root, file))
            
# filtrar los PDF
archivos_pdf = [archivo for archivo in archivos if archivo.endswith(".pdf")]
                
for i in archivos_pdf:
    old_stderr = sys.stderr
    sys.stderr = StringIO()
    try:
        with pdfplumber.open(i) as pdf:
            page = pdf.pages[0]
            text = page.extract_text()
            
            # encontrar el primer CUIT en el texto
            cuit_contenido = re.search(patron_cuit, text)
            if cuit_contenido:
                cuit_contenido = cuit_contenido.group(0)
            
            # obtener solamente el nombre del archivo de i
            i = os.path.basename(i)
            
            # encontrar el cuit en el nombre del archivo
            cuit_nombre = re.search(patron_cuit, i)
            
            if cuit_nombre:
                cuit_nombre = cuit_nombre.group(0)
            
            # comparar el cuit del archivo con el cuit del nombre del archivo
            if cuit_contenido and cuit_nombre:
                if cuit_contenido == cuit_nombre:
                    print(f"El CUIT {cuit_contenido} coincide con el nombre del archivo.")
                else:
                    print(f"XXXXXXXXXXXXXXXXXXXXXX\nEl CUIT no coincide con el nombre del archivo: {i}.\nXXXXXXXXXXXXXXXXXXXXXX")
            else:
                print(f"XXXXXXXXXXXXXXXXXXXXXX\nNo se encontró CUIT en el contenido o en el nombre del archivo: {i}.\nXXXXXXXXXXXXXXXXXXXXXX")
    finally:
        sys.stderr = old_stderr
            

