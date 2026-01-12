import os
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from tkinter.filedialog import askdirectory
import io
import pdfplumber
import re

patron_exclusión = r"(TOTAL\n\$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00)"

# Preguntar por la ruta de la carpeta
Carpeta = askdirectory(title='Seleccionar carpeta')
os.chdir(Carpeta)

# Obtener todos los archivos PDF de la carpeta Folder y sus subcarpetas
pdfFiles = []
for foldername, subfolders, filenames in os.walk(Carpeta):
    for filename in filenames:
        if filename.endswith('.pdf'):
            pdfFiles.append(os.path.join(foldername, filename))

# Ordenar los archivos alfabéticamente sin tener en cuenta el path
pdfFiles.sort(key=os.path.basename)

# Crear un objeto PdfMerger
merger = PdfMerger()

# Agregar la última página de cada archivo PDF al merger
merged_files = []
for pdf in pdfFiles:
    with open(pdf, 'rb') as f:
        pdf_reader = PdfReader(f)
        number_of_pages = len(pdf_reader.pages) - 1
        
        # Leer el contenido de la primera página con pdfplumber
        with pdfplumber.open(f) as pdfp:
            primera_pagina = pdfp.pages[0]
            texto = primera_pagina.extract_text()
            
            # Excluir los archivos que contienen el patrón de exclusión
            if re.search(patron_exclusión, texto):
                continue
        
        # Agregar la última página del archivo al merger
        merger.append(pdf_reader, pages=(number_of_pages, (number_of_pages + 1)))
        merged_files.append(pdf)

# Escribir el archivo PDF resultante en memoria
output = io.BytesIO()
merger.write(output)
output.seek(0)

# Guardar el archivo PDF en la carpeta seleccionada
with open(os.path.join(Carpeta, 'Consolidado última Hoja.pdf'), 'wb') as fout:
    fout.write(output.read())
    
with open(os.path.join(Carpeta, 'Archivos Procesados.txt'), 'w') as f:
    for pdf in merged_files:
        f.write(str(pdf) + '\n')