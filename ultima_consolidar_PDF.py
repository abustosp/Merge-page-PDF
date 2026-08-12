import os
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from tkinter.filedialog import askdirectory
import io


def normalize_cropboxes(pdf_reader):
    """Completa CropBox faltante usando MediaBox para evitar warnings de PyPDF2."""
    for page in pdf_reader.pages:
        if "/CropBox" not in page:
            page.cropbox = page.mediabox

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
for pdf in pdfFiles:
    with open(pdf, 'rb') as f:
        pdf_reader = PdfReader(f)
        normalize_cropboxes(pdf_reader)
        number_of_pages = len(pdf_reader.pages) - 1
        # Agregar la última página del archivo al merger
        nombre = os.path.splitext(os.path.basename(pdf))[0]
        merger.append(pdf_reader, pages=(number_of_pages, (number_of_pages + 1)), outline_item=nombre)

# Escribir el archivo PDF resultante en memoria
output = io.BytesIO()
merger.write(output)
output.seek(0)

# Guardar el archivo PDF en la carpeta seleccionada
with open(os.path.join(Carpeta, 'Consolidado última Hoja.pdf'), 'wb') as fout:
    fout.write(output.read())