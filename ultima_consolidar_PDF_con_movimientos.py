import os
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from tkinter.filedialog import askdirectory
import io
import re
from tkinter.filedialog import asksaveasfilename
import logging

logging.getLogger("pdfminer").setLevel(logging.ERROR)

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


def normalize_cropboxes(pdf_reader):
    """Completa CropBox faltante usando MediaBox para evitar warnings de PyPDF2."""
    for page in pdf_reader.pages:
        if "/CropBox" not in page:
            page.cropbox = page.mediabox


def pdf_tiene_movimientos(pdf_path, pdf_reader):
    """Verifica si un PDF tiene movimientos (transacciones) usando pdfplumber o PyPDF2."""
    patron_exclusión = re.compile(
        r"TOTAL[^\n]*\$0[.,]00(?:\s*\$0[.,]00)+"
        r"|SIN\s*MOVIMIENTO",
        re.IGNORECASE,
    )

    # Intentar con pdfplumber primero (mejor extracción de texto)
    if HAS_PDFPLUMBER:
        try:
            texto_completo = ""
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    texto_completo += (page.extract_text() or "") + "\n"
            if not texto_completo.strip():
                return False
            if patron_exclusión.search(texto_completo):
                return False
            return True
        except Exception:
            pass  # Fallback a PyPDF2

    # Fallback: PyPDF2 en todas las páginas
    texto_completo = ""
    for page in pdf_reader.pages:
        texto_completo += (page.extract_text() or "") + "\n"
    if not texto_completo.strip():
        return False
    if patron_exclusión.search(texto_completo):
        return False
    return True


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
        normalize_cropboxes(pdf_reader)

        # Ignorar PDFs sin páginas
        if len(pdf_reader.pages) == 0:
            continue

        number_of_pages = len(pdf_reader.pages) - 1

        # Verificar si tiene movimientos
        if not pdf_tiene_movimientos(pdf, pdf_reader):
            continue

        # Agregar la última página del archivo al merger
        nombre = os.path.splitext(os.path.basename(pdf))[0]
        merger.append(pdf_reader, pages=(number_of_pages, (number_of_pages + 1)), outline_item=nombre)
        merged_files.append(pdf)

# Escribir el archivo PDF resultante en memoria
output = io.BytesIO()
merger.write(output)
output.seek(0)

# Preguntar nombre de archivo para guardar

nombre_archivo = asksaveasfilename(
    initialdir=Carpeta,
    title='Guardar archivo consolidado',
    initialfile='Consolidado última Hoja.pdf',
    defaultextension='.pdf',
    filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')]
)

if nombre_archivo:
    with open(nombre_archivo, 'wb') as fout:
        fout.write(output.read())
    
# Guardar el TXT en la misma carpeta donde se guardó el consolidado
txt_path = os.path.splitext(nombre_archivo)[0] + '_Archivos_Procesados.txt'
with open(txt_path, 'w') as f:
    for pdf in merged_files:
        f.write(str(pdf) + '\n')