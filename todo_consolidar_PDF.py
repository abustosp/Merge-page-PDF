import PyPDF2, os
from tkinter.filedialog import askdirectory
from PyPDF2 import PdfReader, PdfMerger

# Seleccionar la carpeta donde se encuentran los archivos PDF
folder = askdirectory() 
os.chdir(folder) 

# Obtener todos los archivos PDF de la carpeta Folder y sus subcarpetas
pdfFiles = []
for foldername, subfolders, filenames in os.walk(folder):
    for filename in filenames:
        if filename.endswith('.pdf'):
            pdfFiles.append(os.path.join(foldername, filename))

#pdfFiles.sort(key=str.lower) # Ordenar los archivos alfabéticamente 

#Ordenar los archivos alfabéticamente sin tener en cuenta el path
pdfFiles.sort(key=os.path.basename)

 # Crear el objeto PDF merger 
pdfMerger = PdfMerger() 

#Contador de páginas para los marcadores de página
Primer_página = 0

# Añadir cada archivo PDF al objeto merger 
for filename in pdfFiles: 
    with open(filename, 'rb') as f: 
        #Crear un objeto PDF reader para cada archivo PDF
        pdf_reader = PdfReader(f)
        #Obtener el número de páginas del archivo PDF
        Number_of_pages = len(pdf_reader.pages) -1
        #Agregar un marcador de página para cada archivo PDF sin el path
        pdfMerger.append(pdf_reader , pages=(0, Number_of_pages + 1 ))
        #Agregar un marcador de página para cada archivo PDF con el path
        pdfMerger.add_outline_item(title=str(os.path.basename(filename)) , pagenum=Primer_página)
        #Actualizar el contador de páginas
        Primer_página += Number_of_pages +1
    
 # Escribir el contenido del merger en un nuevo archivo PDF 								  
with open('Consolidado Total.pdf', 'wb') as f: 
    pdfMerger.write(f)