import pdfplumber
import os
from tkinter.filedialog import askdirectory

directorio = askdirectory(title="Selecciona una carpeta con PDFs")

# Listar todos los archivos del directorio y subdirectorios y devolver una lista de los archivos
archivos = []

for root, dirs, files in os.walk(directorio):
    for file in files:
        if file.endswith(".pdf"):
            archivos.append(os.path.join(root, file))

# Filtrar los PDF
archivos_pdf = [archivo for archivo in archivos if archivo.endswith(".pdf")]

print(f"Analizando {len(archivos_pdf)} archivos PDF...\n")

archivos_con_inconsistencia = []

for archivo in archivos_pdf:
    try:
        with pdfplumber.open(archivo) as pdf:
            texto_completo = ""
            # Extraer texto de todas las páginas
            for page in pdf.pages:
                texto_completo += page.extract_text() or ""
            
            # Buscar la palabra INCONSISTENCIA
            if "INCONSISTENCIA" in texto_completo:
                nombre_archivo = os.path.basename(archivo)
                archivos_con_inconsistencia.append(nombre_archivo)
                print(f"⚠️  INCONSISTENCIA encontrada en: {nombre_archivo}")
    except Exception as e:
        nombre_archivo = os.path.basename(archivo)
        print(f"❌ Error al procesar {nombre_archivo}: {e}")

print(f"\n{'='*60}")
print(f"Resumen:")
print(f"Total de archivos analizados: {len(archivos_pdf)}")
print(f"Archivos con INCONSISTENCIA: {len(archivos_con_inconsistencia)}")
print(f"{'='*60}")

if archivos_con_inconsistencia:
    print("\nArchivos con inconsistencias:")
    for archivo in archivos_con_inconsistencia:
        print(f"  - {archivo}")
