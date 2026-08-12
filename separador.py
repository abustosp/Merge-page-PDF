import os
from PyPDF2 import PdfReader, PdfWriter
from tkinter.filedialog import askdirectory
import shutil


def normalize_cropboxes(pdf_reader):
    """Completa CropBox faltante usando MediaBox para evitar warnings de PyPDF2."""
    for page in pdf_reader.pages:
        if "/CropBox" not in page:
            page.cropbox = page.mediabox

def extract_pdf_pages():
    # Seleccionar carpeta
    folder_path = askdirectory(title="Selecciona una carpeta")
    if not folder_path:
        print("No se seleccionó ninguna carpeta.")
        return

    # Crear carpeta "Separado" dentro de la carpeta seleccionada
    separated_folder = os.path.join(folder_path, "Separado")
    os.makedirs(separated_folder, exist_ok=True)

    # Recorrer archivos en la carpeta seleccionada
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            pdf_reader = PdfReader(pdf_path)
            normalize_cropboxes(pdf_reader)
            num_pages = len(pdf_reader.pages)

            # Extraer cada página y guardarla como un nuevo PDF
            for page_num in range(num_pages):
                pdf_writer = PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[page_num])

                output_filename = f"{os.path.splitext(filename)[0]}_separado_página_{page_num + 1}.pdf"
                output_path = os.path.join(separated_folder, output_filename)

                with open(output_path, "wb") as output_pdf:
                    pdf_writer.write(output_pdf)

    print("Páginas extraídas y guardadas en la carpeta 'Separado'.")

if __name__ == "__main__":
    extract_pdf_pages()
