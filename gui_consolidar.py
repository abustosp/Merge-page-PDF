import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import io
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import pdfplumber
import re

# Colores del tema
BG = "#2e2e2e"
FG = "#ffffff"
ACCENT = "#d35400"


class ConsolidadorPDFGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Consolidador de PDFs")
        
        # Intentar cargar el icono
        try:
            self.root.iconbitmap(os.path.join("bin", "ABP-blanco-en-fondo-negro.ico"))
        except Exception:
            pass
        
        self.root.geometry("650x600")
        self.root.configure(background=BG)
        self.root.resizable(False, False)
        
        # Configurar estilo
        self.setup_styles()
        
        # Crear interfaz
        self.create_widgets()
        
    def setup_styles(self):
        """Configurar estilos de la aplicación"""
        style = ttk.Style(self.root)
        try:
            style.theme_use('clam')
        except Exception:
            pass
        
        # Estilos generales
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=FG)
        style.configure("TButton", foreground="#000000", padding=10, font=('Arial', 10))
        style.configure("Action.TButton", font=('Arial', 11, 'bold'), padding=12, justify='center')
        style.configure("TLabelframe", background=BG, foreground=FG)
        style.configure("TLabelframe.Label", background=BG, foreground=FG, font=("Arial", 11, "bold"))
        
    def create_widgets(self):
        """Crear los widgets de la interfaz"""
        # Configurar grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Container principal
        container = ttk.Frame(self.root, padding=10)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        
        # Header con título y logo
        header = ttk.Frame(container, padding=10)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        # Cargar logo si existe
        logo_path = os.path.join("bin", "MrBot.png")
        self.root.logo_img = None
        if os.path.exists(logo_path):
            try:
                self.root.logo_img = tk.PhotoImage(file=logo_path)
                tk.Label(header, image=self.root.logo_img, background=BG).pack(side="top", pady=(0, 8))
            except Exception:
                self.root.logo_img = None
        
        title_lbl = ttk.Label(
            header,
            text="CONSOLIDADOR DE PDFs",
            font=("Arial", 18, "bold"),
            foreground=FG,
            background=BG,
        )
        title_lbl.pack(anchor="center")
        
        subtitle = ttk.Label(
            header,
            text="Papeles de Trabajo y Libros de Compras y Ventas",
            font=("Arial", 11),
            foreground=FG,
            background=BG,
        )
        subtitle.pack(anchor="center", pady=(5, 0))
        
        # Frame de acciones
        actions_frame = ttk.LabelFrame(container, text="Seleccione el tipo de consolidación", padding=15)
        actions_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        actions_frame.columnconfigure(0, weight=1)
        
        # Botón 1: Consolidación Total
        btn_total = ttk.Button(
            actions_frame,
            text="Consolidación Total\n(Todos los PDFs completos)",
            style='Action.TButton',
            command=self.consolidacion_total,
            width=40
        )
        btn_total.grid(row=0, column=0, pady=8, sticky="ew")
        
        # Botón 2: Consolidación Última Hoja
        btn_ultima = ttk.Button(
            actions_frame,
            text="Consolidación Última Hoja\n(Solo última página)",
            style='Action.TButton',
            command=self.consolidacion_ultima_hoja,
            width=40
        )
        btn_ultima.grid(row=1, column=0, pady=8, sticky="ew")
        
        # Botón 3: Consolidación Última Hoja con Movimientos
        btn_movimientos = ttk.Button(
            actions_frame,
            text="Consolidación con Movimientos\n(Solo PDFs con transacciones)",
            style='Action.TButton',
            command=self.consolidacion_ultima_con_movimientos,
            width=40
        )
        btn_movimientos.grid(row=2, column=0, pady=8, sticky="ew")
        
        # Frame de información
        info_frame = ttk.LabelFrame(container, text="Información", padding=15)
        info_frame.grid(row=2, column=0, sticky="ew")
        
        info_text = (
            "• Cada opción solicitará la carpeta con los PDFs\n"
            "• Podrá seleccionar dónde guardar el resultado\n"
            "• Se generará un archivo TXT con la lista de archivos procesados"
        )
        
        info_label = ttk.Label(
            info_frame,
            text=info_text,
            font=('Arial', 9),
            foreground="#cccccc",
            background=BG,
            justify=tk.LEFT
        )
        info_label.pack(anchor="w")
        
    def consolidacion_total(self):
        """Consolidar todos los PDFs completos"""
        try:
            # Seleccionar carpeta
            folder = filedialog.askdirectory(title="Seleccionar carpeta con PDFs")
            if not folder:
                return
            
            os.chdir(folder)
            
            # Obtener todos los archivos PDF
            pdfFiles = []
            for foldername, subfolders, filenames in os.walk(folder):
                for filename in filenames:
                    if filename.endswith('.pdf'):
                        pdfFiles.append(os.path.join(foldername, filename))
            
            if not pdfFiles:
                messagebox.showwarning("Advertencia", "No se encontraron archivos PDF en la carpeta seleccionada.")
                return
            
            # Ordenar archivos
            pdfFiles.sort(key=os.path.basename)
            
            # Crear merger
            pdfMerger = PdfMerger()
            Primer_página = 0
            
            # Añadir cada PDF
            for filename in pdfFiles:
                with open(filename, 'rb') as f:
                    pdf_reader = PdfReader(f)
                    Number_of_pages = len(pdf_reader.pages) - 1
                    pdfMerger.append(pdf_reader, pages=(0, Number_of_pages + 1))
                    pdfMerger.add_outline_item(title=str(os.path.basename(filename)), pagenum=Primer_página)
                    Primer_página += Number_of_pages + 1
            
            # Guardar archivo
            nombre_archivo = filedialog.asksaveasfilename(
                initialdir=folder,
                title='Guardar archivo consolidado',
                initialfile='Consolidado Total.pdf',
                defaultextension='.pdf',
                filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')]
            )
            
            if nombre_archivo:
                with open(nombre_archivo, 'wb') as f:
                    pdfMerger.write(f)
                
                # Guardar TXT con archivos procesados
                txt_path = os.path.splitext(nombre_archivo)[0] + '_Archivos_Procesados.txt'
                with open(txt_path, 'w') as f:
                    for pdf in pdfFiles:
                        f.write(str(pdf) + '\n')
                
                messagebox.showinfo(
                    "Éxito", 
                    f"Consolidación completada.\n\nArchivos procesados: {len(pdfFiles)}\nGuardado en:\n{nombre_archivo}"
                )
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al consolidar: {str(e)}")
    
    def consolidacion_ultima_hoja(self):
        """Consolidar solo la última hoja de cada PDF"""
        try:
            # Seleccionar carpeta
            Carpeta = filedialog.askdirectory(title='Seleccionar carpeta con PDFs')
            if not Carpeta:
                return
            
            os.chdir(Carpeta)
            
            # Obtener todos los archivos PDF
            pdfFiles = []
            for foldername, subfolders, filenames in os.walk(Carpeta):
                for filename in filenames:
                    if filename.endswith('.pdf'):
                        pdfFiles.append(os.path.join(foldername, filename))
            
            if not pdfFiles:
                messagebox.showwarning("Advertencia", "No se encontraron archivos PDF en la carpeta seleccionada.")
                return
            
            # Ordenar archivos
            pdfFiles.sort(key=os.path.basename)
            
            # Crear merger
            merger = PdfMerger()
            
            # Agregar última página de cada PDF
            for pdf in pdfFiles:
                with open(pdf, 'rb') as f:
                    pdf_reader = PdfReader(f)
                    number_of_pages = len(pdf_reader.pages) - 1
                    merger.append(pdf_reader, pages=(number_of_pages, (number_of_pages + 1)))
            
            # Escribir en memoria
            output = io.BytesIO()
            merger.write(output)
            output.seek(0)
            
            # Guardar archivo
            nombre_archivo = filedialog.asksaveasfilename(
                initialdir=Carpeta,
                title='Guardar archivo consolidado',
                initialfile='Consolidado última Hoja.pdf',
                defaultextension='.pdf',
                filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')]
            )
            
            if nombre_archivo:
                with open(nombre_archivo, 'wb') as fout:
                    fout.write(output.read())
                
                # Guardar TXT con archivos procesados
                txt_path = os.path.splitext(nombre_archivo)[0] + '_Archivos_Procesados.txt'
                with open(txt_path, 'w') as f:
                    for pdf in pdfFiles:
                        f.write(str(pdf) + '\n')
                
                messagebox.showinfo(
                    "Éxito", 
                    f"Consolidación completada.\n\nArchivos procesados: {len(pdfFiles)}\nGuardado en:\n{nombre_archivo}"
                )
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al consolidar: {str(e)}")
    
    def consolidacion_ultima_con_movimientos(self):
        """Consolidar última hoja solo de PDFs con movimientos"""
        try:
            # Seleccionar carpeta
            Carpeta = filedialog.askdirectory(title='Seleccionar carpeta con PDFs')
            if not Carpeta:
                return
            
            os.chdir(Carpeta)
            
            # Patrón para excluir PDFs sin movimientos
            patron_exclusión = r"(TOTAL\n\$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00 \$0\.00)"
            
            # Obtener todos los archivos PDF
            pdfFiles = []
            for foldername, subfolders, filenames in os.walk(Carpeta):
                for filename in filenames:
                    if filename.endswith('.pdf'):
                        pdfFiles.append(os.path.join(foldername, filename))
            
            if not pdfFiles:
                messagebox.showwarning("Advertencia", "No se encontraron archivos PDF en la carpeta seleccionada.")
                return
            
            # Ordenar archivos
            pdfFiles.sort(key=os.path.basename)
            
            # Crear merger
            merger = PdfMerger()
            merged_files = []
            
            # Agregar última página de cada PDF con movimientos
            for pdf in pdfFiles:
                with open(pdf, 'rb') as f:
                    pdf_reader = PdfReader(f)
                    number_of_pages = len(pdf_reader.pages) - 1
                    
                    # Leer primera página para verificar si tiene movimientos
                    with pdfplumber.open(f) as pdfp:
                        primera_pagina = pdfp.pages[0]
                        texto = primera_pagina.extract_text()
                        
                        # Excluir si no tiene movimientos
                        if re.search(patron_exclusión, texto):
                            continue
                    
                    # Agregar última página
                    merger.append(pdf_reader, pages=(number_of_pages, (number_of_pages + 1)))
                    merged_files.append(pdf)
            
            if not merged_files:
                messagebox.showinfo("Información", "No se encontraron PDFs con movimientos.")
                return
            
            # Escribir en memoria
            output = io.BytesIO()
            merger.write(output)
            output.seek(0)
            
            # Guardar archivo
            nombre_archivo = filedialog.asksaveasfilename(
                initialdir=Carpeta,
                title='Guardar archivo consolidado',
                initialfile='Consolidado última Hoja con Movimientos.pdf',
                defaultextension='.pdf',
                filetypes=[('PDF files', '*.pdf'), ('All files', '*.*')]
            )
            
            if nombre_archivo:
                with open(nombre_archivo, 'wb') as fout:
                    fout.write(output.read())
                
                # Guardar TXT con archivos procesados
                txt_path = os.path.splitext(nombre_archivo)[0] + '_Archivos_Procesados.txt'
                with open(txt_path, 'w') as f:
                    for pdf in merged_files:
                        f.write(str(pdf) + '\n')
                
                messagebox.showinfo(
                    "Éxito", 
                    f"Consolidación completada.\n\nArchivos procesados: {len(merged_files)} de {len(pdfFiles)}\n(Se excluyeron {len(pdfFiles) - len(merged_files)} sin movimientos)\n\nGuardado en:\n{nombre_archivo}"
                )
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al consolidar: {str(e)}")


def main():
    root = tk.Tk()
    app = ConsolidadorPDFGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
