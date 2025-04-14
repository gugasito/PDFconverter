from pptx import Presentation
from pptx.util import Inches, Pt
import os
import fitz  # PyMuPDF
from tkinter import Tk, Button, Label
from tkinter.filedialog import askopenfilename, asksaveasfilename

def pdf_to_pptx(pdf_path, pptx_path):
    # Verificar si el archivo PDF existe
    if not os.path.exists(pdf_path):
        print(f"El archivo {pdf_path} no existe.")
        return

    # Crear una presentación PPTX
    presentation = Presentation()

    # Abrir el archivo PDF
    pdf_document = fitz.open(pdf_path)

    # Ajustar el tamaño de la diapositiva al tamaño del PDF
    first_page = pdf_document[0]
    pdf_width, pdf_height = first_page.rect.width, first_page.rect.height
    presentation.slide_width = Pt(pdf_width)
    presentation.slide_height = Pt(pdf_height)

    for page_number in range(len(pdf_document)):
        # Obtener la página del PDF
        page = pdf_document[page_number]

        # Renderizar la página como imagen
        pix = page.get_pixmap(dpi=150)  # DPI ajustable
        image_path = f"page_{page_number + 1}.png"
        pix.save(image_path)

        # Crear una diapositiva en el PPTX
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])  # Layout vacío

        # Agregar la imagen a la diapositiva
        left = Inches(0)
        top = Inches(0)
        slide.shapes.add_picture(image_path, left, top, Pt(pdf_width), Pt(pdf_height))

        # Eliminar la imagen temporal
        os.remove(image_path)

    # Guardar el archivo PPTX
    presentation.save(pptx_path)
    print(f"Archivo PPTX guardado en: {pptx_path}")

def open_file_selection():
    # Ocultar la ventana principal de Tkinter
    root.withdraw()

    # Abrir ventana para seleccionar el archivo PDF
    pdf_path = askopenfilename(
        title="Selecciona el archivo PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )

    if not pdf_path:
        print("No se seleccionó ningún archivo PDF.")
    else:
        # Abrir ventana para seleccionar la ruta de guardado del PPTX
        pptx_path = asksaveasfilename(
            title="Guardar como",
            defaultextension=".pptx",
            filetypes=[("Archivos PPTX", "*.pptx")]
        )

        if not pptx_path:
            print("No se seleccionó una ruta para guardar el archivo PPTX.")
        else:
            pdf_to_pptx(pdf_path, pptx_path)

if __name__ == "__main__":
    # Crear la ventana principal
    root = Tk()
    root.title("PDF a PPTX Converter")
    root.geometry("300x150")
    root.iconbitmap("logo.ico")  # Cambia "icon.ico" por la ruta de tu icono

    # Etiqueta de bienvenida
    label = Label(root, text="Bienvenido al convertidor PDF a PPTX")
    label.pack(pady=10)

    # Botón para iniciar la selección de archivos
    start_button = Button(root, text="Seleccionar PDF", command=open_file_selection)
    start_button.pack(pady=20)

    # Iniciar el bucle principal de la ventana
    root.mainloop()