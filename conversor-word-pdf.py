import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import threading
from docx2pdf import convert
import os
from PIL import Image, ImageTk

def seleccionar_archivo():
    ruta_docx = filedialog.askopenfilename(title="Seleccionar archivo Word",filetypes=[("Archivos Word", "*.docx")],)
    entrada_ruta.delete(0, tk.END)
    entrada_ruta.insert(0, ruta_docx)
    etiqueta_estado.config(text="")

def convertir_a_pdf():
    archivo_docx = entrada_ruta.get()
    if archivo_docx:
        archivo_pdf = archivo_docx.replace(".docx", ".pdf")
        hilo = threading.Thread(target=convertir_en_segundo_plano, args=(archivo_docx, archivo_pdf))
        hilo.start()
    else:
        etiqueta_estado.config(text="Selecciona un archivo Word antes de convertir.")
        boton_descargar.config(state=tk.DISABLED)
        boton_abrir.config(state=tk.DISABLED)

def convertir_en_segundo_plano(archivo_docx, archivo_pdf):
    try:
        convert(archivo_docx, archivo_pdf)
        etiqueta_estado.config(text="Conversión completada", foreground="green")
        etiqueta_ruta_pdf.config(text=f"Ruta del archivo PDF: {archivo_pdf}", foreground="green")
        boton_descargar.config(state=tk.NORMAL)
        boton_abrir.config(state=tk.NORMAL)
    except Exception as e:
        etiqueta_estado.config(text=f"Error durante la conversión: {str(e)}", foreground="red")
        boton_descargar.config(state=tk.DISABLED)
        boton_abrir.config(state=tk.DISABLED)

def descargar_pdf():
    archivo_pdf = entrada_ruta.get().replace(".docx", ".pdf")
    etiqueta_estado.config(text="Descargando...", foreground="blue")
    ventana.update()
    ventana.after(2000, lambda: actualizar_interfaz_despues_descarga(archivo_pdf))

def actualizar_interfaz_despues_descarga(archivo_pdf):
    etiqueta_estado.config(text="Descarga completada", foreground="blue")
    boton_descargar.pack_forget()  
    etiqueta_ruta_pdf.config(text=f"Ruta del archivo PDF: {archivo_pdf}", foreground="green")
    boton_abrir.pack(pady=20)  

def abrir_pdf():
    archivo_pdf = etiqueta_ruta_pdf.cget("text").replace("Ruta del archivo PDF: ", "")
    if os.path.exists(archivo_pdf):
        os.system(f'start "" "{archivo_pdf}"') 
    else:
        etiqueta_estado.config(text="El archivo PDF no se encuentra.", foreground="blue")

ventana = tk.Tk()
ventana.title("SISTEMA DE CONVERSIÓN DE WORD A PDF")
ventana.geometry("600x400")
etiqueta_estilo = ("Helvetica", 12, "bold")
entrada_estilo = ("Helvetica", 10)
boton_estilo = ("Helvetica", 12, "bold")
etiqueta_estado_estilo = ("Helvetica", 10, "italic")

def generar_degradado(ancho, alto):
    imagen = Image.new("RGB", (ancho, alto), color="white")
    for y in range(alto):
        color = (
            int(173 + (255 - 173) * (y / alto)),
            int(216 + (255 - 216) * (y / alto)),
            int(230 + (255 - 230) * (y / alto))
        )
        for x in range(ancho):
            imagen.putpixel((x, y), color)
    return imagen

imagen_degradado = generar_degradado(600, 400)
imagen_tk = ImageTk.PhotoImage(imagen_degradado)

fondo = tk.Label(ventana, image=imagen_tk)
fondo.place(relwidth=1, relheight=1)

# Widgets
etiqueta_ruta = tk.Label(ventana, text="RUTA DEL ARCHIVO WORD :", font=etiqueta_estilo, bg="#F0F0F0")
entrada_ruta = tk.Entry(ventana, width=50, font=entrada_estilo)
boton_seleccionar = tk.Button(ventana, text="SELECCIONAR ARCHIVO", command=seleccionar_archivo, font=boton_estilo, bg="#4CAF50", fg="white")
boton_convertir = tk.Button(ventana, text="CONVERTIR A PDF", command=convertir_a_pdf, font=boton_estilo, bg="#008CBA", fg="white")
etiqueta_estado = tk.Label(ventana, text="", font=etiqueta_estado_estilo, bg="#F0F0F0")
boton_descargar = tk.Button(ventana, text="DESCARGAR PDF", state=tk.DISABLED, command=descargar_pdf, font=boton_estilo, bg="#FF9800", fg="blue")
etiqueta_ruta_pdf = tk.Label(ventana, text="", font=etiqueta_estado_estilo, bg="#F0F0F0")
boton_abrir = tk.Button(ventana, text="ABRIR PDF", state=tk.DISABLED, command=abrir_pdf, font=boton_estilo, bg="#4CAF50", fg="blue")

# Alineación y disposición
etiqueta_ruta.pack(pady=10)
entrada_ruta.pack(pady=10)
boton_seleccionar.pack(pady=10)
boton_convertir.pack(pady=20)
etiqueta_estado.pack()
boton_descargar.pack(pady=10)
etiqueta_ruta_pdf.pack(pady=10)
boton_abrir.pack(pady=20)

# Iniciar la interfaz
ventana.mainloop()
