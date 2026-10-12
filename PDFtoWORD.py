import os
import threading
import subprocess
import platform
import customtkinter as ctk
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
from pdf2docx import Converter

# Integración de TkinterDnD con CustomTkinter
class CTkApp(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

ctk.set_appearance_mode("Dark")

class PDFToWordApp(CTkApp):
    def __init__(self):
        super().__init__()

        # --- Dimensiones y Posición (Abajo a la Derecha) ---
        width = 620
        height = 490

        self.title("PDF to WORD Converter")
        self.resizable(True, True)

        # Paleta de colores
        self.BG_COLOR = "#0f172a"      # Azul oscuro uniforme (Slate 900)
        self.CARD_BG = "#1e293b"       # Fondo de la tarjeta central (Slate 800)
        self.ACCENT_BLUE = "#38bdf8"   # Azul celeste neón

        self.configure(fg_color=self.BG_COLOR)

        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = screen_width - width - 30
        y = screen_height - height - 60
        self.geometry(f"{width}x{height}+{x}+{y}")

        self.pdf_path = ""
        self.converted_docx_path = ""
        self.is_converting = False

        # Configurar la ventana para recibir archivos al soltar
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.al_arrastrar_archivo)

        # --- INTERFAZ GRÁFICA ---

        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=self.BG_COLOR)
        self.main_frame.pack(fill="both", expand=True, padx=25, pady=20)

        # Header / Encabezado
        self.lbl_title = ctk.CTkLabel(
            self.main_frame, 
            text="📄 PDF to WORD", 
            font=("Segoe UI", 26, "bold"),
            text_color="#ffffff"
        )
        self.lbl_title.pack(pady=(15, 4))

        self.lbl_subtitle = ctk.CTkLabel(
            self.main_frame, 
            text="Arrastra tu PDF a la ventana o selecciónalo manualmente", 
            font=("Segoe UI", 13),
            text_color="#cbd5e1"
        )
        self.lbl_subtitle.pack(pady=(0, 20))

        # --- ZONA CENTRAL ---
        self.drop_zone = ctk.CTkFrame(
            self.main_frame, 
            corner_radius=16, 
            fg_color=self.CARD_BG, 
            border_width=2, 
            border_color="#334155"
        )
        self.drop_zone.pack(fill="x", padx=10, pady=5)

        self.drop_zone.bind("<Button-1>", lambda e: self.seleccionar_pdf())

        self.lbl_icon = ctk.CTkLabel(
            self.drop_zone, 
            text="📥", 
            font=("Segoe UI", 44)
        )
        self.lbl_icon.pack(pady=(22, 6))
        self.lbl_icon.bind("<Button-1>", lambda e: self.seleccionar_pdf())

        self.lbl_drop_text = ctk.CTkLabel(
            self.drop_zone, 
            text="Arrastra y suelta tu archivo PDF aquí", 
            font=("Segoe UI", 15, "bold"),
            text_color=self.ACCENT_BLUE
        )
        self.lbl_drop_text.pack(pady=(0, 4))
        self.lbl_drop_text.bind("<Button-1>", lambda e: self.seleccionar_pdf())

        self.lbl_archivo = ctk.CTkLabel(
            self.drop_zone, 
            text="O haz clic para explorar tus archivos", 
            font=("Segoe UI", 12),
            text_color="#94a3b8"
        )
        self.lbl_archivo.pack(pady=(0, 22))
        self.lbl_archivo.bind("<Button-1>", lambda e: self.seleccionar_pdf())

        # --- BOTONES DE ACCIÓN ---
        self.button_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.button_frame.pack(pady=22)

        self.btn_convert = ctk.CTkButton(
            self.button_frame, 
            text="🚀 Convertir Documento", 
            font=("Segoe UI", 14, "bold"),
            height=46,
            width=210,
            corner_radius=10,
            fg_color="#16a34a",
            text_color="#ffffff",
            hover_color="#15803d",
            command=self.iniciar_conversion
        )
        self.btn_convert.grid(row=0, column=0, padx=10)

        self.btn_open_folder = ctk.CTkButton(
            self.button_frame, 
            text="📂 Abrir Resultado", 
            font=("Segoe UI", 14, "bold"),
            height=46,
            width=180,
            corner_radius=10,
            fg_color="#0284c7",
            text_color="#ffffff",
            hover_color="#0369a1",
            command=self.abrir_carpeta
        )
        self.btn_open_folder.grid(row=0, column=1, padx=10)

        # --- BARRA DE PROGRESO Y ESTADO ---
        self.progress_bar = ctk.CTkProgressBar(
            self.main_frame, 
            height=6, 
            corner_radius=3,
            progress_color="#22c55e",
            fg_color="#334155"
        )

        self.lbl_status = ctk.CTkLabel(
            self.main_frame, 
            text="", 
            font=("Segoe UI", 13, "bold")
        )
        self.lbl_status.pack(pady=(5, 10))

    # --- LÓGICA DE CARGA Y CONVERSIÓN ---

    def al_arrastrar_archivo(self, event):
        archivos = self.tk.splitlist(event.data)
        if archivos:
            ruta = archivos[0]
            if ruta.lower().endswith(".pdf"):
                self.cargar_pdf(ruta)
            else:
                self.mostrar_estado("⚠️ Por favor, selecciona solo archivos .PDF", "#f87171")

    def seleccionar_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if file_path:
            self.cargar_pdf(file_path)

    def cargar_pdf(self, ruta):
        self.pdf_path = ruta
        nombre = os.path.basename(ruta)
        
        self.lbl_icon.configure(text="📄")
        self.drop_zone.configure(border_color="#22c55e", fg_color=self.CARD_BG)
        self.lbl_drop_text.configure(text=nombre, text_color="#4ade80")
        self.lbl_archivo.configure(
            text=f"Ruta: {ruta[:45]}..." if len(ruta) > 45 else ruta, 
            text_color="#cbd5e1"
        )
        self.mostrar_estado("Documento cargado correctamente", "#38bdf8")
        
        if self.progress_bar.winfo_ismapped():
            self.progress_bar.pack_forget()

    def iniciar_conversion(self):
        if not self.pdf_path:
            self.mostrar_estado("⚠️ Primero debes seleccionar o arrastrar un archivo PDF", "#f87171")
            return

        if self.is_converting:
            return

        self.is_converting = True
        self.drop_zone.configure(border_color="#facc15")
        self.mostrar_estado("Convirtiendo documento...", "#facc15")
        
        self.progress_bar.pack(fill="x", padx=20, pady=(0, 10))
        self.progress_bar.start()
        
        threading.Thread(target=self.procesar_conversion, daemon=True).start()

    def procesar_conversion(self):
        try:
            self.converted_docx_path = os.path.splitext(self.pdf_path)[0] + ".docx"
            
            cv = Converter(self.pdf_path)
            cv.convert(self.converted_docx_path, start=0, end=None)
            cv.close()

            self.progress_bar.stop()
            self.progress_bar.set(1.0)
            self.drop_zone.configure(border_color="#4ade80")
            self.mostrar_estado("¡Conversión completada con éxito!", "#4ade80")
        except Exception as e:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.drop_zone.configure(border_color="#f87171")
            self.mostrar_estado(f"Error: {str(e)}", "#f87171")
        finally:
            self.is_converting = False

    def abrir_carpeta(self):
        if not self.converted_docx_path or not os.path.exists(self.converted_docx_path):
            self.mostrar_estado("⚠️ Primero debes convertir un documento", "#f87171")
            return

        carpeta = os.path.dirname(self.converted_docx_path)
        if platform.system() == "Windows":
            os.startfile(carpeta)
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", carpeta])
        else:
            subprocess.Popen(["xdg-open", carpeta])

    def mostrar_estado(self, mensaje, color):
        self.lbl_status.configure(text=mensaje, text_color=color)

if __name__ == "__main__":
    app = PDFToWordApp()
    app.mainloop()