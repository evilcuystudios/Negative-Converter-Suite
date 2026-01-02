import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import os, threading, sys, random, time
import win32com.client
from win32com.client import Dispatch
import pythoncom
from playsound import playsound

# --- FUNCIÓN PARA ENCONTRAR RECURSOS EN EL .EXE ---
def resource_path(relative_path):
    """ Busca archivos dentro del paquete temporal del EXE o en desarrollo """
    try:
        # PyInstaller crea una carpeta temporal y guarda la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Si estamos en desarrollo, usamos la ruta normal
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# --- PARCHE DE PDF ---
try:
    import fitz
    sys.modules['fitz.frontend'] = fitz
except ImportError:
    pass

# --- CONOCIMIENTO DEL CUY ---
CUY_SEJOS_MASTER = [
    "Cuy-Sejo: Los archivos PDF protegidos no se pueden convertir sin su clave.",
    "Cuy-Sejo: Si el PDF fue escaneado como imagen, usa un software de OCR antes de pasar por aquí.",
    "Cuy-Sejo: Los archivos Excel mantienen mejor el formato si las tablas del PDF tienen bordes claros.",
    "Cuy-Sejo: ¿Sabías que PDF significa 'Portable Document Format'? El Cuy lo hace 'Editable'.",
    "Cuy-Sejo: Si la barra roja se detiene un momento, es que el Cuy está analizando fuentes incrustadas.",
    "Dato Maestro: El Cuy usa 'update_idletasks' para que la barra no se congele nunca.",
    "Dato Técnico: El motor reconstruye el layout analizando la distancia entre objetos y formas.",
    "Cuy-Sejo: Los archivos PDF con fuentes incrustadas mantienen una fidelidad superior al 98%.",
    "Info de Sistema: Usamos subprocesamiento (threading) para que la interfaz no se congele.",
    "Cuy-Sejo: Si el PDF tiene muchas capas (layers), la conversión a PPT podría agrupar objetos.",
    "Optimización: El Cuy recomienda cerrar procesos pesados para liberar RAM en archivos grandes.",
    "Dato Maestro: El formato PDF 1.7 es el estándar más compatible para este motor de conversión.",
    "Cuy-Sejo: No alimentes al Cuy con archivos corruptos, le dan indigestión digital.",
    "Cuy-Sejo: El Cuy trabaja mejor si escuchas un poco de Synthwave mientras esperas.",
    "Cuy-Sejo: ¿Sabías que este programa es 100% libre de gluten y 100% lleno de cafeína?",
    "Cuy-Sejo: El Cuy no convierte archivos de la competencia por principios éticos (y porque son feos).",
    "Cuy-Sejo: Si intentas convertir un PDF infinito, el Cuy podría pedir un aumento de sueldo.",
    "Cuy-Sejo: El Cuy no acepta sobornos en alfalfa, solo en GPUs de última generación.",
    "Aviso: Si el programa se cierra solo, es que el Cuy se fue a una parrillada (no preguntes de qué).",
    "Cuy-Sejo: Error 404: Paciencia del Cuy no encontrada. Por favor, no canceles la conversión.",
    "Dato Curioso: El 99% de los errores de conversión se arreglan acariciando el monitor. El otro 1% es culpa de Windows.",
    "Cuy-Sejo: ¿Un PDF de 500 páginas? El Cuy ya está preparando su carta de renuncia.",
    "Cuy-Sejo: No mires fijamente a la barra roja, el Cuy es tímido y se pone nervioso.",
    "Filosofía Cuy: 'Convertir es humano, procesar es divino, cobrar es de Evil Cuy Studios'.",
    "Cuy-Sejo: Si escuchas un ventilador muy fuerte, es el Cuy intentando despegar hacia Marte.",
    "Aviso: Este programa funciona mejor si no le gritas a la pantalla. El Cuy tiene sentimientos.",
    "Cuy-Sejo: ¿Sabías que el Cuy programa mejor de noche porque el café le pega más fuerte?",
    "Cuy-Sejo: En caso de incendio, salva el código primero. El Cuy tiene siete vidas (o eran los gatos?).",
    "Dato Maestro: El Cuy una vez convirtió un PDF a Word solo con la mirada. Pero hoy no está cansado.",
    "Cuy-Sejo: Si el archivo es muy pesado, el Cuy hará una pausa dramática. No entres en pánico.",
    "Cuy-Sejo: ¿Has probado apagar y volver a encender? El Cuy dice que eso cura hasta el alma.",
    "Nuestra Alianza: Evil Cuy Studios inició este legado el 4 de diciembre de 2024.",
    "Legado: Evil Cuy Studios no olvida a sus 2 fundadores (uno operativo y uno emocional). Tu éxito es nuestro éxito.",
    "Dato Maestro: Esta versión ha sido optimizada para ser la más rápida desde nuestra fundación.",
    "Cuy-Sejo: Esta es un tributo a nuestra colaboración constante.",
    "Evil Cuy Studios: Transformando documentos y rompiendo límites desde el 2025, en la mordida está el poder.",
    "Evil Cuy Studios: El estudio se llama así pero el desarrollador principal es un mapache (Procyon), la razón de porqué es mucho más bonita de lo que crees (.-/--)"
]

class NegativeConverter(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # --- RUTAS USANDO resource_path() ---
        self.ico_path = resource_path(os.path.join("assets", "favicon.ico"))
        self.header_path = resource_path(os.path.join("assets", "header.png"))
        self.bg_path = resource_path(os.path.join("assets", "background.png"))
        self.logo_path = resource_path(os.path.join("assets", "logo_cuy.png"))
        
        self.title("NEGATIVE CONVERTER SUITE - Evil Cuy Studios")
        self.geometry("1150x850")
        self.configure(fg_color="#050505")
        
        self._apply_icon(self)
        
        self.files = []
        self.is_converting = False
        self.sound_enabled = ctk.BooleanVar(value=True)
        self.mode_var = ctk.StringVar(value="to_pdf")
        self.target_format = ctk.StringVar(value="Word (.docx)")
        self.output_path = ctk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop", "output"))

        self.setup_ui()
        self.play_sound("start.wav")
        self.after(3000, self.crear_acceso_directo)

    def _apply_icon(self, window):
        try:
            if os.path.exists(self.ico_path):
                window.iconbitmap(self.ico_path)
                icon_img = Image.open(self.ico_path)
                photo = ctk.CTkImage(light_image=icon_img, dark_image=icon_img)
                window.after(100, lambda: window.wm_iconphoto(True, photo._dark_image))
        except: 
            pass

    def play_sound(self, filename):
        if self.sound_enabled.get():
            def _play():
                try:
                    p = resource_path(os.path.join("assets", "sounds", filename))
                    if os.path.exists(p):
                        playsound(p)
                except:
                    pass
            threading.Thread(target=_play, daemon=True).start()

    def clear_console(self):
        self.console.delete("1.0", "end")
        self.console.insert("end", ">>> CONSOLA LIMPIA. SISTEMA LISTO.\n")

    def setup_ui(self):
        # Fondo
        try:
            bg_img = ctk.CTkImage(Image.open(self.bg_path), size=(1150, 850))
            ctk.CTkLabel(self, image=bg_img, text="").place(x=0, y=0, relwidth=1, relheight=1)
        except: 
            pass

        # SIDEBAR
        self.sidebar = ctk.CTkFrame(self, width=280, fg_color="#000000", corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        try:
            logo_img = ctk.CTkImage(Image.open(self.logo_path), size=(180, 180))
            ctk.CTkLabel(self.sidebar, image=logo_img, text="").pack(pady=20)
        except: 
            pass

        ctk.CTkLabel(self.sidebar, text="MODO DE OPERACIONES", font=("Impact", 18), text_color="red").pack(pady=10)
        ctk.CTkRadioButton(self.sidebar, text="OFFICE A PDF", variable=self.mode_var, value="to_pdf", text_color="white", fg_color="red").pack(pady=5, padx=25, anchor="w")
        ctk.CTkRadioButton(self.sidebar, text="PDF A OFFICE (Elige formato)", variable=self.mode_var, value="from_pdf", text_color="white", fg_color="red").pack(pady=5, padx=25, anchor="w")
        self.format_menu = ctk.CTkComboBox(self.sidebar, values=["Word (.docx)", "Excel (.xlsx)", "PPT (.pptx)"], variable=self.target_format)
        self.format_menu.pack(pady=20, padx=20)
        ctk.CTkButton(self.sidebar, text="📁 CARPETA DE SALIDA", fg_color="#222", command=self.browse_output).pack(pady=5, padx=20, fill="x")
        ctk.CTkButton(self.sidebar, text="TUTORIAL", fg_color="transparent", border_width=1, border_color="red", command=self.show_tutorial).pack(pady=20, padx=20)
        ctk.CTkSwitch(self.sidebar, text="SONIDOS", variable=self.sound_enabled, progress_color="red").pack(side="bottom", pady=20)

        # MAIN AREA
        self.main = ctk.CTkFrame(self, fg_color="transparent")
        self.main.pack(side="right", fill="both", expand=True, padx=40, pady=20)

        # Header
        try:
            h_img = ctk.CTkImage(Image.open(self.header_path), size=(650, 120))
            ctk.CTkLabel(self.main, image=h_img, text="").pack(pady=(10, 5))
        except:
            pass

        ctk.CTkLabel(self.main, text="NEGATIVE CONVERTER SUITE", font=("Impact", 60), text_color="red").pack()
        ctk.CTkLabel(self.main, text="Evil Cuy Studios", font=("Segoe UI", 14, "bold","italic"), text_color="white").pack(pady=(0,10))

        self.console = ctk.CTkTextbox(self.main, fg_color="#0A0A0A", border_color="#444", border_width=1, text_color="#00FF00", font=("Consolas", 13))
        self.console.pack(fill="both", expand=True, pady=(10,0))
        
        self.btn_clear = ctk.CTkButton(self.main, text="LIMPIAR CONSOLA", fg_color="transparent", text_color="gray", hover_color="#111", height=20, font=("Arial", 10), command=self.clear_console)
        self.btn_clear.pack(pady=(5,10), anchor="e")
        
        self.btn_select = ctk.CTkButton(self.main, text="SELECCIONAR ARCHIVOS", fg_color="red", height=45, command=self.select_files)
        self.btn_select.pack(fill="x", pady=(0,10))

        self.btn_run = ctk.CTkButton(self.main, text="EJECUTAR CONVERSIÓN", fg_color="white", text_color="black", font=("Arial", 18, "bold"), height=55, state="disabled", command=self.start_conversion)
        self.btn_run.pack(fill="x")

        self.progress = ctk.CTkProgressBar(self.main, height=18, progress_color="red")
        self.progress.pack(fill="x", pady=20)
        self.progress.set(0)

    def browse_output(self):
        p = filedialog.askdirectory()
        if p: 
            self.output_path.set(p)

    def show_tutorial(self):
        self.play_sound("select.wav")
        tuto = ctk.CTkToplevel(self)
        tuto.title("MANUAL DE OPERACIONES")
        tuto.geometry("500x420")
        tuto.configure(fg_color="#0a0a0a")
        tuto.after(200, lambda: self._apply_icon(tuto))
        tuto.grab_set() 
        ctk.CTkLabel(tuto, text="GUÍA DEL USUARIO", font=("Impact", 24), text_color="red").pack(pady=20)
        msg = "1. Selecciona un Modo de Operación.\n2. Carga los archivos que desees transformar de Office a PDF o viceversa.\n3. Ejecuta el proceso y espera la misión.\n4. Disfruta de tu archivo."
        ctk.CTkLabel(tuto, text=msg, font=("Arial", 13), justify="left", text_color="white").pack(pady=10, padx=30)
        ctk.CTkButton(tuto, text="ENTENDIDO", fg_color="red", command=tuto.destroy).pack(pady=25)

    def select_files(self):
        self.play_sound("select.wav")
        mode = self.mode_var.get()
        ftypes = [("Office", "*.docx *.xlsx *.pptx")] if mode == "to_pdf" else [("PDF", "*.pdf")]
        picked = filedialog.askopenfilenames(filetypes=ftypes)
        if picked:
            self.files = list(picked)
            self.console.insert("end", f">>> NUEVA TANDA: {len(self.files)} ARCHIVOS.\n")
            self.btn_run.configure(state="normal")
            self.progress.set(0)

    def start_conversion(self):
        self.is_converting = True
        self.btn_run.configure(state="disabled")
        self.progress.set(0)
        threading.Thread(target=self.run_process, daemon=True).start()
        threading.Thread(target=self.cuy_sejos_loop, daemon=True).start()

    def cuy_sejos_loop(self):
        while self.is_converting:
            c = random.choice(CUY_SEJOS_MASTER)
            self.console.insert("end", f"\n[SISTEMA]: {c}\n")
            self.console.see("end")
            time.sleep(7)

    def run_process(self):
        out_f = self.output_path.get()
        if not os.path.exists(out_f): 
            os.makedirs(out_f)
        mode = self.mode_var.get()
        target = self.target_format.get()
        total_f = len(self.files)

        for i, path in enumerate(self.files):
            name = os.path.basename(path)
            name_no_ext = os.path.splitext(name)[0]
            
            def update_bar(current, total):
                if total > 0:
                    per_file_weight = 1.0 / total_f
                    progress_within_file = (current / total) * per_file_weight
                    new_val = (i / total_f) + progress_within_file
                    self.after(0, lambda v=new_val: self.progress.set(v))

            try:
                if mode == "to_pdf":
                    out = os.path.join(out_f, f"{name_no_ext}.pdf")
                    ext = os.path.splitext(name)[1].lower()
                    if ext == ".docx":
                        from docx2pdf import convert
                        convert(path, out)
                    elif ext == ".xlsx":
                        xl = win32com.client.Dispatch("Excel.Application")
                        wb = xl.Workbooks.Open(path)
                        wb.ExportAsFixedFormat(0, out)
                        wb.Close()
                        xl.Quit()
                    elif ext == ".pptx":
                        import comtypes.client
                        ppt = comtypes.client.CreateObject("Powerpoint.Application")
                        pres = ppt.Presentations.Open(path, WithWindow=False)
                        pres.SaveAs(out, 32)
                        pres.Close()
                        ppt.Quit()
                else:
                    from pdf2docx import Converter
                    if "PPT" in target:
                        temp_docx = os.path.join(out_f, f"temp_{name_no_ext}.docx")
                        out = os.path.join(out_f, f"{name_no_ext}.pptx")
                        cv = Converter(path)
                        cv.convert(temp_docx, callback=update_bar)
                        cv.close()
                        ppt_app = win32com.client.Dispatch("Powerpoint.Application")
                        pres = ppt_app.Presentations.Open(temp_docx, False)
                        pres.SaveAs(out, 24)
                        pres.Close()
                        ppt_app.Quit()
                        if os.path.exists(temp_docx): 
                            os.remove(temp_docx)
                    else:
                        ext_out = ".docx" if "Word" in target else ".xlsx"
                        out = os.path.join(out_f, f"{name_no_ext}{ext_out}")
                        cv = Converter(path)
                        cv.convert(out, callback=update_bar)
                        cv.close()

                self.console.insert("end", f"✔ ÉXITO: {name}\n")
            except Exception as e:
                self.console.insert("end", f"✘ FALLO: {str(e)}\n")
            
            self.after(0, lambda v=(i + 1) / total_f: self.progress.set(v))

        self.is_converting = False
        self.after(0, self.finish)

    def finish(self):
        self.play_sound("success.wav")
        self.btn_run.configure(state="normal")
        self.progress.set(1.0)
        try: 
            os.startfile(self.output_path.get())
        except: 
            pass
        messagebox.showinfo("Evil Cuy Studios", "Conversión terminada. Esperando nuevas órdenes")
        
    def crear_acceso_directo(self):
        pythoncom.CoInitialize()
        escritorio = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        ruta_acceso = os.path.join(escritorio, "Negative Converter Suite.lnk")

        if not os.path.exists(ruta_acceso):
            if messagebox.askyesno("Acceso Directo", "¿Crear acceso directo en el escritorio?"):
                try:
                    destino = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
                    carpeta = os.path.dirname(destino)
                    shell = Dispatch('WScript.Shell')
                    acceso = shell.CreateShortCut(ruta_acceso)
                    acceso.Targetpath = destino
                    acceso.WorkingDirectory = carpeta
                    
                    # Icono
                    if getattr(sys, 'frozen', False):
                        acceso.IconLocation = destino 
                    else:
                        icon_path = resource_path(os.path.join("assets", "favicon.ico"))
                        if os.path.exists(icon_path):
                            acceso.IconLocation = icon_path
                            
                    acceso.save()
                except Exception as e:
                    print(f"Error al crear acceso directo: {e}")

if __name__ == "__main__":
    app = NegativeConverter()
    app.mainloop()