import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import time
import win32gui
import win32con
from automation_engine import AutomationEngine

class AutoInputApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.engine = AutomationEngine()
        self.reset_state()
        self.setup_ui()

    def reset_state(self):
        self.running = False
        self.paused = False
        self.file_path = ""
        self.total_processed = 0
        self.current_total_lines = 0

    def setup_ui(self):
        self.title("AutoInput")
        self.geometry("550x720")
        
        ctk.CTkLabel(self, text="AutoInput", font=("Arial", 28, "bold")).pack(pady=20)
        
        # UI Elements
        self.btn_file = ctk.CTkButton(self, text="Seleccionar TXT", command=self.select_file)
        self.btn_file.pack(pady=10)

        self.window_list = ctk.CTkComboBox(self, values=self.engine.get_windows(), width=400)
        self.window_list.pack(pady=15)
        
        self.check_clear = ctk.CTkCheckBox(self, text="Limpiar celda/campo antes de escribir")
        self.check_clear.pack(pady=5)

        self.delay_label = ctk.CTkLabel(self, text="Velocidad: 1.5s")
        self.delay_label.pack(pady=(15,0))
        self.delay_slider = ctk.CTkSlider(self, from_=0.1, to=5.0, command=self.update_delay_label)
        self.delay_slider.set(1.5)
        self.delay_slider.pack(pady=10)

        # Control Group
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(pady=30)
        
        self.btn_start = ctk.CTkButton(self.btn_frame, text="Comenzar", fg_color="#27ae60", hover_color="#2ecc71", width=160, command=self.start_thread)
        self.btn_start.pack(side="left", padx=10)

        self.btn_stop = ctk.CTkButton(self.btn_frame, text="STOP", fg_color="#2c3e50", state="disabled", width=160, command=self.handle_stop)
        self.btn_stop.pack(side="left", padx=10)

        self.progress = ctk.CTkProgressBar(self, width=420)
        self.progress.pack(pady=10)
        self.progress.set(0)
        
        # Este es el label que actualizaremos
        self.status = ctk.CTkLabel(self, text="Estado: Listo", font=("Arial", 13), text_color="gray", wraplength=500)
        self.status.pack(pady=20)

    def update_delay_label(self, val):
        self.delay_label.configure(text=f"Delay después de Enter: {float(val):.1f}s")

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if self.file_path:
            try:
                # Leemos las líneas para contar cuántas hay (ignorando vacías)
                with open(self.file_path, 'r', encoding='utf-8') as f:
                    lines = [l.strip() for l in f if l.strip()]
                
                self.current_total_lines = len(lines)
                file_name = self.file_path.split('/')[-1]
                
                # Formato solicitado: Cargado: filename.txt | Se encontro xx items/lineas
                new_text = f"Cargado: {file_name} | Se encontró {self.current_total_lines} ítems/líneas"
                self.status.configure(text=new_text, text_color="blue")
                
            except Exception as e:
                self.status.configure(text=f"Error al leer archivo: {str(e)}", text_color="red")

    def handle_stop(self):
        if self.running and not self.paused:
            self.paused = True
            self.btn_start.configure(text="Continuar", state="normal")
            self.btn_stop.configure(text="Cancelar Escritura", fg_color="#c0392b")
            self.status.configure(text="Estado: Pausado", text_color="orange")
            self.minimize_target()
        elif self.paused:
            if messagebox.askyesno("Confirmar", f"¿Desea cancelar?\nProcesados: {self.total_processed}/{self.current_total_lines}"):
                self.running = False
                self.finalize_ui("cancelado")

    def start_thread(self):
        if self.paused:
            self.paused = False
            self.btn_start.configure(state="disabled")
            self.btn_stop.configure(text="STOP", fg_color="#e74c3c")
            self.status.configure(text="Estado: Escribiendo...", text_color="#27ae60")
            return

        if not self.file_path or self.current_total_lines == 0: 
            messagebox.showwarning("Error", "Por favor, carga un archivo .txt con contenido")
            return
            
        self.running = True
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal", fg_color="#e74c3c")
        threading.Thread(target=self.worker, daemon=True).start()

    def worker(self):
        try:
            target_title = self.window_list.get()
            is_excel = "Excel" in target_title
            parent_hwnd = win32gui.FindWindow(None, target_title)
            
            # Recargamos las líneas por si el usuario editó el archivo antes de dar START
            with open(self.file_path, 'r', encoding='utf-8') as f:
                lines = [l.strip() for l in f if l.strip()]
            
            self.current_total_lines = len(lines)
            self.engine.dock_window_right(parent_hwnd)

            for i, line in enumerate(lines):
                while self.paused and self.running: time.sleep(0.2)
                if not self.running: break

                if win32gui.IsIconic(parent_hwnd): 
                    self.engine.dock_window_right(parent_hwnd)
                
                self.status.configure(text=f"Escribiendo {i+1} de {self.current_total_lines}", text_color="#2ecc71")
                self.progress.set((i+1)/self.current_total_lines)

                child = win32gui.FindWindowEx(parent_hwnd, 0, "XLDESK", None)
                if child: child = win32gui.FindWindowEx(child, 0, "EXCEL7", None)
                final_hwnd = child if child else parent_hwnd

                self.engine.send_input(final_hwnd, line, self.check_clear.get(), is_excel)
                self.total_processed = i + 1
                time.sleep(self.delay_slider.get())

            if self.running: self.finalize_ui("exito")
        except Exception as e:
            print(f"Worker Error: {e}")
            self.finalize_ui("error")

    def finalize_ui(self, result):
        self.running = False
        self.paused = False
        self.btn_start.configure(text="Comenzar", state="normal")
        self.btn_stop.configure(text="STOP", state="disabled", fg_color="#2c3e50")
        self.minimize_target()
        
        if result == "exito":
            self.status.configure(text="Estado: Items escritos con éxito", text_color="#27ae60")
            messagebox.showinfo("Carga Finalizada", "Se han escrito todos los registros correctamente.")
        elif result == "cancelado":
            self.status.configure(text="Estado: Proceso Cancelado", text_color="#c0392b")
        else:
            self.status.configure(text="Estado: Error en la automatización", text_color="red")

    def minimize_target(self):
        hwnd = win32gui.FindWindow(None, self.window_list.get())
        if hwnd: win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)