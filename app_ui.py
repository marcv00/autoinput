import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import time
import win32gui
import win32con
from automation_engine import AutomationEngine
from PIL import Image
import sys
import os


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


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
        self.start_time = None

    def setup_ui(self):

        self.title("AutoInput Workflows")
        self.geometry("550x720")
        self.minsize(550, 720)

        # ===== TITLE =====

        ctk.CTkLabel(
            self,
            text="AutoInput Workflows",
            font=("Arial", 28, "bold")
        ).pack(pady=(40, 20))

        # ===============================
        # FILE SECTION
        # ===============================

        self.file_section = ctk.CTkFrame(self, fg_color="transparent")
        self.file_section.pack(fill="x")

        self.file_header = ctk.CTkFrame(self.file_section, fg_color="transparent")
        self.file_header.pack(fill="x", padx=40)

        ctk.CTkLabel(
            self.file_header,
            text="Archivo de entrada",
            font=("Arial", 16, "bold")
        ).pack(side="left")

        self.btn_file = ctk.CTkButton(
            self.file_header,
            text="Seleccionar TXT",
            width=140,
            command=self.select_file
        )
        self.btn_file.pack(side="right")

        ctk.CTkFrame(self.file_section, height=2).pack(fill="x", padx=40, pady=(10, 15))

        self.file_info = ctk.CTkLabel(
            self.file_section,
            text="Ningún archivo seleccionado",
            text_color="gray"
        )
        self.file_info.pack()

        # ===============================
        # WINDOW SECTION
        # ===============================

        self.window_section = ctk.CTkFrame(self, fg_color="transparent")
        self.window_section.pack(fill="x")

        ctk.CTkLabel(
            self.window_section,
            text="Ventana objetivo",
            font=("Arial", 16, "bold")
        ).pack(pady=(25, 5))

        ctk.CTkFrame(self.window_section, height=2).pack(fill="x", padx=40, pady=(0, 15))

        self.window_frame = ctk.CTkFrame(self.window_section, fg_color="transparent")
        self.window_frame.pack()

        self.window_list = ctk.CTkComboBox(
            self.window_frame,
            values=self.engine.get_windows(),
            width=360
        )
        self.window_list.pack(side="left", padx=(0, 8))

        refresh_icon = ctk.CTkImage(
            light_image=Image.open(resource_path("refresh.png")),
            dark_image=Image.open(resource_path("refresh.png")),
            size=(18, 18)
        )

        self.btn_refresh = ctk.CTkButton(
            self.window_frame,
            text="",
            width=40,
            height=40,
            fg_color="#2979FF",
            hover_color="#1565C0",
            image=refresh_icon,
            command=self.refresh_windows,
            corner_radius=20
        )
        self.btn_refresh.pack(side="left")

        # ===============================
        # OPTIONS
        # ===============================

        self.options_section = ctk.CTkFrame(self, fg_color="transparent")
        self.options_section.pack(fill="x")

        ctk.CTkLabel(
            self.options_section,
            text="Opciones",
            font=("Arial", 16, "bold")
        ).pack(pady=(25, 5))

        ctk.CTkFrame(self.options_section, height=2).pack(fill="x", padx=40, pady=(0, 15))

        self.check_clear = ctk.CTkCheckBox(
            self.options_section,
            text="Limpiar campo antes del siguiente item (Recomendado)"
        )
        self.check_clear.pack(pady=5)
        self.check_clear.select()

        self.delay_label = ctk.CTkLabel(
            self.options_section,
            text="Delay después de Enter: 0.7s"
        )
        self.delay_label.pack(pady=(15, 0))

        self.delay_slider = ctk.CTkSlider(
            self.options_section,
            from_=0.7,
            to=5.0,
            command=self.update_delay_label
        )
        self.delay_slider.set(0.7)
        self.delay_slider.pack(pady=10)

        # ===============================
        # CONTROL
        # ===============================

        self.control_section = ctk.CTkFrame(self, fg_color="transparent")
        self.control_section.pack(fill="x")

        ctk.CTkLabel(
            self.control_section,
            text="Control",
            font=("Arial", 16, "bold")
        ).pack(pady=(25, 5))

        ctk.CTkFrame(self.control_section, height=2).pack(fill="x", padx=40, pady=(0, 15))

        self.btn_frame = ctk.CTkFrame(self.control_section, fg_color="transparent")
        self.btn_frame.pack(pady=20)

        self.btn_start = ctk.CTkButton(
            self.btn_frame,
            text="Comenzar",
            width=160,
            fg_color="#27ae60",
            hover_color="#2ecc71",
            command=self.start_thread
        )
        self.btn_start.pack(side="left", padx=10)

        self.btn_stop = ctk.CTkButton(
            self.btn_frame,
            text="STOP",
            width=160,
            state="disabled",
            fg_color="#2c3e50",
            command=self.handle_stop
        )
        self.btn_stop.pack(side="left", padx=10)

        # ===============================
        # PROGRESS (hidden initially)
        # ===============================

        self.progress_section = ctk.CTkFrame(self, fg_color="transparent")

        ctk.CTkLabel(
            self.progress_section,
            text="Progreso",
            font=("Arial", 16, "bold")
        ).pack(pady=(20, 5))

        ctk.CTkFrame(self.progress_section, height=2).pack(fill="x", padx=40, pady=(0, 15))

        self.progress = ctk.CTkProgressBar(self.progress_section, width=420)
        self.progress.pack(pady=10)
        self.progress.set(0)

        self.status = ctk.CTkLabel(
            self.progress_section,
            text="Estado: Listo",
            text_color="gray",
            wraplength=480
        )
        self.status.pack(pady=10)

        self.runtime_label = ctk.CTkLabel(
            self.progress_section,
            text="",
            font=("Arial", 13, "bold")
        )
        self.runtime_label.pack(pady=5)

        self.btn_back = ctk.CTkButton(
            self.progress_section,
            text="< Volver",
            width=120,
            command=self.reset_to_main
        )

        # ===============================
        # INSTRUCTIONS
        # ===============================

        self.instructions_panel = ctk.CTkFrame(self, corner_radius=10)

        self.instructions_label = ctk.CTkLabel(
            self.instructions_panel,
            text="",
            font=("Arial", 18, "bold"),
            justify="center",
            wraplength=420
        )

        self.instructions_label.pack(padx=20, pady=20)

    def update_delay_label(self, val):
        self.delay_label.configure(
            text=f"Delay después de Enter: {float(val):.1f}s"
        )

    def select_file(self):

        self.file_path = filedialog.askopenfilename(
            filetypes=[("Text files", "*.txt")]
        )

        if self.file_path:

            with open(self.file_path, 'r', encoding='utf-8') as f:
                lines = [l.strip() for l in f if l.strip()]

            self.current_total_lines = len(lines)

            file_name = os.path.basename(self.file_path)

            self.file_info.configure(
                text=f"{file_name}  |  {self.current_total_lines} líneas",
                text_color="blue"
            )

    def start_thread(self):

        if not self.file_path:
            messagebox.showwarning("Error", "Carga un archivo primero.")
            return

        

        self.file_section.pack_forget()
        self.window_section.pack_forget()
        self.options_section.pack_forget()

        self.progress_section.pack(fill="x")

        self.geometry("550x420")

        threading.Thread(target=self.worker, daemon=True).start()

    def worker(self):

        try:

            target_title = self.window_list.get()
            parent_hwnd = win32gui.FindWindow(None, target_title)

            self.instructions_panel.pack(pady=20)

            for s in range(10, 0, -1):

                msg = (
                    f"INICIANDO EN {s}s\n\n"
                    "• Coloque el cursor en el campo\n"
                    "• Haz clic izquierdo\n"
                    "• No toques el teclado"
                )

                self.instructions_label.configure(text=msg)
                self.update()
                time.sleep(1)

            self.instructions_panel.pack_forget()

            self.start_time = time.time() # Empezar timer

            with open(self.file_path, 'r', encoding='utf-8') as f:
                lines = [l.strip() for l in f if l.strip()]

            i = 0

            while i < len(lines):

                line = lines[i]

                self.status.configure(
                    text=f"Escribiendo {i+1} de {len(lines)}"
                )

                self.progress.set((i+1)/len(lines))

                completed = self.engine.send_input(
                    parent_hwnd,
                    line,
                    self.check_clear.get(),
                    False
                )

                if completed:
                    self.total_processed = i + 1
                    i += 1
                    time.sleep(self.delay_slider.get())

            self.finalize_ui("exito")

        except:
            self.finalize_ui("error")

    def finalize_ui(self, result):

        runtime = int(time.time() - self.start_time)

        mins = runtime // 60
        secs = runtime % 60

        self.runtime_label.configure(
            text=f"Tiempo total: {mins}m {secs}s"
        )

        self.btn_back.pack(pady=10)

        self.status.configure(
            text=f"Items Procesados {self.total_processed}/{self.current_total_lines}",
            text_color="#27ae60"
        )

    def reset_to_main(self):

        self.reset_state()

        self.progress_section.pack_forget()

        # Hide everything first
        self.file_section.pack_forget()
        self.window_section.pack_forget()
        self.options_section.pack_forget()
        self.control_section.pack_forget()

        # Re-pack in correct order
        self.file_section.pack(fill="x")
        self.window_section.pack(fill="x")
        self.options_section.pack(fill="x")
        self.control_section.pack(fill="x")

        self.geometry("550x720")

        self.progress.set(0)
        self.runtime_label.configure(text="")
        self.file_info.configure(text="Ningún archivo seleccionado", text_color="gray")

    def handle_stop(self):
        self.running = False

    def refresh_windows(self):

        windows = self.engine.get_windows()

        self.window_list.configure(values=windows)

        if windows:
            self.window_list.set(windows[0])