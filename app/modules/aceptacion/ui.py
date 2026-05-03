import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.modules.aceptacion.config import load_config, save_config
from app.modules.aceptacion.service import process_aceptacion


class AceptacionFrame(ctk.CTkScrollableFrame):
    def __init__(self, master):
        super().__init__(master)

        self.config_data = load_config()

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(10, weight=1)

        self._build_ui()

    def _build_ui(self):
        title = ctk.CTkLabel(
            self,
            text="Módulo de Aceptación",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))

        guide = ctk.CTkTextbox(self, height=150, font=ctk.CTkFont(size=15))
        guide.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        guide.insert(
            "1.0",
            "Preparación previa:\n\n"
            "1) Contar con el Excel final de estudiantes aceptados.\n"
            "2) Verificar que cada fila corresponda a un estudiante aceptado.\n"
            "3) Confirmar que el archivo contenga RUT, nombre completo, carrera y facultad.\n"
            "4) Ejecutar este módulo para generar el documento consolidado de ingreso al Minor.\n"
        )
        guide.configure(state="disabled")

        ctk.CTkLabel(self, text="Archivo Excel:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.excel_entry = ctk.CTkEntry(self)
        self.excel_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self, text="Buscar", command=self.select_excel).grid(row=2, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Carpeta salida:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.output_entry = ctk.CTkEntry(self)
        self.output_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.output_entry.insert(0, self.config_data.get("base_output_folder", "output/aceptacion"))
        ctk.CTkButton(self, text="Buscar", command=self.select_output_folder).grid(row=3, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Semestre ingreso:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.semestre_entry = ctk.CTkEntry(self)
        self.semestre_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="Año:").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.anio_entry = ctk.CTkEntry(self)
        self.anio_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="Iniciales director:").grid(row=6, column=0, padx=10, pady=5, sticky="w")
        self.iniciales_director_entry = ctk.CTkEntry(self)
        self.iniciales_director_entry.grid(row=6, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="Iniciales coordinador:").grid(row=7, column=0, padx=10, pady=5, sticky="w")
        self.iniciales_coordinador_entry = ctk.CTkEntry(self)
        self.iniciales_coordinador_entry.grid(row=7, column=1, padx=10, pady=5, sticky="ew")

        self.save_config_checkbox = ctk.CTkCheckBox(self, text="Guardar carpeta de salida")
        self.save_config_checkbox.grid(row=8, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        self.save_config_checkbox.select()

        buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        buttons_frame.grid(row=9, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(buttons_frame, text="Generar documento", command=self.run_process).pack(side="left", padx=(0, 10))
        ctk.CTkButton(buttons_frame, text="Limpiar log", command=self.clear_log).pack(side="left")

        self.log_box = ctk.CTkTextbox(self)
        self.log_box.grid(row=10, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

        self.grid_rowconfigure(10, weight=1)

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def select_excel(self):
        path = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.excel_entry.delete(0, "end")
            self.excel_entry.insert(0, path)

    def select_output_folder(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if path:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)

    def run_process(self):
        excel_path = self.excel_entry.get().strip()
        output_folder = self.output_entry.get().strip()
        semestre = self.semestre_entry.get().strip()
        anio = self.anio_entry.get().strip()
        iniciales_director = self.iniciales_director_entry.get().strip()
        iniciales_coordinador = self.iniciales_coordinador_entry.get().strip()

        if not excel_path:
            messagebox.showerror("Error", "Debes seleccionar un archivo Excel.")
            return

        if not output_folder:
            messagebox.showerror("Error", "Debes seleccionar una carpeta de salida.")
            return

        if not semestre or not anio:
            messagebox.showerror("Error", "Debes indicar semestre y año.")
            return

        if not iniciales_director or not iniciales_coordinador:
            messagebox.showerror("Error", "Debes indicar las iniciales requeridas.")
            return

        self.config_data["base_output_folder"] = output_folder

        if self.save_config_checkbox.get() == 1:
            save_config(self.config_data)

        self.clear_log()
        self.log("Iniciando proceso de aceptación...")

        try:
            result = process_aceptacion(
                excel_path=excel_path,
                output_folder=output_folder,
                semestre=semestre,
                anio=anio,
                iniciales_director=iniciales_director,
                iniciales_coordinador=iniciales_coordinador,
                config=self.config_data,
                logger=self.log
            )

            summary = (
                f"\nProceso finalizado.\n"
                f"Total registros: {result['total']}\n"
                f"Advertencias: {len(result['warnings'])}\n"
                f"Archivo generado: {result['output_path']}\n"
            )
            self.log(summary)
            messagebox.showinfo("Proceso finalizado", summary)

        except Exception as e:
            self.log(f"Error general: {e}")
            messagebox.showerror("Error", str(e))