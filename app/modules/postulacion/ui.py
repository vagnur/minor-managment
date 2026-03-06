import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.modules.postulacion.config import load_config, save_config
from app.modules.postulacion.service import process_postulacion


class PostulacionFrame(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.config_data = load_config()

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(8, weight=1)

        self._build_ui()

    def _build_ui(self):
        title = ctk.CTkLabel(
            self,
            text="Módulo de Postulación",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))

        guide = ctk.CTkTextbox(self, height=160, font=ctk.CTkFont(size=15))
        guide.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        guide.insert(
            "1.0",
            "Preparación previa del archivo de entrada:\n\n"
            "1) Descargar las postulaciones desde el formulario Google Forms en formato Excel.\n"
            "2) Revisar cada postulante para verificar si cumple las condiciones para participar en el Minor.\n"
            "3) Filtrar únicamente a los postulantes aceptados.\n"
            "4) Definir manualmente las secciones según disponibilidad de horarios.\n"
            "5) Ajustar los horarios de cada estudiante dejando un único horario final de cátedra y laboratorio.\n"
            "6) Consolidar toda la información en un único archivo Excel final.\n\n"
            "Ese archivo Excel consolidado es el que debe seleccionarse como entrada en este módulo "
            "para generar automáticamente los formularios Word de postulación."
        )
        guide.configure(state="disabled")

        ctk.CTkLabel(self, text="Archivo Excel:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.excel_entry = ctk.CTkEntry(self)
        self.excel_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self, text="Buscar", command=self.select_excel).grid(row=2, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Carpeta salida:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.output_entry = ctk.CTkEntry(self)
        self.output_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.output_entry.insert(0, self.config_data.get("base_output_folder", "output/formularios TIC I"))
        ctk.CTkButton(self, text="Buscar", command=self.select_output_folder).grid(row=3, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Hoja Excel:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.sheet_entry = ctk.CTkEntry(self)
        self.sheet_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")
        self.sheet_entry.insert(0, self.config_data.get("sheet_name", "Hoja 1"))

        ctk.CTkLabel(self, text="Plantilla fija:").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.template_entry = ctk.CTkEntry(self)
        self.template_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")
        self.template_entry.insert(0, self.config_data.get("template_path", "templates/proto.docx"))

        self.save_config_checkbox = ctk.CTkCheckBox(self, text="Guardar hoja/carpeta en configuración")
        self.save_config_checkbox.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        self.save_config_checkbox.select()

        buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        buttons_frame.grid(row=7, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(buttons_frame, text="Generar formularios", command=self.run_process).pack(side="left", padx=(0, 10))
        ctk.CTkButton(buttons_frame, text="Limpiar log", command=self.clear_log).pack(side="left")

        self.log_box = ctk.CTkTextbox(self)
        self.log_box.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

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
        sheet_name = self.sheet_entry.get().strip()
        template_path = self.template_entry.get().strip()

        if not excel_path:
            messagebox.showerror("Error", "Debes seleccionar un archivo Excel.")
            return

        if not output_folder:
            messagebox.showerror("Error", "Debes seleccionar una carpeta de salida.")
            return

        self.config_data["sheet_name"] = sheet_name
        self.config_data["base_output_folder"] = output_folder
        self.config_data["template_path"] = template_path

        if self.save_config_checkbox.get() == 1:
            save_config(self.config_data)

        self.clear_log()
        self.log("Iniciando proceso de postulación...")

        try:
            result = process_postulacion(
                excel_path=excel_path,
                output_folder=output_folder,
                config=self.config_data,
                logger=self.log
            )

            summary = (
                f"\nProceso finalizado.\n"
                f"Total registros: {result['total']}\n"
                f"Documentos generados: {result['ok']}\n"
                f"Errores: {result['errors']}\n"
            )
            self.log(summary)
            messagebox.showinfo("Proceso finalizado", summary)

        except Exception as e:
            self.log(f"Error general: {e}")
            messagebox.showerror("Error", str(e))