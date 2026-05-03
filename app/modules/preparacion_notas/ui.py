import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.modules.preparacion_notas.config import load_config, save_config
from app.modules.preparacion_notas.service import (
    preview_source_folder,
    generate_prepared_grade_excels,
)


class PreparacionNotasFrame(ctk.CTkScrollableFrame):
    def __init__(self, master):
        super().__init__(master)

        self.config_data = load_config()
        self.rut_entries = {}
        self.requirements = []

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(12, weight=1)

        self._build_ui()

    def _build_ui(self):
        title = ctk.CTkLabel(
            self,
            text="Preparación de Notas",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))

        guide = ctk.CTkTextbox(self, height=130, font=ctk.CTkFont(size=15))
        guide.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        guide.insert(
            "1.0",
            "Preparación previa:\n\n"
            "1) Seleccionar carpeta con los Excel originales de notas.\n"
            "2) Seleccionar carpeta de salida.\n"
            "3) Cargar secciones detectadas.\n"
            "4) Ingresar el RUT del profesor correspondiente a cada sección.\n"
            "5) Generar los Excel preparados para el módulo de Notas.\n"
        )
        guide.configure(state="disabled")

        ctk.CTkLabel(self, text="Carpeta originales:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.source_entry = ctk.CTkEntry(self)
        self.source_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self, text="Buscar", command=self.select_source_folder).grid(row=2, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Carpeta salida:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.output_entry = ctk.CTkEntry(self)
        self.output_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.output_entry.insert(0, self.config_data.get("base_output_folder", "output/preparacion_notas"))
        ctk.CTkButton(self, text="Buscar", command=self.select_output_folder).grid(row=3, column=2, padx=10, pady=5)

        self.save_config_checkbox = ctk.CTkCheckBox(self, text="Guardar carpeta de salida")
        self.save_config_checkbox.grid(row=4, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        self.save_config_checkbox.select()

        buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        buttons_frame.grid(row=5, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(
            buttons_frame,
            text="Cargar secciones",
            command=self.load_sections
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            buttons_frame,
            text="Generar Excel preparados",
            command=self.run_process
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            buttons_frame,
            text="Limpiar log",
            command=self.clear_log
        ).pack(side="left")

        self.requirements_frame = ctk.CTkFrame(self)
        self.requirements_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=10, pady=10)
        self.requirements_frame.grid_columnconfigure(1, weight=1)

        req_title = ctk.CTkLabel(
            self.requirements_frame,
            text="RUT de profesores por sección",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        req_title.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 5))

        self.empty_requirements_label = ctk.CTkLabel(
            self.requirements_frame,
            text="Carga una carpeta para detectar las secciones disponibles.",
            font=ctk.CTkFont(size=14)
        )
        self.empty_requirements_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 10))

        self.log_box = ctk.CTkTextbox(self)
        self.log_box.grid(row=12, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def select_source_folder(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta con Excel originales")
        if path:
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, path)

    def select_output_folder(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if path:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)

    def validate_source_input(self):
        source_folder = self.source_entry.get().strip()

        if not source_folder:
            raise ValueError("Debes seleccionar la carpeta con los Excel originales.")

        return source_folder

    def validate_process_inputs(self):
        source_folder = self.source_entry.get().strip()
        output_folder = self.output_entry.get().strip()

        if not source_folder:
            raise ValueError("Debes seleccionar la carpeta con los Excel originales.")

        if not output_folder:
            raise ValueError("Debes seleccionar una carpeta de salida.")

        if not self.requirements:
            raise ValueError("Debes cargar las secciones antes de generar los Excel.")

        professor_ruts = {}

        for requirement in self.requirements:
            key = requirement["key"]
            entry = self.rut_entries.get(key)

            if entry is None:
                raise ValueError("Falta un campo de RUT en la interfaz.")

            rut = entry.get().strip()

            if not rut:
                raise ValueError(
                    f"Debes ingresar RUT para {requirement['subject']} "
                    f"sección {requirement['section_name']} - "
                    f"{requirement['professor_name']}."
                )

            professor_ruts[key] = rut

        return source_folder, output_folder, professor_ruts

    def save_output_config_if_needed(self, output_folder: str):
        self.config_data["base_output_folder"] = output_folder

        if self.save_config_checkbox.get() == 1:
            save_config(self.config_data)

    def clear_requirements_frame(self):
        for widget in self.requirements_frame.winfo_children():
            widget.destroy()

        self.rut_entries = {}

    def build_requirements_ui(self, requirements: list[dict]):
        self.clear_requirements_frame()

        title = ctk.CTkLabel(
            self.requirements_frame,
            text="RUT de profesores por sección",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 5))

        if not requirements:
            empty_label = ctk.CTkLabel(
                self.requirements_frame,
                text="No se detectaron secciones.",
                font=ctk.CTkFont(size=14)
            )
            empty_label.grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 10))
            return

        for idx, requirement in enumerate(requirements, start=1):
            label_text = (
                f"{requirement['subject']} | Sección {requirement['section_name']} | "
                f"Profesor {requirement['rut_type']}: {requirement['professor_name']}"
            )

            ctk.CTkLabel(
                self.requirements_frame,
                text=label_text
            ).grid(row=idx, column=0, padx=10, pady=5, sticky="w")

            entry = ctk.CTkEntry(self.requirements_frame, placeholder_text="Ej: 12.345.678-9")
            entry.grid(row=idx, column=1, padx=10, pady=5, sticky="ew")

            self.rut_entries[requirement["key"]] = entry

    def load_sections(self):
        try:
            source_folder = self.validate_source_input()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.clear_log()
        self.log("Cargando secciones desde archivos originales...")

        try:
            result = preview_source_folder(
                folder_path=source_folder,
                config=self.config_data,
                logger=self.log,
            )

            self.requirements = result["requirements"]
            self.build_requirements_ui(self.requirements)

            summary = (
                "\nCarga finalizada.\n"
                f"Secciones detectadas: {result['total_sections']}\n"
                f"Estudiantes detectados: {result['total_students']}\n"
                f"RUT requeridos: {result['total_ruts_required']}"
            )

            self.log(summary)
            messagebox.showinfo("Carga finalizada", summary)

        except Exception as e:
            self.log(f"Error cargando secciones: {e}")
            messagebox.showerror("Error", str(e))

    def run_process(self):
        try:
            source_folder, output_folder, professor_ruts = self.validate_process_inputs()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.save_output_config_if_needed(output_folder)

        self.clear_log()
        self.log("Generando Excel preparados de notas...")

        try:
            result = generate_prepared_grade_excels(
                folder_path=source_folder,
                output_folder=output_folder,
                professor_ruts=professor_ruts,
                config=self.config_data,
                logger=self.log,
            )

            summary = (
                "\nProceso finalizado.\n"
                f"Excel generados correctamente: {result['total_ok']}\n"
                f"Errores totales: {result['total_errors']}"
            )

            self.log(summary)
            messagebox.showinfo("Proceso finalizado", summary)

        except Exception as e:
            self.log(f"Error general: {e}")
            messagebox.showerror("Error", str(e))