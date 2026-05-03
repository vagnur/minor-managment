import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.modules.inscripcion.config import load_config, save_config
from app.modules.inscripcion.service import (
    get_available_subjects,
    validate_excel_workbook,
    process_inscripcion,
)


class InscripcionFrame(ctk.CTkScrollableFrame):
    def __init__(self, master):
        super().__init__(master)

        self.config_data = load_config()
        self.subject_vars = {}
        self.subject_inputs = {}
        self.subject_frames = {}

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(12, weight=1)

        self._build_ui()

    def _build_ui(self):
        title = ctk.CTkLabel(
            self,
            text="Módulo de Inscripción",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))

        guide = ctk.CTkTextbox(self, height=145, font=ctk.CTkFont(size=15))
        guide.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        guide.insert(
            "1.0",
            "Preparación previa:\n\n"
            "1) Contar con el Excel final de inscripción, con una hoja por asignatura.\n"
            "2) Verificar que cada hoja tenga la estructura esperada y los horarios finales asignados.\n"
            "3) Seleccionar una o más asignaturas a procesar.\n"
            "4) Validar el archivo si se desea revisar hojas faltantes o vacías.\n"
            "5) Generar los formularios para firma de jefaturas de carrera.\n"
        )
        guide.configure(state="disabled")

        ctk.CTkLabel(self, text="Archivo Excel:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.excel_entry = ctk.CTkEntry(self)
        self.excel_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self, text="Buscar", command=self.select_excel).grid(row=2, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Carpeta salida:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.output_entry = ctk.CTkEntry(self)
        self.output_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.output_entry.insert(0, self.config_data.get("base_output_folder", "output/inscripcion"))
        ctk.CTkButton(self, text="Buscar", command=self.select_output_folder).grid(row=3, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Semestre:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.semestre_entry = ctk.CTkEntry(self)
        self.semestre_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="Fecha documento:").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.fecha_entry = ctk.CTkEntry(self)
        self.fecha_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        subjects_frame = ctk.CTkFrame(self)
        subjects_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=10, pady=10)
        subjects_frame.grid_columnconfigure(0, weight=1)
        subjects_frame.grid_columnconfigure(1, weight=1)
        subjects_frame.grid_columnconfigure(2, weight=1)

        subjects_label = ctk.CTkLabel(
            subjects_frame,
            text="Asignaturas a procesar",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        subjects_label.grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 5), sticky="w")

        available_subjects = get_available_subjects(self.config_data)

        columns_per_row = 3

        for idx, subject_name in enumerate(available_subjects):
            row = 1 + (idx // columns_per_row)
            col = idx % columns_per_row

            var = ctk.BooleanVar(value=True)
            checkbox = ctk.CTkCheckBox(
                subjects_frame,
                text=subject_name,
                variable=var
            )
            checkbox.grid(row=row, column=col, padx=10, pady=3, sticky="w")
            self.subject_vars[subject_name] = var

        buttons_row = 1 + ((len(available_subjects) - 1) // columns_per_row) + 1

        subjects_buttons = ctk.CTkFrame(subjects_frame, fg_color="transparent")
        ctk.CTkButton(
            subjects_buttons,
            text="Configurar asignaturas",
            command=self.build_subject_inputs
        ).pack(side="left", padx=(10, 0))
        subjects_buttons.grid(row=buttons_row, column=0, columnspan=3, padx=10, pady=(8, 10), sticky="w")

        ctk.CTkButton(subjects_buttons, text="Seleccionar todas", command=self.select_all_subjects).pack(side="left", padx=(0, 10))
        ctk.CTkButton(subjects_buttons, text="Limpiar selección", command=self.clear_subject_selection).pack(side="left")

        self.save_config_checkbox = ctk.CTkCheckBox(self, text="Guardar carpeta de salida")
        self.save_config_checkbox.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        self.save_config_checkbox.select()

        self.buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.buttons_frame.grid(row=8, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(self.buttons_frame, text="Validar archivo", command=self.run_validation).pack(side="left", padx=(0, 10))
        ctk.CTkButton(self.buttons_frame, text="Generar formularios", command=self.run_process).pack(side="left", padx=(0, 10))
        ctk.CTkButton(self.buttons_frame, text="Limpiar log", command=self.clear_log).pack(side="left")

        self.log_box = ctk.CTkTextbox(self)
        self.log_box.grid(row=12, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

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

    def get_selected_subjects(self) -> list[str]:
        return [
            subject_name
            for subject_name, var in self.subject_vars.items()
            if var.get()
        ]

    def select_all_subjects(self):
        for var in self.subject_vars.values():
            var.set(True)

    def clear_subject_selection(self):
        for var in self.subject_vars.values():
            var.set(False)

    def validate_common_inputs(self):
        excel_path = self.excel_entry.get().strip()
        output_folder = self.output_entry.get().strip()
        semestre = self.semestre_entry.get().strip()
        fecha_documento = self.fecha_entry.get().strip()
        selected_subjects = self.get_selected_subjects()

        if not excel_path:
            raise ValueError("Debes seleccionar un archivo Excel.")

        if not output_folder:
            raise ValueError("Debes seleccionar una carpeta de salida.")

        if not semestre:
            raise ValueError("Debes indicar el semestre.")

        if not fecha_documento:
            raise ValueError("Debes indicar la fecha del documento.")

        if not selected_subjects:
            raise ValueError("Debes seleccionar al menos una asignatura.")

        return excel_path, output_folder, semestre, fecha_documento, selected_subjects

    def run_validation(self):
        try:
            excel_path, _, _, _, selected_subjects = self.validate_common_inputs()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.clear_log()
        self.log("Iniciando validación de archivo...")

        try:
            result = validate_excel_workbook(
                excel_path=excel_path,
                config=self.config_data,
                selected_subjects=selected_subjects,
            )

            self.log(f"Hojas disponibles en Excel: {', '.join(result['available_sheets'])}")

            if result["missing_sheets"]:
                self.log("Hojas faltantes:")
                for sheet in result["missing_sheets"]:
                    self.log(f" - {sheet}")
            else:
                self.log("No hay hojas faltantes.")

            if result["empty_sheets"]:
                self.log("Hojas vacías:")
                for sheet in result["empty_sheets"]:
                    self.log(f" - {sheet}")
            else:
                self.log("No hay hojas vacías entre las seleccionadas.")

            messagebox.showinfo("Validación finalizada", "La validación del archivo terminó correctamente.")

        except Exception as e:
            self.log(f"Error de validación: {e}")
            messagebox.showerror("Error", str(e))

    def run_process(self):
        try:
            excel_path, output_folder, semestre, fecha_documento, selected_subjects = self.validate_common_inputs()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.config_data["base_output_folder"] = output_folder

        if self.save_config_checkbox.get() == 1:
            save_config(self.config_data)

        self.clear_log()
        self.log("Iniciando proceso de inscripción...")

        try:
            subject_runtime_configs = {}
            for subject_name in selected_subjects:
                inputs = self.subject_inputs.get(subject_name, {})
                subject_config = self.config_data["subjects"][subject_name]

                horarios_catedra = []
                horarios_lab = []

                if subject_config.get("has_catedra", False):
                    horarios_catedra = self.parse_schedule_list(inputs["horarios_catedra"].get())

                    if not horarios_catedra:
                        messagebox.showerror("Error", f"Debes ingresar horarios de cátedra para {subject_name}.")
                        return

                if subject_config.get("has_lab", False):
                    horarios_lab = self.parse_schedule_list(inputs["horarios_lab"].get())

                    if not horarios_lab:
                        messagebox.showerror("Error", f"Debes ingresar horarios de laboratorio para {subject_name}.")
                        return

                subject_runtime_configs[subject_name] = {
                    "horarios_catedra": horarios_catedra,
                    "horarios_lab": horarios_lab,
                }
            result = process_inscripcion(
                excel_path=excel_path,
                output_folder=output_folder,
                semestre=semestre,
                fecha_documento=fecha_documento,
                selected_subjects=selected_subjects,
                subject_runtime_configs=subject_runtime_configs,
                config=self.config_data,
                logger=self.log,
            )

            summary_lines = [
                "",
                "Proceso finalizado.",
                f"Asignaturas procesadas: {', '.join(result['subjects_processed']) if result['subjects_processed'] else 'Ninguna'}",
                f"Asignaturas omitidas por hoja vacía: {', '.join(result['subjects_skipped']) if result['subjects_skipped'] else 'Ninguna'}",
                f"Hojas faltantes: {', '.join(result['missing_sheets']) if result['missing_sheets'] else 'Ninguna'}",
                f"Formularios generados correctamente: {result['total_ok']}",
                f"Errores totales: {result['total_errors']}",
            ]

            summary = "\n".join(summary_lines)
            self.log(summary)
            messagebox.showinfo("Proceso finalizado", summary)

        except Exception as e:
            self.log(f"Error general: {e}")
            messagebox.showerror("Error", str(e))

    def parse_schedule_list(self, text: str) -> list:
        return [
            item.strip()
            for item in text.split(";")
            if item.strip()
        ]

    def build_subject_inputs(self):
        for frame in self.subject_frames.values():
            frame.destroy()

        self.subject_frames = {}
        self.subject_inputs = {}

        subjects_config = self.config_data["subjects"]
        start_row = 10

        for i, subject_name in enumerate(self.get_selected_subjects()):
            subject_config = subjects_config[subject_name]
            row = start_row + i

            frame = ctk.CTkFrame(self)
            frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=10, pady=8)
            frame.grid_columnconfigure(1, weight=1)

            display_name = subject_config.get("display_name", subject_name)

            ctk.CTkLabel(
                frame,
                text=display_name,
                font=ctk.CTkFont(size=16, weight="bold")
            ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 5))

            entries = {}
            current_row = 1

            if subject_config.get("has_catedra", False):
                ctk.CTkLabel(frame, text="Horarios cátedra:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
                catedra_entry = ctk.CTkEntry(frame, placeholder_text="Ej: L7 W7; M7 J7")
                catedra_entry.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
                entries["horarios_catedra"] = catedra_entry
                current_row += 1

            if subject_config.get("has_lab", False):
                ctk.CTkLabel(frame, text="Horarios laboratorio:").grid(row=current_row, column=0, padx=10, pady=5, sticky="w")
                lab_entry = ctk.CTkEntry(frame, placeholder_text="Ej: L7 W7; M3 W2; J7 V6")
                lab_entry.grid(row=current_row, column=1, padx=10, pady=5, sticky="ew")
                entries["horarios_lab"] = lab_entry

            self.subject_frames[subject_name] = frame
            self.subject_inputs[subject_name] = entries

        next_row = start_row + len(self.get_selected_subjects())

        self.buttons_frame.grid(row=next_row, column=0, columnspan=3, sticky="ew", padx=10, pady=10)
        self.log_box.grid(row=next_row + 1, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

        self.grid_rowconfigure(next_row + 1, weight=1)