import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.modules.actualizacion.config import load_config
from app.modules.actualizacion.service import (
    UpdateAnalysisError,
    analyze_update,
    current_input_signature,
    generate_updated_workbook,
)


class ActualizacionFrame(ctk.CTkScrollableFrame):
    def __init__(self, master):
        super().__init__(master)

        self.config_data = load_config()
        self.analysis = None

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(8, weight=1)

        self._build_ui()

    def _build_ui(self):
        title = ctk.CTkLabel(
            self,
            text="Actualización de datos",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))

        guide = ctk.CTkTextbox(self, height=165, font=ctk.CTkFont(size=15))
        guide.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        guide.insert(
            "1.0",
            "Actualización semestral del archivo maestro de estudiantes:\n\n"
            "1) Selecciona el archivo estudiantes_actualizado del semestre anterior.\n"
            "2) Selecciona la carpeta que contiene los Excel finales de notas.\n"
            "3) Indica el semestre al que corresponden las notas (ejemplo: 2-2025).\n"
            "4) Presiona 'Analizar actualización' y revisa el resumen y el log.\n"
            "5) Si el análisis es correcto, genera el nuevo archivo. El original nunca se sobrescribe.\n\n"
            "La herramienta utiliza siempre la columna Promedio, cruza estudiantes por RUT y agrega "
            "estudiantes nuevos únicamente cuando registran FPpCD o TIC I en el semestre procesado."
        )
        guide.configure(state="disabled")

        ctk.CTkLabel(self, text="Archivo estudiantes:").grid(
            row=2, column=0, padx=10, pady=5, sticky="w"
        )
        self.master_entry = ctk.CTkEntry(self)
        self.master_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self, text="Buscar", command=self.select_master_file).grid(
            row=2, column=2, padx=10, pady=5
        )

        ctk.CTkLabel(self, text="Carpeta de notas:").grid(
            row=3, column=0, padx=10, pady=5, sticky="w"
        )
        self.notes_entry = ctk.CTkEntry(self)
        self.notes_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self, text="Buscar", command=self.select_notes_folder).grid(
            row=3, column=2, padx=10, pady=5
        )

        ctk.CTkLabel(self, text="Semestre de notas:").grid(
            row=4, column=0, padx=10, pady=5, sticky="w"
        )
        self.semester_entry = ctk.CTkEntry(self, placeholder_text="Ej: 2-2025")
        self.semester_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        output_note = ctk.CTkLabel(
            self,
            text=(
                "El nuevo archivo se guardará en la misma carpeta del archivo de estudiantes, "
                "con nombre estudiantes_actualizado_SEMESTRE_AÑO.xlsx."
            ),
            font=ctk.CTkFont(size=13),
            wraplength=850,
            justify="left",
        )
        output_note.grid(row=5, column=0, columnspan=3, sticky="w", padx=10, pady=(3, 8))

        buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        buttons_frame.grid(row=6, column=0, columnspan=3, sticky="ew", padx=10, pady=8)

        ctk.CTkButton(
            buttons_frame,
            text="Analizar actualización",
            command=self.run_analysis,
        ).pack(side="left", padx=(0, 10))

        self.generate_button = ctk.CTkButton(
            buttons_frame,
            text="Generar Excel actualizado",
            command=self.run_generation,
            state="disabled",
        )
        self.generate_button.pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            buttons_frame,
            text="Limpiar log",
            command=self.clear_log,
        ).pack(side="left")

        self.summary_frame = ctk.CTkFrame(self)
        self.summary_frame.grid(row=7, column=0, columnspan=3, sticky="ew", padx=10, pady=(5, 10))
        self.summary_frame.grid_columnconfigure(0, weight=1)

        self.summary_label = ctk.CTkLabel(
            self.summary_frame,
            text="Aún no se ha realizado un análisis.",
            font=ctk.CTkFont(size=15),
            justify="left",
            anchor="w",
        )
        self.summary_label.grid(row=0, column=0, sticky="ew", padx=12, pady=12)

        self.log_box = ctk.CTkTextbox(self, height=330)
        self.log_box.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def invalidate_analysis(self):
        self.analysis = None
        self.generate_button.configure(state="disabled")
        self.summary_label.configure(text="Aún no se ha realizado un análisis.")

    def select_master_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo maestro de estudiantes",
            filetypes=[("Excel", "*.xlsx *.xlsm")],
        )
        if path:
            self.master_entry.delete(0, "end")
            self.master_entry.insert(0, path)
            self.invalidate_analysis()

    def select_notes_folder(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta con Excel finales de notas")
        if path:
            self.notes_entry.delete(0, "end")
            self.notes_entry.insert(0, path)
            self.invalidate_analysis()

    def _get_inputs(self):
        master_path = self.master_entry.get().strip()
        notes_folder = self.notes_entry.get().strip()
        semester = self.semester_entry.get().strip()

        if not master_path:
            raise ValueError("Debes seleccionar el archivo maestro de estudiantes.")
        if not notes_folder:
            raise ValueError("Debes seleccionar la carpeta con las notas.")
        if not semester:
            raise ValueError("Debes indicar el semestre de las notas.")

        return master_path, notes_folder, semester

    def _summary_text(self, stats: dict) -> str:
        duplicates = stats["source_duplicates"] + stats["existing_duplicates"]
        return (
            "Resumen del análisis\n"
            f"Archivos encontrados: {stats['files_found']}  |  "
            f"Procesados: {stats['files_processed']}  |  "
            f"Registros leídos: {stats['raw_records']}\n"
            f"Notas a actualizar: {stats['grade_updates']}  |  "
            f"Estudiantes nuevos: {stats['new_students']}  |  "
            f"Nuevos bloqueados: {stats['new_students_blocked']}\n"
            f"Duplicados ignorados: {duplicates}  |  "
            f"Advertencias: {stats['warnings']}  |  "
            f"Errores: {stats['errors']}"
        )

    def run_analysis(self):
        try:
            master_path, notes_folder, semester = self._get_inputs()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return

        self.clear_log()
        self.invalidate_analysis()
        self.log("Iniciando análisis de actualización...")

        try:
            analysis = analyze_update(
                master_path=master_path,
                notes_folder=notes_folder,
                semester=semester,
                config=self.config_data,
                logger=self.log,
            )
            self.analysis = analysis
            summary = self._summary_text(analysis["stats"])
            self.summary_label.configure(text=summary)

            if analysis["has_changes"]:
                self.generate_button.configure(state="normal")
                messagebox.showinfo(
                    "Análisis finalizado",
                    summary + "\n\nRevisa el log antes de generar el nuevo archivo.",
                )
            else:
                self.generate_button.configure(state="disabled")
                messagebox.showinfo(
                    "Análisis finalizado",
                    summary + "\n\nNo se detectaron cambios válidos para generar un nuevo archivo.",
                )

        except Exception as exc:
            self.log(f"[ERROR GENERAL] {exc}")
            messagebox.showerror("Error", str(exc))

    def run_generation(self):
        if self.analysis is None:
            messagebox.showerror("Error", "Debes analizar la actualización antes de generar el Excel.")
            return

        try:
            master_path, notes_folder, semester = self._get_inputs()

            if current_input_signature(master_path, notes_folder, semester) != self.analysis.get(
                "input_signature"
            ):
                raise UpdateAnalysisError(
                    "Las entradas cambiaron desde el último análisis. "
                    "Vuelve a ejecutar 'Analizar actualización'."
                )

            self.log("\nGenerando archivo actualizado...")
            result = generate_updated_workbook(self.analysis, logger=self.log)

            validation = result["validation"]
            summary = (
                "Archivo generado correctamente.\n\n"
                f"Notas actualizadas: {result['grade_updates']}\n"
                f"Estudiantes nuevos: {result['new_students']}\n"
                f"RUT duplicados detectados en validación final: {validation['duplicate_ruts']}\n"
                f"RUT vacíos/invalidos históricos: {validation['invalid_ruts']}\n\n"
                f"Salida:\n{result['output_path']}"
            )

            self.generate_button.configure(state="disabled")
            messagebox.showinfo("Proceso finalizado", summary)

        except Exception as exc:
            self.log(f"[ERROR GENERAL] {exc}")
            messagebox.showerror("Error", str(exc))
