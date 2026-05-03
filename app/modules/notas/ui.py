import customtkinter as ctk
from tkinter import filedialog, messagebox

from app.modules.notas.config import load_config, save_config
from app.modules.notas.service import (
    preview_grades_folder,
    generate_faculty_excels,
    generate_memos,
)


class NotasFrame(ctk.CTkScrollableFrame):
    def __init__(self, master):
        super().__init__(master)

        self.config_data = load_config()

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(10, weight=1)

        self._build_ui()

    def _build_ui(self):
        title = ctk.CTkLabel(
            self,
            text="Módulo de Notas",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 5))

        guide = ctk.CTkTextbox(self, height=135, font=ctk.CTkFont(size=15))
        guide.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 10))
        guide.insert(
            "1.0",
            "Preparación previa:\n\n"
            "1) Seleccionar la carpeta que contiene los Excel de notas por sección/asignatura.\n"
            "2) Seleccionar la carpeta de salida.\n"
            "3) Completar semestre, fecha del documento y nombre de la vicedecana.\n"
            "4) Validar carpeta para revisar archivos detectados.\n"
            "5) Generar Excel por facultad o memorándums internos de notas.\n"
        )
        guide.configure(state="disabled")

        ctk.CTkLabel(self, text="Carpeta notas:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.folder_entry = ctk.CTkEntry(self)
        self.folder_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        ctk.CTkButton(self, text="Buscar", command=self.select_grades_folder).grid(row=2, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Carpeta salida:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.output_entry = ctk.CTkEntry(self)
        self.output_entry.grid(row=3, column=1, padx=10, pady=5, sticky="ew")
        self.output_entry.insert(0, self.config_data.get("base_output_folder", "output/notas"))
        ctk.CTkButton(self, text="Buscar", command=self.select_output_folder).grid(row=3, column=2, padx=10, pady=5)

        ctk.CTkLabel(self, text="Semestre:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.semestre_entry = ctk.CTkEntry(self, placeholder_text="Ej: 2025-2")
        self.semestre_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="Fecha documento:").grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.fecha_entry = ctk.CTkEntry(self, placeholder_text="Ej: 16 de marzo de 2026")
        self.fecha_entry.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="Vicedecana:").grid(row=6, column=0, padx=10, pady=5, sticky="w")
        self.vice_entry = ctk.CTkEntry(self, placeholder_text="Ej: Dra. Karina Acosta Barbosa")
        self.vice_entry.grid(row=6, column=1, padx=10, pady=5, sticky="ew")

        self.save_config_checkbox = ctk.CTkCheckBox(self, text="Guardar carpeta de salida")
        self.save_config_checkbox.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        self.save_config_checkbox.select()

        buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        buttons_frame.grid(row=8, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        ctk.CTkButton(
            buttons_frame,
            text="Validar carpeta",
            command=self.run_validation
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            buttons_frame,
            text="Generar Excel por facultad",
            command=self.run_faculty_excels
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            buttons_frame,
            text="Generar memorándums",
            command=self.run_memos
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            buttons_frame,
            text="Limpiar log",
            command=self.clear_log
        ).pack(side="left")

        self.log_box = ctk.CTkTextbox(self)
        self.log_box.grid(row=10, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def select_grades_folder(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta con Excel de notas")
        if path:
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, path)

    def select_output_folder(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if path:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, path)

    def validate_folder_inputs(self):
        folder_path = self.folder_entry.get().strip()

        if not folder_path:
            raise ValueError("Debes seleccionar la carpeta con los Excel de notas.")

        return folder_path

    def validate_output_inputs(self):
        folder_path = self.folder_entry.get().strip()
        output_folder = self.output_entry.get().strip()

        if not folder_path:
            raise ValueError("Debes seleccionar la carpeta con los Excel de notas.")

        if not output_folder:
            raise ValueError("Debes seleccionar una carpeta de salida.")

        return folder_path, output_folder

    def validate_memo_inputs(self):
        folder_path, output_folder = self.validate_output_inputs()

        semestre = self.semestre_entry.get().strip()
        fecha = self.fecha_entry.get().strip()
        vice = self.vice_entry.get().strip()

        if not semestre:
            raise ValueError("Debes indicar el semestre.")

        if not fecha:
            raise ValueError("Debes indicar la fecha del documento.")

        if not vice:
            raise ValueError("Debes indicar el nombre de la vicedecana.")

        return folder_path, output_folder, semestre, fecha, vice

    def save_output_config_if_needed(self, output_folder: str):
        self.config_data["base_output_folder"] = output_folder

        if self.save_config_checkbox.get() == 1:
            save_config(self.config_data)

    def run_validation(self):
        try:
            folder_path = self.validate_folder_inputs()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.clear_log()
        self.log("Iniciando validación de carpeta de notas...")

        try:
            result = preview_grades_folder(
                folder_path=folder_path,
                config=self.config_data,
                logger=self.log,
            )

            summary = (
                "\nValidación finalizada.\n"
                f"Archivos leídos: {result['files_read']}\n"
                f"Asignaturas detectadas: {', '.join(result['subjects_found'])}\n"
                f"Registros totales: {result['total_students']}"
            )

            self.log(summary)
            messagebox.showinfo("Validación finalizada", summary)

        except Exception as e:
            self.log(f"Error de validación: {e}")
            messagebox.showerror("Error", str(e))

    def run_faculty_excels(self):
        try:
            folder_path, output_folder = self.validate_output_inputs()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.save_output_config_if_needed(output_folder)

        self.clear_log()
        self.log("Iniciando generación de Excel por facultad...")

        try:
            result = generate_faculty_excels(
                folder_path=folder_path,
                output_folder=output_folder,
                config=self.config_data,
                logger=self.log,
            )

            summary = (
                "\nProceso finalizado.\n"
                f"Archivos por facultad generados: {result['total_faculties']}\n"
                f"Facultades: {', '.join(result['faculties']) if result['faculties'] else 'Ninguna'}"
            )

            self.log(summary)
            messagebox.showinfo("Proceso finalizado", summary)

        except Exception as e:
            self.log(f"Error general: {e}")
            messagebox.showerror("Error", str(e))

    def run_memos(self):
        try:
            folder_path, output_folder, semestre, fecha, vice = self.validate_memo_inputs()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.save_output_config_if_needed(output_folder)

        self.clear_log()
        self.log("Iniciando generación de memorándums...")

        try:
            result = generate_memos(
                folder_path=folder_path,
                output_folder=output_folder,
                fecha=fecha,
                vice=vice,
                semestre=semestre,
                config=self.config_data,
                logger=self.log,
            )

            summary = (
                "\nProceso finalizado.\n"
                f"Memorándums generados correctamente: {result['total_ok']}\n"
                f"Errores totales: {result['total_errors']}"
            )

            self.log(summary)
            messagebox.showinfo("Proceso finalizado", summary)

        except Exception as e:
            self.log(f"Error general: {e}")
            messagebox.showerror("Error", str(e))