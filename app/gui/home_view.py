import customtkinter as ctk


class HomeView(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(master)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        title = ctk.CTkLabel(
            self,
            text="Programa de Gestión del Minor en Ciencia de Datos",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        subtitle = ctk.CTkLabel(
            self,
            text=(
                "Esta herramienta centraliza los procesos operativos del Minor.\n"
                "Selecciona un módulo en el panel izquierdo para comenzar."
            ),
            font=ctk.CTkFont(size=16)
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="w")

        info_box = ctk.CTkTextbox(self, height=240, font=ctk.CTkFont(size=15))
        info_box.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        info_box.insert(
            "1.0",
            "Módulos considerados en el programa:\n\n"
            "• Postulación\n"
            "• Aceptación\n"
            "• Actualización de datos\n"
            "• Filtrado\n"
            "• Inscripción\n"
            "• Notas\n"
            "• Finalizados\n\n"
            "Estado actual:\n"
            "• Postulación: operativo\n"
            "• Resto de módulos: en construcción\n"
        )
        info_box.configure(state="disabled")