import customtkinter as ctk


class NavigationPanel(ctk.CTkFrame):
    def __init__(self, master, on_select):
        super().__init__(master, corner_radius=0)

        self.on_select = on_select
        self.buttons = {}

        self.grid_rowconfigure(99, weight=1)
        self.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            self,
            text="Minor CD",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.grid(row=0, column=0, padx=20, pady=(20, 20), sticky="w")

        modules = [
            ("Inicio", "home"),
            ("Postulación", "postulacion"),
            ("Aceptación", "aceptacion"),
            ("Actualización de datos", "actualizacion"),
            ("Filtrado", "filtrado"),
            ("Inscripción", "inscripcion"),
            ("Notas", "notas"),
            ("Finalizados", "finalizados"),
        ]

        for idx, (label, key) in enumerate(modules, start=1):
            button = ctk.CTkButton(
                self,
                text=label,
                anchor="w",
                height=42,
                command=lambda module_key=key: self.select(module_key)
            )
            button.grid(row=idx, column=0, padx=15, pady=6, sticky="ew")
            self.buttons[key] = button

    def select(self, key: str):
        self.highlight_selected(key)
        self.on_select(key)

    def highlight_selected(self, selected_key: str):
        for key, button in self.buttons.items():
            if key == selected_key:
                button.configure(fg_color=("gray75", "gray25"))
            else:
                button.configure(fg_color="transparent")