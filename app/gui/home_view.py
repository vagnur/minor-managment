import customtkinter as ctk


MODULE_CARD_INFO = {
    "filtrado": {
        "icon": "⌕",
        "description": (
            "Accede al espacio reservado para el filtrado inicial de estudiantes."
        ),
    },
    "postulacion": {
        "icon": "+",
        "description": (
            "Procesa y consolida las postulaciones recibidas al Minor."
        ),
    },
    "aceptacion": {
        "icon": "✓",
        "description": (
            "Revisa postulaciones y genera los resultados del proceso de aceptación."
        ),
    },
    "inscripcion": {
        "icon": "→",
        "description": (
            "Prepara la información necesaria para la inscripción de estudiantes."
        ),
    },
    "preparacion_notas": {
        "icon": "▤",
        "description": (
            "Prepara y organiza la información necesaria antes del cierre de notas."
        ),
    },
    "notas": {
        "icon": "Σ",
        "description": (
            "Consolida las calificaciones finales y genera los archivos de notas."
        ),
    },
    "actualizacion": {
        "icon": "↻",
        "description": (
            "Incorpora las notas del semestre al archivo maestro de estudiantes."
        ),
    },
}


class ModuleCard(ctk.CTkFrame):
    """Tarjeta clickeable utilizada como acceso directo a un módulo."""

    def __init__(self, master, module, command):
        self._normal_color = ("gray96", "gray17")
        self._hover_color = ("gray91", "gray22")

        super().__init__(
            master,
            height=112,
            corner_radius=12,
            border_width=1,
            border_color=("gray80", "gray30"),
            fg_color=self._normal_color,
        )

        self.module = module
        self.command = command

        info = MODULE_CARD_INFO.get(
            module["key"],
            {
                "icon": "•",
                "description": "Accede a este módulo del programa.",
            },
        )

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_propagate(False)

        icon_box = ctk.CTkFrame(
            self,
            width=52,
            height=52,
            corner_radius=10,
            fg_color=("gray88", "gray25"),
        )
        icon_box.grid(
            row=0,
            column=0,
            rowspan=2,
            padx=(16, 12),
            pady=16,
            sticky="nw",
        )
        icon_box.grid_propagate(False)

        icon = ctk.CTkLabel(
            icon_box,
            text=info["icon"],
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        icon.place(relx=0.5, rely=0.5, anchor="center")

        title = ctk.CTkLabel(
            self,
            text=module["label"],
            font=ctk.CTkFont(size=17, weight="bold"),
            anchor="w",
        )
        title.grid(
            row=0,
            column=1,
            padx=(0, 16),
            pady=(16, 3),
            sticky="ew",
        )

        description = ctk.CTkLabel(
            self,
            text=info["description"],
            font=ctk.CTkFont(size=13),
            anchor="nw",
            justify="left",
            wraplength=390,
        )
        description.grid(
            row=1,
            column=1,
            padx=(0, 16),
            pady=(0, 16),
            sticky="nsew",
        )

        self._interactive_widgets = [
            self,
            icon_box,
            icon,
            title,
            description,
        ]

        for widget in self._interactive_widgets:
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)

    def _on_click(self, _event=None):
        if self.command is not None:
            self.command()

    def _on_enter(self, _event=None):
        self.configure(fg_color=self._hover_color)

    def _on_leave(self, _event=None):
        self.configure(fg_color=self._normal_color)


class HomeView(ctk.CTkFrame):
    def __init__(self, master, on_select=None, modules=None, groups=None):
        super().__init__(master)

        self.on_select = on_select
        self.modules = modules or []
        self.groups = groups or []

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        title = ctk.CTkLabel(
            self,
            text="Programa de Gestión del Minor en Ciencia de Datos",
            font=ctk.CTkFont(size=28, weight="bold"),
        )
        title.grid(
            row=0,
            column=0,
            padx=20,
            pady=(20, 6),
            sticky="w",
        )

        subtitle = ctk.CTkLabel(
            self,
            text="Selecciona el proceso que deseas realizar.",
            font=ctk.CTkFont(size=15),
        )
        subtitle.grid(
            row=1,
            column=0,
            padx=20,
            pady=(0, 14),
            sticky="w",
        )

        content = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            corner_radius=0,
        )
        content.grid(
            row=2,
            column=0,
            padx=20,
            pady=(0, 20),
            sticky="nsew",
        )
        content.grid_columnconfigure(0, weight=1)

        current_row = 0

        for group_index, group in enumerate(self.groups):
            group_modules = [
                module
                for module in self.modules
                if module.get("group") == group["key"]
            ]

            if not group_modules:
                continue

            section_title = ctk.CTkLabel(
                content,
                text=group["label"],
                font=ctk.CTkFont(size=19, weight="bold"),
                anchor="w",
            )
            section_title.grid(
                row=current_row,
                column=0,
                padx=2,
                pady=((4 if group_index == 0 else 18), 8),
                sticky="ew",
            )
            current_row += 1

            card_grid = ctk.CTkFrame(
                content,
                fg_color="transparent",
                corner_radius=0,
            )
            card_grid.grid(
                row=current_row,
                column=0,
                sticky="ew",
            )
            card_grid.grid_columnconfigure(0, weight=1, uniform="cards")
            card_grid.grid_columnconfigure(1, weight=1, uniform="cards")

            for index, module in enumerate(group_modules):
                row = index // 2
                column = index % 2

                card = ModuleCard(
                    card_grid,
                    module=module,
                    command=lambda key=module["key"]: self._select_module(key),
                )
                card.grid(
                    row=row,
                    column=column,
                    padx=(0, 6) if column == 0 else (6, 0),
                    pady=6,
                    sticky="nsew",
                )

            current_row += 1

    def _select_module(self, module_key: str):
        if self.on_select is not None:
            self.on_select(module_key)
