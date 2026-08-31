import customtkinter as ctk


class NavigationPanel(ctk.CTkFrame):

    def __init__(self, master, on_select, modules, groups):
        super().__init__(master, corner_radius=0)

        self.on_select = on_select
        self.modules = modules
        self.groups = groups

        self.buttons = {}
        self.group_buttons = {}
        self.group_frames = {}

        self.active_group = None

        self.grid_rowconfigure(99, weight=1)
        self.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            self,
            text="Minor CD",
            font=ctk.CTkFont(size=24, weight="bold"),
        )
        title.grid(
            row=0,
            column=0,
            padx=20,
            pady=(20, 20),
            sticky="w",
        )

        # ---------------------------------------------------------
        # Inicio
        # ---------------------------------------------------------
        home_module = next(
            module for module in self.modules
            if module["key"] == "home"
        )

        self._create_module_button(
            parent=self,
            module=home_module,
            row=1,
            padx=15,
        )

        # ---------------------------------------------------------
        # Grupos
        # ---------------------------------------------------------
        current_row = 2

        for group in self.groups:
            group_key = group["key"]

            group_button = ctk.CTkButton(
                self,
                text=f"▸ {group['label']}",
                anchor="w",
                height=42,
                fg_color="transparent",
                text_color=("black", "white"),
                hover_color=("gray85", "gray20"),
                font=ctk.CTkFont(weight="bold"),
                command=lambda key=group_key: self.toggle_group(key),
            )
            group_button.grid(
                row=current_row,
                column=0,
                padx=15,
                pady=(10, 2),
                sticky="ew",
            )

            self.group_buttons[group_key] = group_button
            current_row += 1

            group_frame = ctk.CTkFrame(
                self,
                fg_color="transparent",
                corner_radius=0,
            )
            group_frame.grid_columnconfigure(0, weight=1)
            group_frame.grid(
                row=current_row,
                column=0,
                sticky="ew",
            )

            self.group_frames[group_key] = group_frame

            group_modules = [
                module
                for module in self.modules
                if module.get("group") == group_key
            ]

            for module_row, module in enumerate(group_modules):
                self._create_module_button(
                    parent=group_frame,
                    module=module,
                    row=module_row,
                    padx=(30, 15),
                )

            # Todos los grupos parten compactados.
            group_frame.grid_remove()
            current_row += 1

    def _create_module_button(self, parent, module, row, padx):
        button = ctk.CTkButton(
            parent,
            text=module["label"],
            anchor="w",
            height=42,
            fg_color="transparent",
            text_color=("black", "white"),
            hover_color=("gray85", "gray20"),
            command=lambda module_key=module["key"]: self.select(module_key),
        )
        button.grid(
            row=row,
            column=0,
            padx=padx,
            pady=4,
            sticky="ew",
        )

        self.buttons[module["key"]] = button

    def toggle_group(self, group_key: str):
        if self.active_group == group_key:
            self._close_group(group_key)
            return

        if self.active_group is not None:
            self._close_group(self.active_group)

        self._open_group(group_key)

    def _open_group(self, group_key: str):
        if group_key not in self.group_frames:
            return

        if self.active_group is not None and self.active_group != group_key:
            self._close_group(self.active_group)

        self.group_frames[group_key].grid()

        group = next(
            group
            for group in self.groups
            if group["key"] == group_key
        )

        self.group_buttons[group_key].configure(
            text=f"▾ {group['label']}"
        )
        self.active_group = group_key

    def _close_group(self, group_key: str):
        if group_key not in self.group_frames:
            return

        self.group_frames[group_key].grid_remove()

        group = next(
            group
            for group in self.groups
            if group["key"] == group_key
        )

        self.group_buttons[group_key].configure(
            text=f"▸ {group['label']}"
        )

        if self.active_group == group_key:
            self.active_group = None

    def close_all_groups(self):
        """Compacta completamente el menú lateral."""
        for group_key in self.group_frames:
            self.group_frames[group_key].grid_remove()

            group = next(
                group
                for group in self.groups
                if group["key"] == group_key
            )

            self.group_buttons[group_key].configure(
                text=f"▸ {group['label']}"
            )

        self.active_group = None

    def select(self, key: str):
        self.highlight_selected(key)
        self.on_select(key)

    def highlight_selected(self, selected_key: str):
        selected_module = next(
            (
                module
                for module in self.modules
                if module["key"] == selected_key
            ),
            None,
        )

        # Inicio representa el nivel superior: al volver a él,
        # el menú lateral queda completamente compactado.
        if selected_key == "home":
            self.close_all_groups()

        # Si el módulo pertenece a un grupo cerrado,
        # abrir automáticamente ese grupo.
        elif selected_module is not None:
            group_key = selected_module.get("group")

            if group_key is not None and group_key != self.active_group:
                self._open_group(group_key)

        for key, button in self.buttons.items():
            if key == selected_key:
                button.configure(
                    fg_color=("gray75", "gray25"),
                    text_color=("black", "white"),
                )
            else:
                button.configure(
                    fg_color="transparent",
                    text_color=("black", "white"),
                )
