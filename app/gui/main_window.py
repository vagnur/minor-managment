import customtkinter as ctk

from app.gui.navigation import NavigationPanel
from app.core.module_registry import MODULES

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")
ctk.set_widget_scaling(1.3)
ctk.set_window_scaling(1.3)


class PlaceholderView(ctk.CTkFrame):
    def __init__(self, master, module_name: str):
        super().__init__(master)

        self.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            self,
            text=module_name,
            font=ctk.CTkFont(size=26, weight="bold")
        )
        title.grid(row=0, column=0, padx=20, pady=(30, 10), sticky="w")

        message = ctk.CTkLabel(
            self,
            text="Módulo en construcción.",
            font=ctk.CTkFont(size=18)
        )
        message.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")

        detail = ctk.CTkLabel(
            self,
            text="Este espacio ya está reservado dentro del programa para integrar esta funcionalidad.",
            font=ctk.CTkFont(size=15),
            wraplength=700,
            justify="left"
        )
        detail.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="w")


class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Minor en Ciencia de Datos - Herramientas de Gestión")
        self.geometry("1280x780")
        self.minsize(1100, 720)

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.navigation = NavigationPanel(self, self.show_view, MODULES)
        self.navigation.grid(row=0, column=0, sticky="nsew")

        self.content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.views = self._build_views()

        self.current_view = None
        self.show_view("home")
        self.navigation.highlight_selected("home")
        self.bind_all("<MouseWheel>", self._on_global_mousewheel)
        self.bind_all("<Button-4>", self._on_global_mousewheel)
        self.bind_all("<Button-5>", self._on_global_mousewheel)

    def _build_views(self):
        views = {}

        for module in MODULES:
            key = module["key"]
            label = module["label"]
            view_class = module["view_class"]
            enabled = module["enabled"]

            if enabled and view_class is not None:
                views[key] = view_class(self.content_frame)
            else:
                views[key] = PlaceholderView(self.content_frame, label)

        return views

    def show_view(self, key: str):
        if self.current_view is not None:
            self.current_view.grid_forget()

        view = self.views[key]
        view.grid(row=0, column=0, sticky="nsew")
        self.current_view = view

    def _on_global_mousewheel(self, event):
        widget = event.widget

        while widget is not None:
            if hasattr(widget, "_parent_canvas"):
                if event.num == 4:
                    widget._parent_canvas.yview_scroll(-1, "units")
                elif event.num == 5:
                    widget._parent_canvas.yview_scroll(1, "units")
                else:
                    widget._parent_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                return

            widget = widget.master