import customtkinter as ctk

from app.gui.navigation import NavigationPanel
from app.gui.home_view import HomeView
from app.modules.postulacion.ui import PostulacionFrame

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

        self.navigation = NavigationPanel(self, self.show_view)
        self.navigation.grid(row=0, column=0, sticky="nsew")

        self.content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.content_frame.grid(row=0, column=1, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.views = {
            "home": HomeView(self.content_frame),
            "postulacion": PostulacionFrame(self.content_frame),
            "aceptacion": PlaceholderView(self.content_frame, "Aceptación"),
            "actualizacion": PlaceholderView(self.content_frame, "Actualización de datos"),
            "filtrado": PlaceholderView(self.content_frame, "Filtrado"),
            "inscripcion": PlaceholderView(self.content_frame, "Inscripción"),
            "notas": PlaceholderView(self.content_frame, "Notas"),
            "finalizados": PlaceholderView(self.content_frame, "Finalizados"),
        }

        self.current_view = None
        self.show_view("home")
        self.navigation.highlight_selected("home")

    def show_view(self, key: str):
        if self.current_view is not None:
            self.current_view.grid_forget()

        view = self.views[key]
        view.grid(row=0, column=0, sticky="nsew")
        self.current_view = view