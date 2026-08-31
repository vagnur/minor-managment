from app.gui.home_view import HomeView
from app.modules.postulacion.ui import PostulacionFrame
from app.modules.aceptacion.ui import AceptacionFrame
from app.modules.inscripcion.ui import InscripcionFrame
from app.modules.notas.ui import NotasFrame
from app.modules.preparacion_notas.ui import PreparacionNotasFrame
from app.modules.actualizacion.ui import ActualizacionFrame


NAVIGATION_GROUPS = [
    {
        "key": "inicio_semestre",
        "label": "Inicio de semestre",
    },
    {
        "key": "fin_semestre",
        "label": "Fin de semestre",
    },
    {
        "key": "utilidades",
        "label": "Utilidades",
    },
]


MODULES = [
    {
        "key": "home",
        "label": "Inicio",
        "view_class": HomeView,
        "enabled": True,
    },

    # Inicio de semestre
    {
        "key": "filtrado",
        "label": "Filtrado",
        "group": "inicio_semestre",
        "view_class": None,
        "enabled": False,
    },
    {
        "key": "postulacion",
        "label": "Postulación",
        "group": "inicio_semestre",
        "view_class": PostulacionFrame,
        "enabled": True,
    },
    {
        "key": "aceptacion",
        "label": "Aceptación",
        "group": "inicio_semestre",
        "view_class": AceptacionFrame,
        "enabled": True,
    },
    {
        "key": "inscripcion",
        "label": "Inscripción",
        "group": "inicio_semestre",
        "view_class": InscripcionFrame,
        "enabled": True,
    },

    # Fin de semestre
    {
        "key": "preparacion_notas",
        "label": "Preparación de notas",
        "group": "fin_semestre",
        "view_class": PreparacionNotasFrame,
        "enabled": True,
    },
    {
        "key": "notas",
        "label": "Notas",
        "group": "fin_semestre",
        "view_class": NotasFrame,
        "enabled": True,
    },

    # Utilidades
    {
        "key": "actualizacion",
        "label": "Actualización de datos",
        "group": "utilidades",
        "view_class": ActualizacionFrame,
        "enabled": True,
    },
]