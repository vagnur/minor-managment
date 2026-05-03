from app.gui.home_view import HomeView
from app.modules.postulacion.ui import PostulacionFrame
from app.modules.aceptacion.ui import AceptacionFrame
from app.modules.inscripcion.ui import InscripcionFrame
from app.modules.notas.ui import NotasFrame
from app.modules.preparacion_notas.ui import PreparacionNotasFrame


MODULES = [
    {
        "key": "home",
        "label": "Inicio",
        "view_class": HomeView,
        "enabled": True,
    },
    {
        "key": "postulacion",
        "label": "Postulación",
        "view_class": PostulacionFrame,
        "enabled": True,
    },
    {
        "key": "aceptacion",
        "label": "Aceptación",
        "view_class": AceptacionFrame,
        "enabled": True,
    },
    {
        "key": "actualizacion",
        "label": "Actualización de datos",
        "view_class": None,
        "enabled": False,
    },
    {
        "key": "filtrado",
        "label": "Filtrado",
        "view_class": None,
        "enabled": False,
    },
    {
        "key": "inscripcion",
        "label": "Inscripción",
        "view_class": InscripcionFrame,
        "enabled": True,
    },
    {
        "key": "notas",
        "label": "Notas",
        "view_class": NotasFrame,
        "enabled": True,
    },
    {
        "key": "finalizados",
        "label": "Finalizados",
        "view_class": None,
        "enabled": False,
    },
    {
        "key": "preparacion_notas",
        "label": "Preparación de notas",
        "view_class": PreparacionNotasFrame,
        "enabled": True,
    },

]