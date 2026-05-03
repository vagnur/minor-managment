import json
from copy import deepcopy
from pathlib import Path


DEFAULT_CONFIG = {
    "base_output_folder": "output/preparacion_notas",

    "subjects": {
        "FPpCD": {
            "code": "10150",
            "display_name": "Fundamentos de Programación para Ciencia de Datos",
            "aliases": ["fppcd", "fundamentos"],
            "is_taaa": False,
        },
        "ECeI": {
            "code": "10151",
            "display_name": "Estadística Computacional e Inferencial",
            "aliases": ["ecei", "estadistica", "estadística"],
            "is_taaa": False,
        },
        "TIC I": {
            "code": "10152",
            "display_name": "Técnicas de Inteligencia Computacional I",
            "aliases": ["tic i", "tici"],
            "is_taaa": False,
        },
        "TIC II": {
            "code": "10153",
            "display_name": "Técnicas de Inteligencia Computacional II",
            "aliases": ["tic ii", "ticii", "taic"],
            "is_taaa": False,
        },
        "TAAA": {
            "code": "10154",
            "display_name": "Taller de Aprendizaje Automático Aplicado",
            "aliases": ["taaa", "taller"],
            "is_taaa": True,
        },
    },

    "source_columns": {
        "rut_estudiante": "RUT",
        "nombre": "Nombre Estudiante",
        "correo": "Correo institucional",
        "facultad": "Facultad a la que pertenece",
        "carrera": "Carrera a la que pertenece",
        "profesor": "Profesor",
        "seccion_catedra": "Teo",
        "seccion_laboratorio": "Lab"
        }
}


def load_config(config_path: str = "config/preparacion_notas.json") -> dict:
    path = Path(config_path)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return deepcopy(DEFAULT_CONFIG)

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict, config_path: str = "config/preparacion_notas.json") -> None:
    path = Path(config_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)