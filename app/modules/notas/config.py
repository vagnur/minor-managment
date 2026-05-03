import json
from copy import deepcopy
from pathlib import Path


REGULAR_SUBJECT_BASE = {
    "has_catedra": True,
    "has_lab": True,
    "template_path": "templates/molde_notas.docx",
    "required_columns": [
        "Nombre",
        "RUT Estudiante",
        "Carrera",
        "Facultad",
        "Sección Cátedra",
        "Profesor Cátedra",
        "RUT Profesor Cátedra",
        "Sección Laboratorio",
        "Nota Cátedra",
        "Nota Laboratorio",
        "Promedio",
    ],
}


DEFAULT_CONFIG = {
    "base_output_folder": "output/notas",
    "regular_template_path": "templates/molde_notas.docx",
    "taaa_template_path": "templates/molde_notas_TAAA.docx",

    "subjects": {
        "FPpCD": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "code": "10150",
            "display_name": "Fundamentos de Programación para Ciencia de Datos",
            "aliases": ["fppcd", "fundamentos"],
        },
        "ECeI": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "code": "10151",
            "display_name": "Estadística Computacional e Inferencial",
            "aliases": ["ecei", "estadistica", "estadística"],
        },
        "TIC I": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "code": "10152",
            "display_name": "Técnicas de Inteligencia Computacional I",
            "aliases": ["tic i", "tici"],
        },
        "TIC II": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "code": "10153",
            "display_name": "Técnicas de Inteligencia Computacional II",
            "aliases": ["tic ii", "ticii", "taic"],
        },
        "TAAA": {
            "code": "10154",
            "display_name": "Taller de Aprendizaje Automático Aplicado",
            "aliases": ["taaa", "taller"],
            "has_catedra": False,
            "has_lab": True,
            "template_path": "templates/molde_notas_TAAA.docx",
            "required_columns": [
                "Nombre",
                "RUT Estudiante",
                "Carrera",
                "Facultad",
                "Profesor Laboratorio",
                "RUT Profesor Laboratorio",
                "Sección Laboratorio",
                "Nota Laboratorio",
                "Promedio",
            ],
        },
    },
}


def load_config(config_path: str = "config/notas.json") -> dict:
    path = Path(config_path)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return deepcopy(DEFAULT_CONFIG)

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict, config_path: str = "config/notas.json") -> None:
    path = Path(config_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)