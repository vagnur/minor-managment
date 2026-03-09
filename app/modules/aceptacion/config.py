import json
from pathlib import Path


DEFAULT_CONFIG = {
    "template_path": "templates/molde_decano.docx",
    "base_output_folder": "output/aceptacion",
    "output_filename_pattern": "ingreso_minor_ciencia_datos_{semestre}_{anio}.docx",
    "table_headers": [
        "Nº",
        "RUN",
        "DV",
        "APELLIDOS",
        "NOMBRES",
        "CARRERA DE ORIGEN",
        "FACULTAD"
    ],
    "column_mapping": {
        "RUT": "RUT",
        "Nombre Estudiante": "NombreEstudiante",
        "Carrera a la que pertenece": "Carrera",
        "Facultad a la que pertenece": "Facultad"
    }
}


def load_config(config_path: str = "config/aceptacion.json") -> dict:
    path = Path(config_path)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG.copy()

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict, config_path: str = "config/aceptacion.json") -> None:
    path = Path(config_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)