import json
from pathlib import Path


DEFAULT_CONFIG = {
    "template_path": "templates/proto.docx",
    "sheet_name": "Hoja 1",
    "base_output_folder": "output/formularios TIC I",
    "fecha_documento": "11/08/2025",
    "nombre_asignatura": "Técnicas de Inteligencia Computacional I",
    "semestre": "2-2025",
    "horarios_catedra": ["L7 W7", "M7 J7"],
    "horarios_lab": ["L7 W7","M3 W2", "J7 V6",]
}


def load_config(config_path: str = "config/postulacion.json") -> dict:
    path = Path(config_path)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG.copy()

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict, config_path: str = "config/postulacion.json") -> None:
    path = Path(config_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)