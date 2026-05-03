import json
from copy import deepcopy
from pathlib import Path


REGULAR_SUBJECT_BASE = {
    "template_path": "templates/molde_inscripcion.docx",
    "has_catedra": True,
    "has_lab": True,
    "column_mapping": {
        "Seleccione el o los horarios de cátedra a los cuales puede asistir": "HorariosCatedra",
        "Seleccione el horario de cátedra al cual puede asistir": "HorariosCatedra",
        "Seleccione el o los horarios de laboratorio a los cuales puede asistir": "HorariosLaboratorio",
        "Seleccione el horario de laboratorio al cual puede asistir": "HorariosLaboratorio",
        "Indique todos los horarios disponibles en los cuales podría participar en la asignatura, esto es para hacer un catastro entre aquellas/os estudiantes que no pueden participar en el Minor debido a topes de horarios. Nos permite ver posibilidades de solicitar una nueva coordinación en caso de haber demanda suficiente": "HorariosDisponibles"
    },
    "horarios_catedra_field": "HorariosCatedra",
    "horarios_lab_field": "HorariosLaboratorio",
    "horarios_disponibles_field": "HorariosDisponibles",
}

DEFAULT_CONFIG = {
    "base_output_folder": "output/inscripcion",

    "common_column_mapping": {
        "Marca temporal": "Fecha",
        "Dirección de correo electrónico": "Correo",
        "Primer Nombre": "PrimerNombre",
        "Segundo Nombre": "SegundoNombre",
        "Apellido paterno": "ApellidoPaterno",
        "Apellido materno": "ApellidoMaterno",
        "RUT": "RUT",
        "Correo institucional": "CorreoInstitucional",
        "Carrera a la que pertenece": "Carrera",
        "Nombre y apellido de su Jefe Carrera": "JefeCarrera",
        "Correo electrónico de su Jefe de Carrera": "CorreoJefeCarrera",
        "Duración de la carrera": "DuracionCarrera",
        "Avance curricular": "AvanceCurricular",
        "Facultad a la que pertenece": "Facultad",
        "Indique asignatura que desea inscribir": "AsignaturaInscrita",
    },

    "subjects": {
        "FPpCD": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "sheet_name": "FPpCD",
            "display_name": "Fundamentos de Programación para Ciencia de Datos"
        },

        "ECeI": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "sheet_name": "ECeI",
            "display_name": "Estadística Computacional e Inferencial",
        },

        "TIC I": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "sheet_name": "TIC I",
            "display_name": "Técnicas de Inteligencia Computacional I",
        },

        "TIC II": {
            **deepcopy(REGULAR_SUBJECT_BASE),
            "sheet_name": "TIC II",
            "display_name": "Técnicas de Inteligencia Computacional II",
        },

        "TAAA": {
            "sheet_name": "TAAA",
            "display_name": "Taller de Aprendizaje Automático Aplicado",
            "template_path": "templates/molde_inscripcion_TAAA.docx",
            "has_catedra": False,
            "has_lab": True,
            "column_mapping": {
                "Seleccione el o los horarios de laboratorio a los cuales puede asistir": "HorariosLaboratorio",
                "Seleccione el horario de laboratorio al cual puede asistir": "HorariosLaboratorio",
                "Indique todos los horarios disponibles en los cuales podría participar en la asignatura, esto es para hacer un catastro entre aquellas/os estudiantes que no pueden participar en el Minor debido a topes de horarios. Nos permite ver posibilidades de solicitar una nueva coordinación en caso de haber demanda suficiente": "HorariosDisponibles"
            },
            "horarios_catedra_field": "",
            "horarios_lab_field": "HorariosLaboratorio",
            "horarios_disponibles_field": "HorariosDisponibles",
        }
    }
}


def load_config(config_path: str = "config/inscripcion.json") -> dict:
    path = Path(config_path)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return deepcopy(DEFAULT_CONFIG)

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict, config_path: str = "config/inscripcion.json") -> None:
    path = Path(config_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)