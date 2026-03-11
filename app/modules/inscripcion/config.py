import json
from pathlib import Path


DEFAULT_CONFIG = {
    "base_output_folder": "output/inscripcion",

    "template_paths": {
        "FPpCD": "templates/proto_inscripcion.docx",
        "ECeI": "templates/proto_inscripcion.docx",
        "TIC I": "templates/proto_inscripcion.docx",
        "TIC II": "templates/proto_inscripcion.docx",
        "TAAA": "templates/proto_TAAA.docx"
    },

    "common_column_mapping": {
        "Marca temporal": "Fecha",
        "Dirección de correo electrónico": "Correo",
        "Minor al que pertenece": "Minor",
        "Primer Nombre": "PrimerNombre",
        "Segundo Nombre": "SegundoNombre",
        "Apellido paterno": "ApellidoPaterno",
        "Apellido materno": "ApellidoMaterno",
        "RUT": "RUT",
        "Número de celular o de contacto": "NumeroCelular",
        "Correo institucional": "CorreoInstitucional",
        "Correo personal (diferente al institucional)": "CorreoPersonal",
        "Carrera a la que pertenece": "Carrera",
        "Nombre y apellido de su Jefe Carrera": "JefeCarrera",
        "Correo electrónico de su Jefe de Carrera": "CorreoJefeCarrera",
        "Duración de la carrera": "DuracionCarrera",
        "Avance curricular": "AvanceCurricular",
        "Facultad a la que pertenece": "Facultad",
        "Indique asignatura que desea inscribir": "AsignaturaInscrita",
        "Comentarios adicionales": "Comentarios"
    },

    "subjects": {
        "FPpCD": {
            "sheet_name": "FPpCD",
            "has_catedra": True,
            "has_lab": True,
            "specific_column_mapping": {
                "Seleccione el o los horarios de cátedra a los cuales puede asistir": "HorariosCatedraFPpCD",
                "Seleccione el o los horarios de laboratorio a los cuales puede asistir": "HorariosLaboratorioFPpCD",
                "Indique todos los horarios disponibles en los cuales podría participar en la asignatura, esto es para hacer un catastro entre aquellas/os estudiantes que no pueden participar en el Minor debido a topes de horarios. Nos permite ver posibilidades de solicitar una nueva coordinación en caso de haber demanda suficiente": "HorariosDisponiblesFPpCD"
            },
            "horarios_catedra_field": "HorariosCatedraFPpCD",
            "horarios_lab_field": "HorariosLaboratorioFPpCD",
            "horarios_disponibles_field": "HorariosDisponiblesFPpCD",
            "horarios_catedra": [
                "L3-L4",
                "M3-M4"
            ],
            "horarios_lab": [
                "J1-J2",
                "V1-V2"
            ]
        },

        "ECeI": {
            "sheet_name": "ECeI",
            "has_catedra": True,
            "has_lab": True,
            "specific_column_mapping": {
                "Seleccione el horario de cátedra al cual puede asistir": "HorariosCatedraECeI",
                "Seleccione el horario de laboratorio al cual puede asistir": "HorariosLaboratorioECeI",
                "Indique todos los horarios disponibles en los cuales podría participar en la asignatura, esto es para hacer un catastro entre aquellas/os estudiantes que no pueden participar en el Minor debido a topes de horarios. Nos permite ver posibilidades de solicitar una nueva coordinación en caso de haber demanda suficiente": "HorariosDisponiblesECeI"
            },
            "horarios_catedra_field": "HorariosCatedraECeI",
            "horarios_lab_field": "HorariosLaboratorioECeI",
            "horarios_disponibles_field": "HorariosDisponiblesECeI",
            "horarios_catedra": [
                "L5-L6",
                "M5-M6"
            ],
            "horarios_lab": [
                "J3-J4",
                "V3-V4"
            ]
        },

        "TIC I": {
            "sheet_name": "TIC I",
            "has_catedra": True,
            "has_lab": True,
            "specific_column_mapping": {
                "Seleccione el o los horarios de cátedra a los cuales puede asistir": "HorariosCatedraTICI",
                "Seleccione el o los horarios de laboratorio a los cuales puede asistir": "HorariosLaboratorioTICI",
                "Indique todos los horarios disponibles en los cuales podría participar en la asignatura, esto es para hacer un catastro entre aquellas/os estudiantes que no pueden participar en el Minor debido a topes de horarios. Nos permite ver posibilidades de solicitar una nueva coordinación en caso de haber demanda suficiente": "HorariosDisponiblesTICI"
            },
            "horarios_catedra_field": "HorariosCatedraTICI",
            "horarios_lab_field": "HorariosLaboratorioTICI",
            "horarios_disponibles_field": "HorariosDisponiblesTICI",
            "horarios_catedra": [
                "M1-M2",
                "X1-X2"
            ],
            "horarios_lab": [
                "J5-J6",
                "V5-V6"
            ]
        },

        "TIC II": {
            "sheet_name": "TIC II",
            "has_catedra": True,
            "has_lab": True,
            "specific_column_mapping": {
                "Seleccione el horario de cátedra al cual puede asistir": "HorariosCatedraTICII",
                "Seleccione el horario de laboratorio al cual puede asistir": "HorariosLaboratorioTICII",
                "Indique todos los horarios disponibles en los cuales podría participar en la asignatura, esto es para hacer un catastro entre aquellas/os estudiantes que no pueden participar en el Minor debido a topes de horarios. Nos permite ver posibilidades de solicitar una nueva coordinación en caso de haber demanda suficiente": "HorariosDisponiblesTICII"
            },
            "horarios_catedra_field": "HorariosCatedraTICII",
            "horarios_lab_field": "HorariosLaboratorioTICII",
            "horarios_disponibles_field": "HorariosDisponiblesTICII",
            "horarios_catedra": [
                "X3-X4",
                "J3-J4"
            ],
            "horarios_lab": [
                "V7-V8",
                "L7-L8"
            ]
        },

        "TAAA": {
            "sheet_name": "TAAA",
            "has_catedra": False,
            "has_lab": True,
            "specific_column_mapping": {
                "Seleccione el o los horarios de laboratorio a los cuales puede asistir": "HorariosLaboratorioTAAA",
                "Indique todos los horarios disponibles en los cuales podría participar en la asignatura, esto es para hacer un catastro entre aquellas/os estudiantes que no pueden participar en el Minor debido a topes de horarios. Nos permite ver posibilidades de solicitar una nueva coordinación en caso de haber demanda suficiente": "HorariosDisponiblesTAAA"
            },
            "horarios_catedra_field": "",
            "horarios_lab_field": "HorariosLaboratorioTAAA",
            "horarios_disponibles_field": "HorariosDisponiblesTAAA",
            "horarios_catedra": [],
            "horarios_lab": [
                "L1-L2",
                "M1-M2",
                "X1-X2"
            ]
        }
    }
}


def load_config(config_path: str = "config/inscripcion.json") -> dict:
    path = Path(config_path)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG.copy()

    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict, config_path: str = "config/inscripcion.json") -> None:
    path = Path(config_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)