from copy import deepcopy


DEFAULT_CONFIG = {
    "subjects": {
        "FPpCD": {
            "code": "10150",
            "master_column": "FppCD",
            "aliases": ["fppcd", "fundamentos"],
            "entry_subject": True,
        },
        "ECeI": {
            "code": "10151",
            "master_column": "EceI",
            "aliases": ["ecei", "estadistica", "estadística"],
            "entry_subject": False,
        },
        "TIC I": {
            "code": "10152",
            "master_column": "TIC I",
            "aliases": ["tic i", "tici"],
            "entry_subject": True,
        },
        "TIC II": {
            "code": "10153",
            "master_column": "TIC II",
            "aliases": ["tic ii", "ticii", "taic"],
            "entry_subject": False,
        },
        "TAAA": {
            "code": "10154",
            "master_column": "TAAA",
            "aliases": ["taaa", "taller"],
            "entry_subject": False,
        },
    },
    "master_columns": {
        "name": ["v", "Nombre", "Nombre Estudiante"],
        "rut": ["Rut", "RUT"],
        "email": ["Correo", "Correo institucional"],
        "minor": ["Minor"],
        "career": ["Carrera"],
        "faculty": ["Facultad"],
        "entry": ["Ingreso"],
        "status": ["Estado"],
        "comment": ["Comentario"],
    },
    "grade_columns": {
        "rut": "RUT Estudiante",
        "name": "Nombre",
        "email": "Correo",
        "faculty": "Facultad",
        "career": "Carrera",
        "grade": "Promedio",
    },
    "faculty_mapping": {
        "facultad de ingenieria": "FING",
        "facultad de ingeniería": "FING",
        "ingenieria": "FING",
        "ingeniería": "FING",
        "fing": "FING",
        "facultad de administracion y economia": "FAE",
        "facultad de administración y economía": "FAE",
        "administracion y economia": "FAE",
        "administración y economía": "FAE",
        "fae": "FAE",
        "facultad de ciencia": "FCi",
        "facultad de ciencias": "FCi",
        "ciencia": "FCi",
        "fci": "FCi",
        "facultad de quimica y biologia": "FQyB",
        "facultad de química y biología": "FQyB",
        "quimica y biologia": "FQyB",
        "química y biología": "FQyB",
        "fqyb": "FQyB",
        "facultad de humanidades": "FAHU",
        "humanidades": "FAHU",
        "fahu": "FAHU",
        "facultad de ciencias medicas": "FCM",
        "facultad de ciencias médicas": "FCM",
        "ciencias medicas": "FCM",
        "ciencias médicas": "FCM",
        "fcm": "FCM",
        "facultad tecnologica": "Tecno",
        "facultad tecnológica": "Tecno",
        "tecnologica": "Tecno",
        "tecnológica": "Tecno",
        "tecno": "Tecno",
        "escuela de arquitectura": "Arq",
        "arquitectura": "Arq",
        "arq": "Arq",
    },
}


def load_config() -> dict:
    """Devuelve una copia independiente de la configuración del módulo."""
    return deepcopy(DEFAULT_CONFIG)
