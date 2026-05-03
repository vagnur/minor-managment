from pathlib import Path
import re

import pandas as pd

from app.core.file_utils import ensure_folder, sanitize_filename, safe_str
from app.core.validation_utils import validate_required_columns


def normalize_text(text: str) -> str:
    return (
        safe_str(text)
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace(".", "")
    )


def detect_subject_from_filename(filename: str, config: dict) -> str | None:
    normalized_filename = normalize_text(filename)

    matches = []

    for subject_name, subject_config in config["subjects"].items():
        for alias in subject_config.get("aliases", []):
            normalized_alias = normalize_text(alias)

            if normalized_alias in normalized_filename:
                matches.append((subject_name, len(normalized_alias)))

    if not matches:
        return None

    matches.sort(key=lambda item: item[1], reverse=True)
    return matches[0][0]


def is_section_sheet(sheet_name: str) -> bool:
    return re.search(r"secci[oó]n", safe_str(sheet_name), re.IGNORECASE) is not None


def extract_section_from_sheet_name(sheet_name: str) -> str:
    text = safe_str(sheet_name).strip()

    match = re.search(r"secci[oó]n\s*([A-Za-z0-9]+)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()

    return text


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = [str(col).strip() for col in df.columns]
    df = df.dropna(how="all")

    source_name_columns = [
        "Nombre Estudiante",
        "Nombre",
    ]

    for column in source_name_columns:
        if column in df.columns:
            df = df[df[column].notna()]
            df = df[df[column].astype(str).str.strip() != ""]
            break

    return df


def get_required_source_columns(config: dict) -> list[str]:
    source_columns = config["source_columns"]

    return [
        source_columns["rut_estudiante"],
        source_columns["nombre"],
        source_columns["correo"],
        source_columns["facultad"],
        source_columns["carrera"],
        source_columns["profesor"],
    ]


def read_source_sections(folder_path: str, config: dict, logger=None) -> list[dict]:
    def log(msg: str):
        if logger:
            logger(msg)

    folder = Path(folder_path)

    if not folder.exists():
        raise FileNotFoundError(f"No existe la carpeta seleccionada: {folder}")

    if not folder.is_dir():
        raise NotADirectoryError(f"La ruta seleccionada no corresponde a una carpeta: {folder}")

    excel_files = sorted([
        file_path
        for file_path in folder.iterdir()
        if file_path.suffix.lower() in [".xlsx", ".xls"]
        and not file_path.name.startswith("~$")
    ])

    if not excel_files:
        raise ValueError("La carpeta seleccionada no contiene archivos Excel.")

    required_columns = get_required_source_columns(config)
    sections = []

    for file_path in excel_files:
        subject_name = detect_subject_from_filename(file_path.name, config)

        if subject_name is None:
            log(f"[OMITIDO] No se pudo detectar asignatura: {file_path.name}")
            continue

        workbook = pd.ExcelFile(file_path)
        section_sheets = [
            sheet_name
            for sheet_name in workbook.sheet_names
            if is_section_sheet(sheet_name)
        ]

        if not section_sheets:
            log(f"[OMITIDO] {file_path.name}: no tiene hojas de sección.")
            continue

        for sheet_name in section_sheets:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df = clean_dataframe(df)

            if df.empty:
                log(f"[OMITIDO] {file_path.name} / {sheet_name}: hoja vacía.")
                continue

            validate_required_columns(df, required_columns)

            section_name = extract_section_from_sheet_name(sheet_name)
            professor_column = config["source_columns"]["profesor"]
            professor_name = safe_str(df.iloc[0][professor_column])

            sections.append({
                "file_path": file_path,
                "file_name": file_path.name,
                "sheet_name": sheet_name,
                "section_name": section_name,
                "subject": subject_name,
                "subject_config": config["subjects"][subject_name],
                "professor_name": professor_name,
                "student_count": len(df),
                "dataframe": df,
            })

            log(
                f"[OK] {file_path.name} / {sheet_name} → "
                f"{subject_name} sección {section_name} → "
                f"{professor_name} → {len(df)} estudiantes"
            )

    if not sections:
        raise ValueError("No se pudo leer ninguna sección válida desde los archivos originales.")

    return sections


def build_professor_rut_requirements(sections: list[dict]) -> list[dict]:
    seen = set()
    requirements = []

    for section in sections:
        subject = section["subject"]
        section_name = section["section_name"]
        professor_name = section["professor_name"]
        is_taaa = section["subject_config"].get("is_taaa", False)

        key = (
            subject,
            section_name,
            professor_name,
            "lab" if is_taaa else "catedra",
        )

        if key in seen:
            continue

        seen.add(key)

        requirements.append({
            "key": "|".join(key),
            "subject": subject,
            "section_name": section_name,
            "professor_name": professor_name,
            "rut_type": "laboratorio" if is_taaa else "cátedra",
            "is_taaa": is_taaa,
        })

    return requirements


def preview_source_folder(folder_path: str, config: dict, logger=None) -> dict:
    sections = read_source_sections(folder_path, config, logger)
    requirements = build_professor_rut_requirements(sections)

    return {
        "sections": sections,
        "requirements": requirements,
        "total_sections": len(sections),
        "total_students": sum(section["student_count"] for section in sections),
        "total_ruts_required": len(requirements),
    }


def build_rut_key(section: dict) -> str:
    is_taaa = section["subject_config"].get("is_taaa", False)

    return "|".join([
        section["subject"],
        section["section_name"],
        section["professor_name"],
        "lab" if is_taaa else "catedra",
    ])


def build_prepared_dataframe(section: dict, professor_rut: str, config: dict) -> pd.DataFrame:
    df = section["dataframe"]
    source_columns = config["source_columns"]

    subject_config = section["subject_config"]
    is_taaa = subject_config.get("is_taaa", False)

    output_rows = []

    for _, row in df.iterrows():
        if is_taaa:
            output_rows.append({
                "Código": subject_config["code"],
                "Sección Laboratorio": row.get(source_columns["seccion_laboratorio"], ""),
                "Profesor Laboratorio": section["professor_name"],
                "RUT Profesor Laboratorio": professor_rut,
                "RUT Estudiante": row.get(source_columns["rut_estudiante"], ""),
                "Nombre": row.get(source_columns["nombre"], ""),
                "Correo": row.get(source_columns["correo"], ""),
                "Facultad": row.get(source_columns["facultad"], ""),
                "Carrera": row.get(source_columns["carrera"], ""),
                "Nota Laboratorio": "",
                "Promedio": "",
            })
        else:
            output_rows.append({
                "Código": subject_config["code"],
                "Sección Cátedra": row.get(source_columns["seccion_catedra"], ""),
                "Sección Laboratorio": row.get(source_columns["seccion_laboratorio"], ""),
                "Profesor Cátedra": section["professor_name"],
                "RUT Profesor Cátedra": professor_rut,
                "RUT Estudiante": row.get(source_columns["rut_estudiante"], ""),
                "Nombre": row.get(source_columns["nombre"], ""),
                "Correo": row.get(source_columns["correo"], ""),
                "Facultad": row.get(source_columns["facultad"], ""),
                "Carrera": row.get(source_columns["carrera"], ""),
                "Nota Cátedra": "",
                "Nota Laboratorio": "",
                "Promedio": "",
            })

    if is_taaa:
        columns = [
            "Código",
            "Sección Laboratorio",
            "Profesor Laboratorio",
            "RUT Profesor Laboratorio",
            "RUT Estudiante",
            "Nombre",
            "Correo",
            "Facultad",
            "Carrera",
            "Nota Laboratorio",
            "Promedio",
        ]
    else:
        columns = [
            "Código",
            "Sección Cátedra",
            "Sección Laboratorio",
            "Profesor Cátedra",
            "RUT Profesor Cátedra",
            "RUT Estudiante",
            "Nombre",
            "Correo",
            "Facultad",
            "Carrera",
            "Nota Cátedra",
            "Nota Laboratorio",
            "Promedio",
        ]

    return pd.DataFrame(output_rows, columns=columns)

def build_output_filename(section: dict) -> str:
    subject = sanitize_filename(section["subject"])
    section_name = sanitize_filename(section["section_name"])

    return f"{subject}_{section_name}.xlsx"


def generate_prepared_grade_excels(
    folder_path: str,
    output_folder: str,
    professor_ruts: dict,
    config: dict,
    logger=None,
) -> dict:
    def log(msg: str):
        if logger:
            logger(msg)

    sections = read_source_sections(folder_path, config, logger)
    output_path = ensure_folder(Path(output_folder) / "notas_preparadas")

    result = {
        "total_ok": 0,
        "total_errors": 0,
        "generated_files": [],
        "error_details": [],
    }

    for section in sections:
        try:
            rut_key = build_rut_key(section)
            professor_rut = safe_str(professor_ruts.get(rut_key, ""))

            if not professor_rut:
                raise ValueError(
                    f"No se ingresó RUT para {section['subject']} "
                    f"sección {section['section_name']} - {section['professor_name']}"
                )

            df_output = build_prepared_dataframe(
                section=section,
                professor_rut=professor_rut,
                config=config,
            )

            file_name = build_output_filename(section)
            save_path = output_path / file_name

            df_output.to_excel(save_path, index=False)

            result["total_ok"] += 1
            result["generated_files"].append(str(save_path))
            log(f"[OK] Excel preparado generado: {save_path}")

        except Exception as e:
            result["total_errors"] += 1
            detail = (
                f"[ERROR] {section['subject']} sección "
                f"{section['section_name']}: {e}"
            )
            result["error_details"].append(detail)
            log(detail)

    return result