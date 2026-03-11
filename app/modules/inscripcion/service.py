from pathlib import Path
import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK

from app.core.file_utils import ensure_folder, sanitize_filename, safe_str
from app.core.validation_utils import (
    validate_file_exists,
    validate_required_columns,
)
from app.core.docx_utils import (
    find_table_by_text,
)


SCHEDULE_TABLE_MARKER = "Horario Teoría nombre_asignatura"
LAB_TABLE_MARKER = "Horario Laboratorio nombre_asignatura"


def load_workbook_sheets(excel_path: str) -> dict[str, pd.DataFrame]:
    return pd.read_excel(excel_path, sheet_name=None)


def get_available_subjects(config: dict) -> list[str]:
    return list(config["subjects"].keys())


def normalize_subject_dataframe(df: pd.DataFrame, config: dict, subject_config: dict) -> pd.DataFrame:
    common_mapping = config["common_column_mapping"]
    specific_mapping = subject_config["specific_column_mapping"]

    full_mapping = {}
    full_mapping.update(common_mapping)
    full_mapping.update(specific_mapping)

    df = df.rename(columns=full_mapping)

    # Normalización interna para que el resto del service use nombres genéricos
    if subject_config["has_catedra"]:
        catedra_field = subject_config["horarios_catedra_field"]
        if catedra_field in df.columns:
            df["HorariosCatedra"] = df[catedra_field]

    if subject_config["has_lab"]:
        lab_field = subject_config["horarios_lab_field"]
        if lab_field in df.columns:
            df["HorariosLaboratorio"] = df[lab_field]

    disponibles_field = subject_config.get("horarios_disponibles_field", "")
    if disponibles_field and disponibles_field in df.columns:
        df["HorariosDisponibles"] = df[disponibles_field]

    return df

def get_required_columns(subject_config: dict) -> list[str]:
    required = [
        "Minor",
        "PrimerNombre",
        "ApellidoPaterno",
        "ApellidoMaterno",
        "RUT",
        "CorreoInstitucional",
        "Carrera",
        "JefeCarrera",
        "DuracionCarrera",
        "AvanceCurricular",
        "Facultad"
    ]

    if subject_config["has_catedra"]:
        required.append("HorariosCatedra")

    if subject_config["has_lab"]:
        required.append("HorariosLaboratorio")

    return required

def is_effectively_empty(df: pd.DataFrame) -> bool:
    if df.empty:
        return True

    temp = df.copy()
    temp = temp.dropna(how="all")

    if temp.empty:
        return True

    return False

def validate_subject_dataframe(df: pd.DataFrame, subject_config: dict):
    required_columns = get_required_columns(subject_config)
    validate_required_columns(df, required_columns)


def build_output_path(base_output_folder: str, subject_name: str, row_data: dict) -> Path:
    subject_folder = sanitize_filename(subject_name)
    carrera = sanitize_filename(safe_str(row_data["Carrera"]) or "SinCarrera")
    first_name = sanitize_filename(safe_str(row_data["PrimerNombre"]))
    last_name_1 = sanitize_filename(safe_str(row_data["ApellidoPaterno"]))
    last_name_2 = sanitize_filename(safe_str(row_data["ApellidoMaterno"]))

    folder = ensure_folder(Path(base_output_folder) / subject_folder / carrera)
    filename = f"formulario_{first_name}_{last_name_1}_{last_name_2}.docx"
    return folder / filename


def replace_in_paragraphs(doc, row_data: dict, subject_name: str, semestre: str, fecha_documento: str):
    for paragraph in doc.paragraphs:
        paragraph.text = paragraph.text.replace("fecha_ingreso", safe_str(fecha_documento))
        paragraph.text = paragraph.text.replace("semestre_ingreso", safe_str(semestre))
        paragraph.text = paragraph.text.replace("nombre_jefe_carrera", safe_str(row_data["JefeCarrera"]))
        paragraph.text = paragraph.text.replace(
            "carrera_estudiante",
            f"Jefatura Carrera {safe_str(row_data['Carrera'])}"
        )
        paragraph.text = paragraph.text.replace("nombre_asignatura", safe_str(subject_name))


def fill_schedule_table_regular(doc, row_data: dict, subject_name: str, subject_config: dict):
    horarios_catedra = subject_config["horarios_catedra"]
    horarios_lab = subject_config["horarios_lab"]
    header_text = subject_config["schedule_header_text"]

    tabla = get_schedule_table(doc, subject_config)

    cabecera_index = None
    for i, row in enumerate(tabla.rows):
        if row.cells[0].text.strip() == header_text:
            cabecera_index = i
            break

    if cabecera_index is None:
        raise ValueError(
            f"No se encontró la fila cabecera de horarios para {subject_name}."
        )

    for i, (hora_catedra, hora_lab) in enumerate(zip(horarios_catedra, horarios_lab)):
        nueva_fila = tabla.add_row()

        if len(nueva_fila.cells) != 5:
            raise ValueError(
                f"La fila agregada en la plantilla de {subject_name} tiene {len(nueva_fila.cells)} celdas y se esperaban 5."
            )

        # Estructura real de proto.docx:
        # [0] teoría horario
        # [1] respuesta teoría
        # [2] celda teoría extendida / apoyo visual
        # [3] laboratorio horario
        # [4] respuesta laboratorio

        # Fusionar las celdas 1 y 2 para que visualmente queden 4 columnas
        nueva_fila.cells[1].merge(nueva_fila.cells[2])

        nueva_fila.cells[0].text = safe_str(hora_catedra)
        nueva_fila.cells[1].text = f"respuesta_catedra_{i+1}"
        nueva_fila.cells[3].text = safe_str(hora_lab)
        nueva_fila.cells[4].text = f"respuesta_lab_{i+1}"

        if nueva_fila.cells[0].paragraphs and nueva_fila.cells[0].paragraphs[0].runs:
            nueva_fila.cells[0].paragraphs[0].runs[0].font.bold = False


def fill_schedule_table_taaa(doc, row_data: dict, subject_name: str, subject_config: dict):
    horarios_lab = subject_config["horarios_lab"]
    header_text = subject_config["schedule_header_text"]

    tabla = get_schedule_table(doc, subject_config)

    cabecera_index = None
    for i, row in enumerate(tabla.rows):
        if row.cells[0].text.strip() == header_text:
            cabecera_index = i
            break

    if cabecera_index is None:
        raise ValueError(
            f"No se encontró la fila cabecera de horarios de laboratorio para {subject_name}."
        )

    for i, hora_lab in enumerate(horarios_lab):
        nueva_fila = tabla.add_row()

        if len(nueva_fila.cells) < 5:
            raise ValueError(
                f"La fila agregada en la plantilla de {subject_name} no tiene 5 celdas como se esperaba."
            )

        nueva_fila.cells[0].merge(nueva_fila.cells[1])
        nueva_fila.cells[0].merge(nueva_fila.cells[2])
        nueva_fila.cells[3].merge(nueva_fila.cells[4])

        nueva_fila.cells[0].text = safe_str(hora_lab)
        if nueva_fila.cells[0].paragraphs and nueva_fila.cells[0].paragraphs[0].runs:
            nueva_fila.cells[0].paragraphs[0].runs[0].font.bold = False

        nueva_fila.cells[3].text = f"respuesta_lab_{i+1}"


def replace_in_tables_regular(doc, row_data: dict, subject_name: str, subject_config: dict):
    horarios_catedra = subject_config["horarios_catedra"]
    horarios_lab = subject_config["horarios_lab"]

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "MM" in cell.text:
                    cell.text = "X" if safe_str(row_data["Minor"]) == "Minor en Ciencia de Datos" else ""

                if "AA" in cell.text:
                    cell.text = "X" if safe_str(row_data["Minor"]) == "Minor en Ciencia de Datos Avanzado" else ""

                if "primer_nombresegundo_nombre" in cell.text:
                    cell.text = f"{safe_str(row_data['PrimerNombre'])} {safe_str(row_data.get('SegundoNombre', ''))}".strip()

                if "primer_apellidosegundo_apellido" in cell.text:
                    cell.text = f"{safe_str(row_data['ApellidoPaterno'])} {safe_str(row_data['ApellidoMaterno'])}".strip()

                if "rut_estudiante" in cell.text:
                    cell.text = safe_str(row_data["RUT"])

                if "correo_estudiante" in cell.text:
                    cell.text = safe_str(row_data["CorreoInstitucional"])

                if "carrera_estudiante" in cell.text:
                    cell.text = safe_str(row_data["Carrera"])

                if "facultad_estudiante" in cell.text:
                    cell.text = safe_str(row_data["Facultad"])

                if "duracion_carrera" in cell.text:
                    cell.text = safe_str(row_data["DuracionCarrera"])

                if "nivel_avance" in cell.text:
                    cell.text = safe_str(row_data["AvanceCurricular"])

                if "Horario Teoría nombre_asignatura" in cell.text:
                    cell.text = f"Horario Teoría {subject_name}"
                    if cell.paragraphs:
                        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                if "Horario Laboratorio nombre_asignatura" in cell.text:
                    cell.text = f"Horario Laboratorio {subject_name}"
                    if cell.paragraphs:
                        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                for i in range(len(horarios_catedra)):
                    if f"respuesta_catedra_{i+1}" in cell.text:
                        cell.text = "X" if safe_str(row_data["HorariosCatedra"]) == safe_str(horarios_catedra[i]) else ""
                        break

                for i in range(len(horarios_lab)):
                    if f"respuesta_lab_{i+1}" in cell.text:
                        cell.text = "X" if safe_str(row_data["HorariosLaboratorio"]) == safe_str(horarios_lab[i]) else ""
                        break


def replace_in_tables_taaa(doc, row_data: dict, subject_name: str, subject_config: dict):
    horarios_lab = subject_config["horarios_lab"]

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "MM" in cell.text:
                    cell.text = "X" if safe_str(row_data["Minor"]) == "Minor en Ciencia de Datos" else ""

                if "AA" in cell.text:
                    cell.text = "X" if safe_str(row_data["Minor"]) == "Minor en Ciencia de Datos Avanzado" else ""

                if "primer_nombresegundo_nombre" in cell.text:
                    cell.text = f"{safe_str(row_data['PrimerNombre'])} {safe_str(row_data.get('SegundoNombre', ''))}".strip()

                if "primer_apellidosegundo_apellido" in cell.text:
                    cell.text = f"{safe_str(row_data['ApellidoPaterno'])} {safe_str(row_data['ApellidoMaterno'])}".strip()

                if "rut_estudiante" in cell.text:
                    cell.text = safe_str(row_data["RUT"])

                if "correo_estudiante" in cell.text:
                    cell.text = safe_str(row_data["CorreoInstitucional"])

                if "carrera_estudiante" in cell.text:
                    cell.text = safe_str(row_data["Carrera"])

                if "facultad_estudiante" in cell.text:
                    cell.text = safe_str(row_data["Facultad"])

                if "duracion_carrera" in cell.text:
                    cell.text = safe_str(row_data["DuracionCarrera"])

                if "nivel_avance" in cell.text:
                    cell.text = safe_str(row_data["AvanceCurricular"])

                if "Horario Laboratorio Taller de Aprendizaje Automático Aplicado" in cell.text:
                    cell.text = "Horario Laboratorio Taller de Aprendizaje Automático Aplicado"
                    if cell.paragraphs:
                        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                for i in range(len(horarios_lab)):
                    if f"respuesta_lab_{i+1}" in cell.text:
                        cell.text = "X" if safe_str(row_data["HorariosLaboratorio"]) == safe_str(horarios_lab[i]) else ""
                        break


def insert_page_break_if_needed(doc):
    for i, paragraph in enumerate(doc.paragraphs):
        if "salto_pagina" in paragraph.text:
            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i + 1].insert_paragraph_before("").add_run().add_break(WD_BREAK.PAGE)
            paragraph.clear()
            break


def generate_document_for_row(
    row_data: dict,
    subject_name: str,
    subject_config: dict,
    template_path: str,
    semestre: str,
    fecha_documento: str,
    save_path: str,
):
    doc = Document(template_path)

    if subject_config["has_catedra"]:
        fill_schedule_table_regular(doc, row_data, subject_name, subject_config)
        replace_in_tables_regular(doc, row_data, subject_name, subject_config)
    else:
        fill_schedule_table_taaa(doc, row_data, subject_name, subject_config)
        replace_in_tables_taaa(doc, row_data, subject_name, subject_config)

    replace_in_paragraphs(doc, row_data, subject_name, semestre, fecha_documento)
    insert_page_break_if_needed(doc)

    try:
        doc._element.body.remove(doc._element.body[0])
    except Exception:
        pass

    doc.save(save_path)


def validate_subject_resources(subject_name: str, config: dict):
    template_path = config["template_paths"][subject_name]
    validate_file_exists(template_path, f"plantilla de {subject_name}")

    subject_config = config["subjects"][subject_name]

    if subject_config["has_catedra"] and not subject_config["horarios_catedra"]:
        raise ValueError(f"La asignatura {subject_name} no tiene horarios de cátedra configurados.")

    if subject_config["has_lab"] and not subject_config["horarios_lab"]:
        raise ValueError(f"La asignatura {subject_name} no tiene horarios de laboratorio configurados.")


def process_subject(
    subject_name: str,
    df: pd.DataFrame,
    config: dict,
    output_folder: str,
    semestre: str,
    fecha_documento: str,
    logger=None,
) -> dict:
    def log(msg: str):
        if logger:
            logger(msg)

    result = {
        "subject": subject_name,
        "total": 0,
        "ok": 0,
        "errors": 0,
        "skipped": False,
        "error_details": [],
    }

    subject_config = config["subjects"][subject_name]
    validate_subject_resources(subject_name, config)

    if is_effectively_empty(df):
        result["skipped"] = True
        log(f"[{subject_name}] Hoja vacía. Se omite.")
        return result

    df = normalize_subject_dataframe(df, config, subject_config)
    validate_subject_dataframe(df, subject_config)
    df = df.dropna(how="all")

    result["total"] = len(df)
    template_path = config["template_paths"][subject_name]

    log(f"[{subject_name}] Registros a procesar: {len(df)}")

    for idx, row in df.iterrows():
        row_data = row.to_dict()

        try:
            save_path = build_output_path(output_folder, subject_name, row_data)
            generate_document_for_row(
                row_data=row_data,
                subject_name=subject_name,
                subject_config=subject_config,
                template_path=template_path,
                semestre=semestre,
                fecha_documento=fecha_documento,
                save_path=str(save_path),
            )
            result["ok"] += 1
            log(f"[{subject_name}] OK fila {idx + 1}: {save_path}")
        except Exception as e:
            result["errors"] += 1
            detail = f"[{subject_name}] Error fila {idx + 1}: {e}"
            result["error_details"].append(detail)
            log(detail)

    return result


def validate_excel_workbook(excel_path: str, config: dict, selected_subjects: list[str]) -> dict:
    validate_file_exists(excel_path, "archivo Excel")

    workbook = load_workbook_sheets(excel_path)
    results = {
        "available_sheets": list(workbook.keys()),
        "missing_sheets": [],
        "empty_sheets": [],
    }

    for subject_name in selected_subjects:
        sheet_name = config["subjects"][subject_name]["sheet_name"]

        if sheet_name not in workbook:
            results["missing_sheets"].append(sheet_name)
            continue

        df = workbook[sheet_name]
        if is_effectively_empty(df):
            results["empty_sheets"].append(sheet_name)

    return results


def process_inscripcion(
    excel_path: str,
    output_folder: str,
    semestre: str,
    fecha_documento: str,
    selected_subjects: list[str],
    config: dict,
    logger=None,
) -> dict:
    def log(msg: str):
        if logger:
            logger(msg)

    validate_file_exists(excel_path, "archivo Excel")
    ensure_folder(output_folder)

    workbook = load_workbook_sheets(excel_path)

    global_result = {
        "subjects_processed": [],
        "subjects_skipped": [],
        "missing_sheets": [],
        "total_ok": 0,
        "total_errors": 0,
        "details": [],
    }

    for subject_name in selected_subjects:
        sheet_name = config["subjects"][subject_name]["sheet_name"]

        if sheet_name not in workbook:
            global_result["missing_sheets"].append(sheet_name)
            log(f"[{subject_name}] Hoja no encontrada en Excel: {sheet_name}")
            continue

        df = workbook[sheet_name]

        try:
            result = process_subject(
                subject_name=subject_name,
                df=df,
                config=config,
                output_folder=output_folder,
                semestre=semestre,
                fecha_documento=fecha_documento,
                logger=log,
            )

            if result["skipped"]:
                global_result["subjects_skipped"].append(subject_name)
            else:
                global_result["subjects_processed"].append(subject_name)

            global_result["total_ok"] += result["ok"]
            global_result["total_errors"] += result["errors"]
            global_result["details"].append(result)

        except Exception as e:
            detail = f"[{subject_name}] Error general de asignatura: {e}"
            global_result["total_errors"] += 1
            global_result["details"].append({
                "subject": subject_name,
                "total": 0,
                "ok": 0,
                "errors": 1,
                "skipped": False,
                "error_details": [detail],
            })
            log(detail)

    return global_result

def get_schedule_table(doc, subject_config: dict):
    table_index = subject_config["table_index"]

    if len(doc.tables) <= table_index:
        raise ValueError(
            f"La plantilla no tiene la tabla esperada en índice {table_index}."
        )

    return doc.tables[table_index]