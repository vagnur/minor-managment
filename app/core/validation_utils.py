from pathlib import Path


def validate_file_exists(path: str, label: str = "archivo") -> None:
    if not Path(path).exists():
        raise FileNotFoundError(f"No se encontró el {label}: {path}")


def validate_required_columns(df, required_columns: list[str]) -> None:
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(
            "El archivo Excel no contiene las columnas requeridas: "
            + ", ".join(missing)
        )


def validate_non_empty_dataframe(df) -> None:
    if df.empty:
        raise ValueError("El archivo Excel no contiene registros.")