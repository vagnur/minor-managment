from pathlib import Path
import re


def ensure_folder(path: str | Path) -> Path:
    folder = Path(path)
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def sanitize_filename(text: str) -> str:
    text = str(text).strip()
    text = text.replace(" ", "_")
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    return text


def safe_str(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "nan":
        return ""
    return text