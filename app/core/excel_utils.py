import pandas as pd


def read_excel_file(excel_path: str, sheet_name: str):
    return pd.read_excel(excel_path, sheet_name=sheet_name)