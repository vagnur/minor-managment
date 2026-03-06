def find_table_by_text(doc, target_text: str):
    """
    Busca una tabla que contenga un texto específico en cualquiera de sus celdas.
    """
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if target_text in cell.text:
                    return table
    return None


def find_row_index_by_cell_text(table, target_text: str):
    """
    Busca el índice de la fila que contiene el texto indicado.
    """
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            if target_text in cell.text:
                return i
    return None