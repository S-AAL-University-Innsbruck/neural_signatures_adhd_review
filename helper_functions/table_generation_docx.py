"""Functions for exporting selected dataframes as Word tables."""

from pathlib import Path

import pandas as pd
from docx import Document
from docx.shared import Pt


def generate_docx_table(
    dataframe: pd.DataFrame,
    output_path: str | Path,
) -> Path:
    """Save the provided dataframe as a styled Word table."""
    # Round values before writing them to the document.
    dataframe = dataframe.round(2)

    # Create the table and add its column names as the header.
    document = Document()
    table = document.add_table(rows=1, cols=len(dataframe.columns))
    table.style = "Table Grid"

    # --- Apply font to the entire table in one pass ---
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(10)
                    
    for cell, column_name in zip(table.rows[0].cells, dataframe.columns):
        cell.text = str(column_name)

    # Add one row for each dataframe record.
    for row in dataframe.itertuples(index=False, name=None):
        cells = table.add_row().cells
        for cell, value in zip(cells, row):
            cell.text = str(value)

    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_file))
    return output_file