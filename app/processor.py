from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


HEADER_ROW = 1
COLUMN_G = 7
COLUMN_H = 8


def _is_blank(value: object) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _contains_february(value: object) -> bool:
    if _is_blank(value):
        return False
    return "феврал" in str(value).lower()


def process_excel(input_path: Path, output_path: Path) -> None:
    """Process every sheet and remove rows by filtering rules.

    Rules:
    1) Remove rows if column G is empty.
    2) Remove rows if column G contains "февраль".
    3) Remove rows if column H is empty.
    """
    workbook = load_workbook(input_path)

    for sheet in workbook.worksheets:
        rows_to_delete: list[int] = []
        for row_idx in range(sheet.max_row, HEADER_ROW, -1):
            g_value = sheet.cell(row=row_idx, column=COLUMN_G).value
            h_value = sheet.cell(row=row_idx, column=COLUMN_H).value

            if _is_blank(g_value) or _contains_february(g_value) or _is_blank(h_value):
                rows_to_delete.append(row_idx)

        for row_idx in rows_to_delete:
            sheet.delete_rows(row_idx, 1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
