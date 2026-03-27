from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import re

import pandas as pd
from openpyxl import load_workbook


HEADER_ROW = 1
COLUMN_C = 3
COLUMN_G = 7
COLUMN_H = 8
TARGET_ROLE = "стажер"


def _is_blank(value: object) -> bool:
    if pd.isna(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def _contains_february(value: object) -> bool:
    if _is_blank(value):
        return False

    if isinstance(value, (datetime, date)):
        return value.month == 2

    value_text = str(value).strip().lower()
    if "феврал" in value_text:
        return True

    return re.search(r"(?:^|\D)\d{1,2}\.02\.\d{4}(?:$|\D)", value_text) is not None


def _contains_target_role(value: object) -> bool:
    if _is_blank(value):
        return False

    value_text = str(value).strip().lower().replace("ё", "е")
    normalized_text = value_text.replace("–", "-").replace("—", "-")

    return TARGET_ROLE in normalized_text


def process_excel(input_path: Path, output_path: Path) -> None:
    """Process every sheet and remove rows by filtering rules.

    Rules:
    1) Keep rows only if column C contains "стажер".
    2) Remove rows if column G is empty.
    3) Remove rows if column G contains "февраль".
    4) Remove rows if column H is empty.
    """
    workbook = load_workbook(input_path)

    for sheet in workbook.worksheets:
        rows_to_delete: list[int] = []
        for row_idx in range(sheet.max_row, HEADER_ROW, -1):
            c_value = sheet.cell(row=row_idx, column=COLUMN_C).value
            g_value = sheet.cell(row=row_idx, column=COLUMN_G).value
            h_value = sheet.cell(row=row_idx, column=COLUMN_H).value

            if (
                not _contains_target_role(c_value)
                or _is_blank(g_value)
                or _contains_february(g_value)
                or _is_blank(h_value)
            ):
                rows_to_delete.append(row_idx)

        for row_idx in rows_to_delete:
            sheet.delete_rows(row_idx, 1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
