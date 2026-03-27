from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


HEADER_ROW = 1
COLUMN_C = 3
COLUMN_F = 6
COLUMN_G = 7
COLUMN_H = 8
TARGET_ROLE = "стажер"
INVALID_ROW_FILL = PatternFill(fill_type="solid", start_color="FFFFC7CE", end_color="FFFFC7CE")

MENTOR_ROLE_RULES: dict[str, set[str]] = {
    "бариста-стажер": {"бариста"},
    "кассир-стажер": {"кассир", "старший кассир", "повар-универсал"},
    "старший кассир-стажер": {"старший кассир", "заместитель директора"},
    "повар-стажер": {"повар-универсал", "повар"},
    "повар-универсал стажер": {"повар-универсал", "повар", "старший кассир", "кассир"},
    "работник торгового зала-стажер": {
        "кассир",
        "старший кассир",
        "работник торгового зала",
        "повар-универсал",
    },
}


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


def _normalize_role(value: object) -> str:
    if _is_blank(value):
        return ""

    return (
        str(value)
        .strip()
        .lower()
        .replace("ё", "е")
        .replace("–", "-")
        .replace("—", "-")
        .replace(" - ", "-")
        .replace("- ", "-")
        .replace(" -", "-")
        .replace("  ", " ")
    )


def _mentor_role_is_valid(trainee_role: object, mentor_role: object) -> bool:
    normalized_trainee_role = _normalize_role(trainee_role)
    allowed_mentor_roles = MENTOR_ROLE_RULES.get(normalized_trainee_role)
    if not allowed_mentor_roles:
        return True

    normalized_mentor_role = _normalize_role(mentor_role)
    if not normalized_mentor_role:
        return False

    return normalized_mentor_role in allowed_mentor_roles


def _row_has_mentor_validation_error(trainee_role: object, mentor_role: object) -> bool:
    """Return True when row violates mentor validation rules.

    Validation rules:
    - mentor role (column F) cannot be empty;
    - for supported trainee roles (column C), mentor role (column F)
      must match role-specific allowed values.
    """
    return _is_blank(mentor_role) or not _mentor_role_is_valid(trainee_role, mentor_role)


def _paint_row(sheet, row_idx: int, max_column: int) -> None:
    for column_idx in range(1, max_column + 1):
        sheet.cell(row=row_idx, column=column_idx).fill = INVALID_ROW_FILL


def process_excel(input_path: Path, output_path: Path) -> None:
    """Process every sheet and remove rows by filtering rules.

    Rules:
    1) Keep rows only if column C contains "стажер".
    2) Remove rows if column G is empty.
    3) Remove rows if column G contains "февраль".
    4) Remove rows if column H is empty.
    5) Fill row red if mentor role in column F is empty.
    6) Fill row red if mentor role in column F does not match trainee role rules.
    """
    workbook = load_workbook(input_path)

    for sheet in workbook.worksheets:
        rows_to_delete: list[int] = []
        rows_to_highlight: list[int] = []
        for row_idx in range(sheet.max_row, HEADER_ROW, -1):
            c_value = sheet.cell(row=row_idx, column=COLUMN_C).value
            g_value = sheet.cell(row=row_idx, column=COLUMN_G).value
            f_value = sheet.cell(row=row_idx, column=COLUMN_F).value
            h_value = sheet.cell(row=row_idx, column=COLUMN_H).value

            if (
                not _contains_target_role(c_value)
                or _is_blank(g_value)
                or _contains_february(g_value)
                or _is_blank(h_value)
            ):
                rows_to_delete.append(row_idx)
                continue

            if _row_has_mentor_validation_error(c_value, f_value):
                rows_to_highlight.append(row_idx)

        for row_idx in rows_to_highlight:
            _paint_row(sheet, row_idx, sheet.max_column)

        for row_idx in rows_to_delete:
            sheet.delete_rows(row_idx, 1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
