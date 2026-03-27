from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


HEADER_ROW = 1
TRAINEE_ROLE_COLUMN = 3
MENTOR_ROLE_COLUMN = 8
ROW_HIGHLIGHT_FILL = PatternFill(fill_type="solid", fgColor="80FF0000")

ALLOWED_MENTOR_ROLES_BY_TRAINEE_ROLE = {
    "бариста-стажер": {"бариста"},
    "кассир-стажер": {"кассир", "старший кассир", "повар-универсал"},
    "повар-стажер": {"повар-универсал", "повар"},
    "повар-универсал стажер": {
        "повар-универсал",
        "повар",
        "старший кассир",
        "кассир",
    },
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
    )


def _mentor_role_is_valid(trainee_role: object, mentor_role: object) -> bool:
    normalized_trainee_role = _normalize_role(trainee_role)
    normalized_mentor_role = _normalize_role(mentor_role)

    if normalized_mentor_role == "":
        return False

    allowed_mentor_roles = ALLOWED_MENTOR_ROLES_BY_TRAINEE_ROLE.get(normalized_trainee_role)
    if allowed_mentor_roles is None:
        return True

    return normalized_mentor_role in allowed_mentor_roles


def process_excel(input_path: Path, output_path: Path) -> None:
    """Highlight rows with invalid mentor position rules.

    Row is highlighted in semi-transparent red when:
    1) Mentor role column is empty.
    2) For configured trainee roles, mentor role is outside the allowed list.
    """
    workbook = load_workbook(input_path)

    for sheet in workbook.worksheets:
        for row_idx in range(HEADER_ROW + 1, sheet.max_row + 1):
            trainee_role = sheet.cell(row=row_idx, column=TRAINEE_ROLE_COLUMN).value
            mentor_role = sheet.cell(row=row_idx, column=MENTOR_ROLE_COLUMN).value

            if not _mentor_role_is_valid(trainee_role=trainee_role, mentor_role=mentor_role):
                for col_idx in range(1, sheet.max_column + 1):
                    sheet.cell(row=row_idx, column=col_idx).fill = ROW_HIGHLIGHT_FILL

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
