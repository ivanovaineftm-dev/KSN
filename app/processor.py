from __future__ import annotations

from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
import re

import pandas as pd


COLUMN_C = 2
COLUMN_D = 3
COLUMN_F = 5
COLUMN_H = 7
TARGET_ROLE = "стажер"
NORMALIZED_DEPARTMENT_COLUMN = "Подразделение (нормализованное)"
NOT_FOUND_VALUE = "Не найдено"

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


class ProcessingError(Exception):
    """Domain-specific processing error with user friendly text."""


def _engine_by_extension(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    raise ProcessingError("Поддерживаются только файлы .xlsx и .xls")


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


def _normalize_role(value: object) -> str:
    if _is_blank(value):
        return ""

    normalized = (
        str(value)
        .strip()
        .lower()
        .replace("ё", "е")
        .replace("–", "-")
        .replace("—", "-")
        .replace(" - ", "-")
        .replace("- ", "-")
        .replace(" -", "-")
    )
    return re.sub(r"\s+", " ", normalized)


def _contains_target_role(value: object) -> bool:
    return TARGET_ROLE in _normalize_role(value)


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
    return _is_blank(mentor_role) or not _mentor_role_is_valid(trainee_role, mentor_role)


def _normalize_department_key(value: object) -> str:
    if _is_blank(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def _read_excel_any(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path, header=None, engine=_engine_by_extension(path))
    except Exception as exc:  # noqa: BLE001
        raise ProcessingError(f"Не удалось прочитать файл '{path.name}'. Проверьте формат и целостность файла.") from exc


def _extract_locations(locations_df: pd.DataFrame) -> list[str]:
    if locations_df.shape[1] <= 1:
        raise ProcessingError("В файле 'Локации' отсутствует столбец B с подразделениями")

    values = [str(v).strip() for v in locations_df.iloc[:, 1].tolist() if not _is_blank(v)]
    return [value for value in values if value]


def _find_department_match(department: object, canonical_values: list[str]) -> str:
    if not canonical_values or _is_blank(department):
        return NOT_FOUND_VALUE

    source = _normalize_department_key(department)
    normalized_map = {_normalize_department_key(item): item for item in canonical_values}

    exact = normalized_map.get(source)
    if exact:
        return exact

    normalized_candidates = list(normalized_map.keys())
    contains_matches = [key for key in normalized_candidates if source in key or key in source]
    if contains_matches:
        best_key = sorted(contains_matches, key=lambda item: (abs(len(item) - len(source)), len(item)))[0]
        return normalized_map[best_key]

    fuzzy_scores = [
        (candidate, SequenceMatcher(None, source, candidate).ratio())
        for candidate in normalized_candidates
    ]
    fuzzy_scores.sort(key=lambda item: item[1], reverse=True)
    if fuzzy_scores and fuzzy_scores[0][1] >= 0.8:
        return normalized_map[fuzzy_scores[0][0]]

    return NOT_FOUND_VALUE


def process_excel(
    input_path: Path,
    output_path: Path,
    locations_path: Path | None = None,
) -> list[dict[str, int | str]]:
    main_df = _read_excel_any(input_path)
    if main_df.shape[1] <= COLUMN_H:
        raise ProcessingError("В основном файле отсутствуют необходимые столбцы C, D, F или H")

    locations_values: list[str] = []
    if locations_path:
        locations_df = _read_excel_any(locations_path)
        locations_values = _extract_locations(locations_df)

    filtered_df = main_df[
        main_df.iloc[:, COLUMN_C].apply(_contains_target_role)
        & ~main_df.iloc[:, COLUMN_H].apply(_is_blank)
        & ~main_df.iloc[:, COLUMN_H].apply(_contains_february)
    ].copy()

    validation_errors = filtered_df.apply(
        lambda row: _row_has_mentor_validation_error(row.iloc[COLUMN_C], row.iloc[COLUMN_F]),
        axis=1,
    )

    filtered_df[NORMALIZED_DEPARTMENT_COLUMN] = filtered_df.iloc[:, COLUMN_D].apply(
        lambda value: _find_department_match(value, locations_values)
    )

    department_stats: dict[str, dict[str, int | str]] = {}
    for row_idx, row in filtered_df.iterrows():
        department_value = row.iloc[COLUMN_D]
        department_key = _normalize_department_key(department_value)
        if not department_key:
            continue

        stats = department_stats.setdefault(
            department_key,
            {"department": str(department_value).strip().upper(), "total_rows": 0, "valid_rows": 0},
        )
        stats["total_rows"] += 1
        if not bool(validation_errors.loc[row_idx]):
            stats["valid_rows"] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    filtered_df.to_excel(output_path, index=False, engine="openpyxl")

    analytics: list[dict[str, int | str]] = []
    for stats in department_stats.values():
        total_rows = int(stats["total_rows"])
        valid_rows = int(stats["valid_rows"])
        quality = round((valid_rows / total_rows) * 100) if total_rows else 0
        analytics.append(
            {
                "department": str(stats["department"]),
                "total_rows": total_rows,
                "valid_rows": valid_rows,
                "quality": quality,
            }
        )

    analytics.sort(key=lambda item: (-int(item["quality"]), str(item["department"])))
    return analytics
