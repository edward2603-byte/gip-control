"""Read-only validation and preview for the seven-column calendar-plan source.

No function in this module writes the source file or a GIP project workbook.
The update operation is deliberately implemented separately, after a preview
and an explicit confirmation.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

sys.path.insert(0, str(Path(__file__).with_name("vendor")))
import openpyxl
import xlrd


MIN_STAGES = 2
MAX_STAGES = 75
EXPECTED = (
    "Номер договора/Номер Наряд-Заказа", "Наименование проекта", "№ Этапа",
    "Наименование этапа", "Начало работ", "Окончание работ", "Стоимость этапа, руб без НДС",
)
ALIASES = (
    {"номер договора/номер наряд-заказа", "номер договора"}, {"наименование проекта"},
    {"№ этапа", "этап"}, {"наименование этапа", "наименование этапа по договору"},
    {"начало работ"}, {"окончание работ"}, {"стоимость этапа, руб без ндс"},
)


def norm(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\n", " ")).strip().casefold()


def rows_from_book(path: Path) -> tuple[list[list[object]], int | None]:
    suffix = path.suffix.casefold()
    if suffix == ".xls":
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        return [[sheet.cell_value(row, col) for col in range(sheet.ncols)] for row in range(sheet.nrows)], book.datemode
    if suffix == ".xlsx":
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            sheet = book.active
            return [list(row) for row in sheet.iter_rows(values_only=True)], None
        finally:
            book.close()
    raise ValueError("Поддерживаются только .xls и .xlsx")


def parse_date(value: object, datemode: int | None) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and datemode is not None:
        try:
            return xlrd.xldate_as_datetime(value, datemode).date()
        except (TypeError, ValueError, xlrd.XLDateError) as exc:
            raise ValueError(f"некорректная дата: {value!r}") from exc
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, str):
        cleaned = value.strip()
        for pattern in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d"):
            try:
                return datetime.strptime(cleaned, pattern).date()
            except ValueError:
                pass
    raise ValueError(f"некорректная дата: {value!r}")


def parse_money(value: object) -> Decimal | None:
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    cleaned = str(value).replace("\u00a0", " ").replace("₽", "").replace("руб.", "").replace("руб", "")
    cleaned = cleaned.replace(" ", "").replace(",", ".")
    try:
        return Decimal(cleaned).quantize(Decimal("0.01"))
    except InvalidOperation as exc:
        raise ValueError(f"некорректная стоимость: {value!r}") from exc


def header_index(rows: list[list[object]]) -> int:
    """Locate the strictly ordered seven-field header, allowing a short note above it."""
    for index, row in enumerate(rows[:30]):
        candidate = [norm(cell) for cell in row[:7]]
        if len(candidate) == 7 and all(candidate[col] in ALIASES[col] for col in range(7)):
            if any(norm(cell) for cell in row[7:]):
                raise ValueError("Шапка КП должна содержать ровно семь колонок; обнаружены данные после седьмой колонки.")
            return index
    raise ValueError("Шапка не соответствует семи колонкам КП в заданном порядке; импорт не применён.")


def is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def preview(path: Path) -> dict:
    rows, datemode = rows_from_book(path)
    if not rows:
        raise ValueError("Книга пуста")
    header_row = header_index(rows)
    stages: list[dict] = []
    warnings: list[str] = []
    project_refs: set[tuple[str, str]] = set()

    for row_no, raw in enumerate(rows[header_row + 1 :], header_row + 2):
        if any(not is_blank(cell) for cell in raw[7:]):
            raise ValueError(f"Строка {row_no}: обнаружены данные после седьмой колонки КП.")
        row = list(raw[:7]) + [None] * max(0, 7 - len(raw))
        if all(is_blank(cell) for cell in row):
            continue
        stage = str(row[2] or "").strip()
        name = str(row[3] or "").strip()
        if not stage:
            # Service headings/notes are explicitly reported, rather than being silently dropped.
            warnings.append(f"Строка {row_no}: служебная строка не импортирована: {name or 'без номера этапа'}")
            continue
        if not name:
            raise ValueError(f"Строка {row_no}: не заполнено наименование этапа.")
        if all(is_blank(row[column]) for column in (4, 5, 6)):
            # Source plans often use numbered section headings (for example,
            # "1" before "1.1"). They are not planned work items and are
            # explicitly listed in preview warnings instead of treated as data.
            warnings.append(f"Строка {row_no}: групповой заголовок «{stage}» не импортирован как этап КП.")
            continue
        try:
            start = parse_date(row[4], datemode)
            end = parse_date(row[5], datemode)
            amount = parse_money(row[6])
        except ValueError as exc:
            raise ValueError(f"Строка {row_no}: {exc}") from exc
        if not start or not end:
            raise ValueError(f"Строка {row_no}: должны быть заполнены начало и окончание работ.")
        if start > end:
            raise ValueError(f"Строка {row_no}: начало работ позже окончания работ.")
        if amount is None:
            warnings.append(f"Строка {row_no}: стоимость этапа без НДС не указана; импортируется пустое значение.")
        contract = str(row[0] or "").strip()
        project = str(row[1] or "").strip()
        if not project or not contract:
            raise ValueError(f"Строка {row_no}: отсутствует реквизит проекта.")
        project_refs.add((contract, project))
        stages.append({
            "row": row_no, "contract_or_order": contract, "project": project, "stage": stage,
            "name": name, "start": start, "end": end, "amount_rub_without_vat": amount,
        })

    if not stages:
        raise ValueError("В источнике не найдено ни одного этапа КП.")
    if len(stages) < MIN_STAGES or len(stages) > MAX_STAGES:
        raise ValueError(f"Количество этапов КП должно быть от {MIN_STAGES} до {MAX_STAGES}; получено: {len(stages)}.")
    if len(project_refs) != 1:
        raise ValueError("В источнике смешаны разные проекты или реквизиты; импорт остановлен.")
    stage_keys = [stage["stage"] for stage in stages]
    duplicates = sorted({key for key in stage_keys if stage_keys.count(key) > 1})
    if duplicates:
        raise ValueError("Повторяющиеся номера этапов: " + ", ".join(duplicates) + ".")
    contract, project = next(iter(project_refs))
    return {
        "source": str(path), "header_row": header_row + 1,
        "header_mapping": dict(zip(EXPECTED, rows[header_row][:7])),
        "project": project, "contract_or_order": contract, "stages": stages,
        "stage_count": len(stages),
        "total_rub_without_vat": sum((stage["amount_rub_without_vat"] or Decimal("0") for stage in stages), Decimal("0")),
        "warnings": warnings, "duplicate_stage_numbers": [],
    }


def convert(value: object):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return format(value, ".2f")
    return value


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Предпросмотр импорта КП")
    parser.add_argument("source", type=Path)
    args = parser.parse_args()
    result = preview(args.source.resolve())
    # PowerShell 5 decodes redirected child-process output with its legacy
    # console code page. ASCII JSON escapes keep Cyrillic lossless end-to-end.
    print(json.dumps(result, ensure_ascii=True, default=convert, indent=2))
