"""Stage 1: safe preview of the fixed seven-column КП import."""
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

def rows_from_book(path: Path):
    if path.suffix.casefold() == ".xls":
        book = xlrd.open_workbook(path)
        sheet = book.sheet_by_index(0)
        rows = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        return rows, book.datemode
    if path.suffix.casefold() == ".xlsx":
        book = openpyxl.load_workbook(path, read_only=True, data_only=True)
        sheet = book.active
        return [list(row) for row in sheet.iter_rows(values_only=True)], None
    raise ValueError("Поддерживаются только .xls и .xlsx")

def parse_date(value, datemode):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, float) and datemode is not None: return xlrd.xldate_as_datetime(value, datemode).date()
    if value in (None, ""): return None
    raise ValueError(f"некорректная дата: {value!r}")

def parse_money(value) -> Decimal | None:
    if value in (None, ""): return None
    try: return Decimal(str(value).replace(" ", "").replace(",", ".")).quantize(Decimal("0.01"))
    except InvalidOperation as exc: raise ValueError(f"некорректная стоимость: {value!r}") from exc

def preview(path: Path) -> dict:
    rows, datemode = rows_from_book(path)
    if not rows: raise ValueError("Книга пуста")
    header = [norm(cell) for cell in rows[0][:7]]
    if len(header) != 7 or any(header[i] not in ALIASES[i] for i in range(7)):
        raise ValueError("Шапка не соответствует семи колонкам КП; импорт не применён.")
    stages, warnings, project_refs = [], [], set()
    for row_no, raw in enumerate(rows[1:], 2):
        row = list(raw[:7]) + [None] * max(0, 7 - len(raw))
        stage = str(row[2] or "").strip()
        title = str(row[3] or "").strip()
        if not stage:
            if title: warnings.append(f"Строка {row_no}: служебная строка пропущена: {title}")
            continue
        item = {"row": row_no, "contract_or_order": str(row[0] or "").strip(), "project": str(row[1] or "").strip(), "stage": stage, "name": title, "start": parse_date(row[4], datemode), "end": parse_date(row[5], datemode), "amount_rub_without_vat": parse_money(row[6])}
        if not item["project"] or not item["contract_or_order"]: raise ValueError(f"Строка {row_no}: отсутствует реквизит проекта")
        project_refs.add((item["contract_or_order"], item["project"]))
        stages.append(item)
    if len(project_refs) != 1: raise ValueError("В источнике смешаны разные проекты или реквизиты; импорт остановлен.")
    duplicates = sorted({x["stage"] for x in stages if sum(y["stage"] == x["stage"] for y in stages) > 1})
    return {"source": str(path), "header_mapping": dict(zip(EXPECTED, rows[0][:7])), "project": next(iter(project_refs))[1], "contract_or_order": next(iter(project_refs))[0], "stages": stages, "total_rub_without_vat": sum((x["amount_rub_without_vat"] or Decimal("0")) for x in stages), "warnings": warnings, "duplicate_stage_numbers": duplicates}

def convert(value):
    if isinstance(value, (date, datetime)): return value.isoformat()
    if isinstance(value, Decimal): return format(value, ".2f")
    return value

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Предпросмотр импорта КП")
    parser.add_argument("source", type=Path)
    args = parser.parse_args()
    result = preview(args.source.resolve())
    # PowerShell 5 decodes redirected child-process output with the legacy
    # console code page. ASCII JSON escapes keep Cyrillic lossless end-to-end.
    print(json.dumps(result, ensure_ascii=True, default=convert, indent=2))
