"""Regression tests for the 2–75 stage КП preview contract."""
from __future__ import annotations

import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tools" / "vendor"))
import xlsxwriter
from gipnav_import import EXPECTED, preview


def make_source(path: Path, stage_count: int, *, bad_date: bool = False, duplicate: bool = False) -> None:
    book = xlsxwriter.Workbook(path)
    sheet = book.add_worksheet("КП")
    date_format = book.add_format({"num_format": "dd.mm.yyyy"})
    sheet.write(0, 0, "Примечание перед шапкой: разрешено и явно обработано")
    sheet.write_row(1, 0, EXPECTED)
    start = date(2026, 1, 1)
    for index in range(stage_count):
        stage = "A-1" if duplicate and index == stage_count - 1 else ("A-1" if index == 0 else f"{index + 1}.этап")
        row_start = start + timedelta(days=index * 3)
        row_end = row_start - timedelta(days=1) if bad_date and index == 0 else row_start + timedelta(days=2)
        sheet.write_row(index + 2, 0, ["Договор-ТЕСТ", "Тестовый проект", stage, f"Работа {index + 1}",
                                      None, None, (index + 1) * 1000])
        sheet.write_datetime(index + 2, 4, row_start, date_format)
        sheet.write_datetime(index + 2, 5, row_end, date_format)
    book.close()


def expect_error(path: Path, phrase: str) -> None:
    try:
        preview(path)
    except ValueError as exc:
        if phrase not in str(exc):
            raise AssertionError(f"Expected {phrase!r}, got {exc!s}") from exc
    else:
        raise AssertionError("Invalid КП was accepted")


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="gip-kp-import-") as temp:
        directory = Path(temp)
        for count in (2, 10, 75):
            path = directory / f"kp_{count}.xlsx"
            make_source(path, count)
            result = preview(path)
            assert result["stage_count"] == count
            assert len(result["stages"]) == count
            assert result["stages"][-1]["stage"] == ("A-1" if count == 1 else f"{count}.этап")
            assert result["total_rub_without_vat"] == sum(range(1, count + 1)) * 1000
        too_many = directory / "kp_76.xlsx"; make_source(too_many, 76); expect_error(too_many, "от 2 до 75")
        bad_date = directory / "bad_date.xlsx"; make_source(bad_date, 2, bad_date=True); expect_error(bad_date, "начало работ позже")
        duplicate = directory / "duplicate.xlsx"; make_source(duplicate, 2, duplicate=True); expect_error(duplicate, "Повторяющиеся номера")
    print("PASS: КП валидатор проверил 2, 10, 75 этапов и отклонил 76, дубль и неверные даты.")


if __name__ == "__main__":
    main()
