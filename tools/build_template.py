"""Create the empty, portable GIP Navigator master workbook.

The master is deliberately data-free.  A project workbook is always created by
copying this file with New-ProjectFromTemplate.ps1; project data is never put
back into the master.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).with_name("vendor")))
import xlsxwriter
from sr_library import LIBRARY_ROWS


ROOT = Path(__file__).parents[1]
TEMPLATE_VERSION = "3.0"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output", type=Path, default=ROOT / "templates" / "GIP_Navigator_TEMPLATE.xlsx"
    )
    args = parser.parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)

    workbook = xlsxwriter.Workbook(args.output)
    workbook.set_calc_mode("auto")
    workbook.set_properties({
        "title": "GIP Navigator — универсальный шаблон",
        "subject": "Календарный план и сопутствующие работы",
        "author": "GIP Navigator",
        "company": "",
        "comments": "Обезличенный мастер-шаблон. Не содержит данных проекта.",
    })

    blue, purple, orange, cyan, green = "#125AB6", "#7B3F98", "#F5A623", "#0C9BD6", "#00B050"
    pale, yellow, grid, grey = "#F4F7FB", "#FFF2CC", "#E5E7EB", "#64748B"
    title = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "font_size": 15,
                                 "bg_color": blue, "align": "center", "valign": "vcenter"})
    header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "font_size": 10,
                                  "bg_color": purple, "align": "center", "valign": "vcenter",
                                  "text_wrap": True, "border": 1, "border_color": "#FFFFFF"})
    text = workbook.add_format({"font_name": "Aptos", "font_size": 10, "text_wrap": True,
                                "valign": "top", "border": 1, "border_color": grid})
    input_cell = workbook.add_format({"font_name": "Aptos", "font_size": 10, "text_wrap": True,
                                      "valign": "top", "bg_color": yellow, "border": 1,
                                      "border_color": grid, "locked": False})
    date_cell = workbook.add_format({"num_format": "dd.mm.yyyy", "font_name": "Aptos", "font_size": 10,
                                     "align": "center", "valign": "vcenter", "border": 1, "border_color": grid})
    date_input = workbook.add_format({"num_format": "dd.mm.yyyy", "font_name": "Aptos", "font_size": 10,
                                      "align": "center", "valign": "vcenter", "bg_color": yellow, "border": 1,
                                      "border_color": grid, "locked": False})
    money_input = workbook.add_format({"num_format": "#,##0.00", "font_name": "Aptos", "font_size": 10,
                                       "align": "center", "valign": "vcenter", "bg_color": yellow, "border": 1,
                                       "border_color": grid, "locked": False})
    money = workbook.add_format({"num_format": "#,##0.00", "font_name": "Aptos", "font_size": 10,
                                 "align": "center", "valign": "vcenter", "border": 1, "border_color": grid})
    note = workbook.add_format({"italic": True, "font_color": grey, "bg_color": pale,
                                "text_wrap": True, "valign": "vcenter"})
    service = workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": grey,
                                   "bg_color": pale, "border": 1, "border_color": grid})
    checkbox_cell = workbook.add_format({"font_name": "Aptos", "font_size": 10, "align": "center",
                                         "valign": "vcenter", "num_format": ";;;", "border": 1,
                                         "border_color": grid})
    card_formats = {color: workbook.add_format({"bold": True, "font_color": "#FFFFFF", "font_size": 12,
                                                 "bg_color": color, "align": "center", "valign": "vcenter",
                                                 "text_wrap": True})
                    for color in (purple, cyan, orange, green)}
    card_date = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "font_size": 12,
                                     "bg_color": purple, "align": "center", "valign": "vcenter",
                                     "num_format": "dd.mm.yyyy"})

    sheets = [
        ("00_Главная", blue), ("01_КП", purple), ("02_Библиотека_СР", cyan),
        ("03_СР_проекта", orange), ("04_Согласование_плана", cyan),
        ("05_График", purple), ("06_Контроль", orange), ("99_Служебный", grey),
        ("07_Актирование", green),
    ]
    ws = {name: workbook.add_worksheet(name) for name, _ in sheets}
    for name, color in sheets:
        ws[name].set_tab_color(color)
        ws[name].hide_gridlines(2)

    # 00_Главная
    home = ws["00_Главная"]
    home.set_column("A:P", 13)
    home.merge_range("A1:P2", "GIP Navigator — универсальный шаблон", title)
    home.set_row(0, 27); home.set_row(1, 27)
    card_headers = [("A4:D4", "ПРОЕКТ", purple), ("E4:H4", "ЭТАПОВ КП", cyan),
                    ("I4:L4", "СТОИМОСТЬ КП", orange), ("M4:P4", "ЗААКТИРОВАНО", green)]
    for cell_range, label, color in card_headers:
        home.merge_range(cell_range, label, card_formats[color])
    for cell_range, color in (("A5:D6", purple), ("E5:H6", cyan), ("I5:L6", orange), ("M5:P6", green)):
        home.merge_range(cell_range, "", card_formats[color])
    home.write_formula("A5", "=IF('01_КП'!B5=\"\",\"не импортирован\",'01_КП'!B5)", card_formats[purple], "не импортирован")
    home.write_formula("E5", "=COUNTA('01_КП'!$C$5:$C$79)", card_formats[cyan], 0)
    home.write_formula("I5", "=SUM('01_КП'!$G$5:$G$79)", card_formats[orange], 0)
    home.write_formula("M5", "=SUM('01_КП'!$I$5:$I$79)", card_formats[green], 0)
    home.merge_range("A8:D8", "НАЧАЛО ПРОЕКТА\n(с учётом НЗ)", card_formats[purple])
    home.merge_range("E8:H8", "ОСТАТОК ПО КП", card_formats[cyan])
    home.merge_range("I8:L8", "КРИТИЧЕСКИХ РИСКОВ", card_formats[orange])
    home.merge_range("M8:P8", "НЕДОСТАТОЧНО ДАННЫХ", card_formats[green])
    for cell_range, color in (("A9:D9", purple), ("E9:H9", cyan), ("I9:L9", orange), ("M9:P9", green)):
        home.merge_range(cell_range, "", card_formats[color])
    home.write_formula("A9", "=IF(COUNT('01_КП'!$J$5:$J$79)=0,\"не импортирован\",MIN('01_КП'!$J$5:$J$79))", card_date, "не импортирован")
    home.write_formula("E9", "=SUM('01_КП'!$L$5:$L$79)", card_formats[cyan], 0)
    home.write_formula("I9", "=COUNTIF('06_Контроль'!$C$5:$C$304,\"Критично\")", card_formats[orange], 0)
    home.write_formula("M9", "=COUNTIF('06_Контроль'!$C$5:$C$304,\"Недостаточно данных\")", card_formats[green], 0)
    home.merge_range("A15:P17", "1. Скопируйте этот файл в папку проекта и переименуйте его.\n"
                                  "2. Импортируйте исходный КП после предпросмотра.\n"
                                  "3. Выберите сопутствующие работы, задайте связи на листе 04 и пересчитайте внутренний график.", note)
    home.set_row(14, 42)

    # 01_КП: 75 reserved logical rows, not a filled demonstration plan.
    kp = ws["01_КП"]
    kp.set_column("A:A", 31); kp.set_column("B:B", 42); kp.set_column("C:C", 13)
    kp.set_column("D:D", 56); kp.set_column("E:F", 18); kp.set_column("G:G", 23)
    kp.set_column("H:H", 24); kp.set_column("I:I", 22); kp.set_column("J:K", 18)
    kp.set_column("L:L", 22); kp.set_column("M:M", 16); kp.set_column("N:O", 18)
    kp.set_column("P:P", 15); kp.set_column("Q:Q", 22); kp.set_column("R:V", 18)
    kp.merge_range("A1:V1", "Утверждённый календарный план", title)
    kp.merge_range("A2:V2", "Колонки A:G заполняются только импортом исходного КП. H:I сохраняют прежнее фактическое актирование. Смещённые даты, прогноз этапов, риски и суммы подписанных актов рассчитываются автоматически.", note)
    kp.write("A3", "Дата подписания наряд-заказа", header)
    kp.write_blank("B3", None, date_input)
    kp.write("D3", "Смещение сроков, кал. дн.", header)
    kp.write_formula("E3", "=IF(COUNT($E$5:$E$79)=0,0,MAX(0,IF(ISNUMBER($B$3),$B$3-MIN($E$5:$E$79),0)))", date_cell, 0)
    kp.write("G3", "Статус сроков", header)
    kp.merge_range("H3:M3", "", note)
    kp.write_formula("H3", "=IF(COUNTA($C$5:$C$79)=0,\"ожидается импорт КП\",IF($B$3=\"\",\"Не введена дата подписания НЗ\",IF($E$3=0,\"Сроки КП без смещения\",\"Сроки КП смещены на \"&$E$3&\" кал. дн.\")))", note, "ожидается импорт КП")
    kp.set_row(2, 26)
    kp_headers = ["Номер договора / Номер Наряд-Заказа", "Наименование проекта", "№ Этапа",
                  "Наименование этапа", "Начало по утверждённому КП", "Окончание по утверждённому КП",
                  "Стоимость этапа, руб. без НДС", "Дата фактического актирования",
                  "Фактически заактировано, руб. без НДС", "Смещённое начало", "Смещённое окончание",
                  "Остаток, руб. без НДС", "% актирования", "Прогноз начала этапа", "Прогноз окончания этапа", "Отклонение, кал. дн.", "Статус срока", "Под риском, руб.", "Подписанные акты, руб.", "Остаток по актам, руб.", "% актирования по актам", "Причина прогноза"]
    kp.write_row("A4", kp_headers, header); kp.set_row(3, 34)
    for row in range(4, 79):
        for col in range(22):
            fmt = date_input if col == 7 else money_input if col == 8 else money if col in (6, 11) else text
            if col in (4, 5, 9, 10, 13, 14): fmt = date_cell
            kp.write_blank(row, col, None, fmt)
        excel_row = row + 1
        kp.write_formula(row, 9, f'=IF(C{excel_row}="","",E{excel_row}+$E$3)', date_cell)
        kp.write_formula(row, 10, f'=IF(C{excel_row}="","",F{excel_row}+$E$3)', date_cell)
        kp.write_formula(row, 11, f'=IF(C{excel_row}="","",IF(G{excel_row}="","",G{excel_row}-IF(I{excel_row}="",0,I{excel_row})))', money)
        kp.write_formula(row, 12, f'=IF(OR(C{excel_row}="",I{excel_row}="",G{excel_row}="",G{excel_row}=0),"",I{excel_row}/G{excel_row})', workbook.add_format({"num_format": "0.0%", "font_name": "Aptos", "font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": grid}))
        # S is written by RecalculateActs from the act register.  Until the
        # first signed act, T:U retain the legacy H:I stage summary.
        kp.write_formula(row, 19, f'=IF(C{excel_row}="","",IF(G{excel_row}="","",G{excel_row}-IF(S{excel_row}="",IF(I{excel_row}="",0,I{excel_row}),S{excel_row})))', money)
        kp.write_formula(row, 20, f'=IF(OR(C{excel_row}="",G{excel_row}="",G{excel_row}=0),"",IF(S{excel_row}="",M{excel_row},S{excel_row}/G{excel_row}))', workbook.add_format({"num_format": "0.0%", "font_name": "Aptos", "font_size": 10, "align": "center", "valign": "vcenter", "border": 1, "border_color": grid}))
        kp.set_row(row, 24)
    kp.conditional_format("I5:L79", {"type": "formula", "criteria": "=$L5<0", "format": workbook.add_format({"bg_color": "#FCE4D6", "font_color": "#C00000"})})
    kp.autofilter(3, 0, 78, 21); kp.freeze_panes(4, 0)
    # Contractual dates (E:F) are locked.  The unprotected yellow inputs are
    # the signing date and the factual act data.  Import macros temporarily
    # unprotect the sheet and restore protection in their safe-exit path.
    kp.protect("", {"autofilter": True})

    # 02_Библиотека_СР: a new project starts with the whole standard library
    # selected.  The linked Excel controls created by build_xlsm.ps1 preserve
    # this state when the workbook is opened again.
    library = ws["02_Библиотека_СР"]
    for col, width in enumerate((12, 13, 44, 20, 34, 16)):
        library.set_column(col, col, width)
    library.set_column("H:L", 12)
    library.set_column("M:P", 2, None, {"hidden": True})
    library.merge_range("A1:P1", "Библиотека типовых сопутствующих работ", title)
    library.merge_range("A2:P2", "Все 81 типовые СР включены для нового проекта. Снимите флажок у неприменимой работы. Диаграмма справа показывает рассчитанные сроки выбранных СР.", note)
    library_headers = ["В проект", "Код СР", "Наименование", "Группа", "Ожидаемый результат", "Длительность, кал. дн."]
    library.write_row("A4", library_headers, header); library.set_row(3, 34)
    for row, item in enumerate(LIBRARY_ROWS, 4):
        values = [item["code"], item["name"], item["group"], item["result"], item["duration"]]
        library.write(row, 0, True, checkbox_cell)
        for col, value in enumerate(values, 1):
            library.write_blank(row, col, None, text) if value is None else library.write(row, col, value, text)
        library.set_row(row, 30)
    library.autofilter(3, 0, 84, 5); library.freeze_panes(4, 0)

    # 03_СР_проекта: clear working view for the project manager.  IDs and
    # network diagnostics stay technical and are surfaced on 06_Контроль.
    sr = ws["03_СР_проекта"]
    widths = (2, 13, 38, 17, 18, 18, 16, 16, 18, 18, 23, 28, 2, 2, 2)
    for col, width in enumerate(widths): sr.set_column(col, col, width)
    sr.set_column("A:A", 2, None, {"hidden": True})
    sr.set_column("M:O", 2, None, {"hidden": True})
    sr.merge_range("A1:O1", "План сопутствующих работ проекта", title)
    sr.merge_range("A2:O2", "Все сроки — календарные дни. Жёлтые поля вводит ГИП: факт и ожидаемое окончание. Плановые даты и статус рассчитываются. Связи FS/SS и ошибки сети — на листе 06 «Контроль».", note)
    sr_headers = ["ID экземпляра", "Код СР", "Наименование", "Расчётное плановое начало", "Длительность, кал. дн.",
                  "Расчётное окончание", "Факт. начало", "Факт. окончание", "Ожид. окончание",
                  "Прогноз окончания", "Статус СР", "Комментарий", "Проверка сети", "Источник", "Ограничение начала (не ранее)"]
    sr.write_row("A4", sr_headers, header); sr.set_row(3, 34)
    for row in range(4, 304):
        for col in range(15):
            fmt = date_input if col in (6, 7, 8, 14) else date_cell if col in (3, 5, 9) else input_cell
            sr.write_blank(row, col, None, fmt)
        excel_row = row + 1
        sr.write_formula(row, 2, f'=IF(B{excel_row}="","",IFERROR(VLOOKUP(B{excel_row},\'02_Библиотека_СР\'!$B$5:$C$85,2,FALSE),""))', input_cell)
        # A blank library duration is a deliberate "to be clarified" value,
        # not a zero-day milestone.  COUNTIFS keeps blank source cells blank
        # in the project sheet, allowing the network check to flag them.
        sr.write_formula(row, 4, f'=IF(B{excel_row}="","",IF(COUNTIF(\'02_Библиотека_СР\'!$B$5:$B$85,B{excel_row})=0,"",IF(COUNTIFS(\'02_Библиотека_СР\'!$B$5:$B$85,B{excel_row},\'02_Библиотека_СР\'!$F$5:$F$85,"<>")=0,"",VLOOKUP(B{excel_row},\'02_Библиотека_СР\'!$B$5:$F$85,5,FALSE))))', input_cell)
        sr.set_row(row, 23)
    sr.data_validation("B5:B304", {"validate": "list", "source": "='02_Библиотека_СР'!$B$5:$B$85", "error_type": "warning"})
    sr.data_validation("E5:E304", {"validate": "integer", "criteria": ">=", "value": 0, "error_type": "stop"})
    sr.autofilter(3, 1, 303, 11); sr.freeze_panes(4, 1)

    agreement = ws["04_Согласование_плана"]
    agreement.set_column("A:A", 16); agreement.set_column("B:C", 25); agreement.set_column("D:D", 14)
    agreement.set_column("E:E", 12); agreement.set_column("F:F", 18); agreement.set_column("G:G", 42)
    agreement.set_column("I:I", 27); agreement.set_column("J:J", 38); agreement.set_column("K:K", 13); agreement.set_column("L:L", 42); agreement.set_column("M:M", 25); agreement.set_column("N:N", 22); agreement.set_column("O:O", 48)
    agreement.set_column("Q:Q", 28); agreement.set_column("R:R", 24); agreement.set_column("S:S", 28); agreement.set_column("T:V", 18)
    agreement.merge_range("A1:G1", "Связи СР → СР", title)
    agreement.merge_range("A2:G2", "Каждая строка задаёт связь между СР. FS: старт после окончания; SS: старт от начала. Лаг считается в календарных днях.", note)
    agreement.write_row("A4", ["ID связи", "Последующая СР (ID)", "Предшествующая СР (ID)", "Тип связи", "Лаг, дн.", "Проверка", "Комментарий"], header)
    for row in range(4, 304):
        for col in range(7): agreement.write_blank(row, col, None, input_cell)
    agreement.data_validation("B5:C304", {"validate": "list", "source": "='03_СР_проекта'!$A$5:$A$304", "error_type": "warning"})
    agreement.data_validation("D5:D304", {"validate": "list", "source": ["FS", "SS"], "error_type": "warning"})
    agreement.autofilter(3, 0, 303, 6); agreement.freeze_panes(4, 0)
    agreement.merge_range("I1:O1", "Обязательные связи СР → этап КП", title)
    agreement.merge_range("I2:O2", "Выберите ID СР: рядом отображается название. Условие задаёт крайний срок выбранной СР относительно смещённого этапа КП.", note)
    agreement.write_row("I4", ["ID СР", "Наименование СР", "№ этапа КП", "Наименование этапа КП", "Условие", "Проверка", "Причина / критическая СР"], header)
    for row in range(4, 304):
        for col in range(8, 15): agreement.write_blank(row, col, None, input_cell)
        excel_row = row + 1
        agreement.write_formula(row, 9, f'=IF(I{excel_row}="","",IFERROR(VLOOKUP(I{excel_row},\'03_СР_проекта\'!$A$5:$C$304,3,FALSE),""))', text)
        agreement.write_formula(row, 11, f'=IF(K{excel_row}="","",IFERROR(VLOOKUP(K{excel_row},\'01_КП\'!$C$5:$D$79,2,FALSE),""))', text)
    agreement.data_validation("I5:I304", {"validate": "list", "source": "='03_СР_проекта'!$A$5:$A$304", "error_type": "warning"})
    agreement.data_validation("K5:K304", {"validate": "list", "source": "='01_КП'!$C$5:$C$79", "error_type": "warning"})
    agreement.data_validation("M5:M304", {"validate": "list", "source": ["До начала этапа", "До окончания этапа"], "error_type": "warning"})
    agreement.merge_range("Q1:V1", "Согласование первоначального графика", title)
    agreement.write("Q3", "Версия", header); agreement.write_blank("R3", None, input_cell)
    agreement.write("Q4", "Дата согласования", header); agreement.write_blank("R4", None, date_input)
    agreement.write("Q5", "Ответственный", header); agreement.write_blank("R5", None, input_cell)
    agreement.write("Q6", "Статус", header); agreement.merge_range("R6:V6", "Заполните версию, дату и ответственного; затем нажмите кнопку на главной.", note)
    agreement.write_row("Q9", ["Версия", "Дата", "№ этапа", "Наименование этапа", "Утверждённое начало", "Утверждённое окончание"], header)
    for row in range(9, 304):
        for col in range(16, 22): agreement.write_blank(row, col, None, date_cell if col in (20, 21) else text)

    graph = ws["05_График"]
    graph.set_column("A:A", 18); graph.set_column("B:B", 58); graph.set_column("C:I", 16); graph.set_column("K:Z", 12)
    graph.merge_range("A1:I1", "Диаграмма сроков КП и СР", title)
    graph.merge_range("A2:I2", "Показываются смещённые плановые сроки, прогнозы обязательных СР и фактическое выполнение. Пустые строки скрываются автоматически.", note)
    graph.write_row("A4", ["Тип", "Наименование", "Начало", "Длительность", "КП — план", "КП — прогноз", "СР — план", "СР — прогноз", "СР — факт"], header)
    for row in range(4, 79):
        excel_row = row + 1
        graph.write_formula(row, 0, f"=IF('01_КП'!$C{excel_row}=\"\",\"\",\"КП\")", text)
        graph.write_formula(row, 1, f"=IF($A{excel_row}=\"\",\"\",'01_КП'!$D{excel_row})", text)
        graph.write_formula(row, 2, f"=IF($A{excel_row}=\"\",\"\",'01_КП'!$J{excel_row})", date_cell)
        graph.write_formula(row, 3, f"=IF($A{excel_row}=\"\",\"\",'01_КП'!$K{excel_row}-'01_КП'!$J{excel_row}+1)", text)
        # Empty reserved rows must stay empty in the visible source table.  The
        # chart ignores blank values, whereas NA() leaves distracting #N/A cells
        # for work that has not yet been selected.
        graph.write_formula(row, 4, f"=IF($A{excel_row}=\"\",\"\",$D{excel_row})", text)
        graph.write_formula(row, 5, "=\"\"", text)
    for row in range(79, 379):
        for col in range(9): graph.write_blank(row, col, None, text)
        graph.set_row(row, 23, None, {"hidden": True})
    workbook.define_name("GIP_Gantt_Task_Names", "=OFFSET('05_График'!$B$5,0,0,COUNTA('01_КП'!$C$5:$C$79),1)")
    workbook.define_name("GIP_Gantt_Starts", "=OFFSET('05_График'!$C$5,0,0,COUNTA('01_КП'!$C$5:$C$79),1)")
    workbook.define_name("GIP_Gantt_Durations", "=OFFSET('05_График'!$E$5,0,0,COUNTA('01_КП'!$C$5:$C$79),1)")
    chart = workbook.add_chart({"type": "bar", "subtype": "stacked"})
    chart.add_series({"categories": "=GIP_Gantt_Task_Names", "values": "=GIP_Gantt_Starts",
                      "fill": {"none": True}, "border": {"none": True}})
    chart.add_series({"name": "КП — договорные сроки",
                      "categories": "=GIP_Gantt_Task_Names", "values": "=GIP_Gantt_Durations",
                      "fill": {"color": purple}, "border": {"color": blue}})
    chart.set_title({"name": "Диаграмма Ганта: этапы КП"})
    chart.set_legend({"position": "bottom"}); chart.set_y_axis({"reverse": True})
    chart.set_size({"width": 1050, "height": 600})
    graph.insert_chart("H4", chart)
    graph.freeze_panes(4, 0)

    control = ws["06_Контроль"]
    control.set_column("A:A", 18); control.set_column("B:B", 15); control.set_column("C:C", 22); control.set_column("D:D", 38); control.set_column("E:F", 18); control.set_column("G:G", 14); control.set_column("H:H", 48)
    control.merge_range("A1:H1", "Контроль сроков и данных", title)
    control.merge_range("A2:H2", "Сигналы формируются из единого расчёта КП и СР. Пустой прогноз никогда не считается нормальным. Порог предупреждения: 7 календарных дней.", note)
    control.write_row("A4", ["Категория", "№ этапа", "Состояние", "СР / причина", "Смещённый дедлайн", "Прогноз", "Отклонение, дн.", "Рекомендация"], header)
    control.freeze_panes(4, 0)

    service_ws = ws["99_Служебный"]
    service_ws.set_column("A:A", 28); service_ws.set_column("B:B", 42)
    service_ws.merge_range("A1:B1", "Служебные параметры книги", title)
    service_ws.write_row("A3", ["Параметр", "Значение"], header)
    service_rows = [
        ("template_version", TEMPLATE_VERSION),
        ("workbook_id", ""),
        ("project_created_at", ""),
        ("kp_imported_at", ""),
        ("kp_baseline_revision", ""),
        ("schema_version", "2"),
        ("schedule_rule", "Срок включителен; FS = следующий день после окончания предшественника; SS = дата начала предшественника; лаг в календаре последующей СР."),
    ]
    for row, values in enumerate(service_rows, 3): service_ws.write_row(row, 0, values, service)
    service_ws.hide()

    acts = ws["07_Актирование"]
    acts.set_column("A:A", 18); acts.set_column("B:B", 13); acts.set_column("C:C", 44); acts.set_column("D:D", 16); acts.set_column("E:E", 22); acts.set_column("F:F", 18); acts.set_column("G:G", 16); acts.set_column("H:H", 38)
    acts.merge_range("A1:H1", "Реестр актирования этапов КП", title)
    acts.merge_range("A2:H2", "Одна строка — один акт. Черновики не входят в сумму актирования. Подписанные частичные акты суммируются по этапу; оплаты в книге не ведутся.", note)
    acts.write_row("A4", ["ID акта", "№ этапа КП", "Наименование этапа", "Дата акта", "Номер акта", "Сумма, руб. без НДС", "Статус", "Проверка"], header)
    for row in range(4, 304):
        for col in range(8): acts.write_blank(row, col, None, date_input if col == 3 else money_input if col == 5 else input_cell)
        excel_row = row + 1
        acts.write_formula(row, 2, f'=IF(B{excel_row}="","",IFERROR(VLOOKUP(B{excel_row},\'01_КП\'!$C$5:$D$79,2,FALSE),""))', text)
    acts.data_validation("B5:B304", {"validate": "list", "source": "='01_КП'!$C$5:$C$79", "error_type": "warning"})
    acts.data_validation("G5:G304", {"validate": "list", "source": ["Черновик", "Подписан"], "error_type": "warning"})
    acts.autofilter(3, 0, 303, 7); acts.freeze_panes(4, 0)

    workbook.close()
    print(args.output)


if __name__ == "__main__":
    main()
