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
TEMPLATE_VERSION = "1.0"


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
                                      "border_color": grid})
    date_cell = workbook.add_format({"num_format": "dd.mm.yyyy", "font_name": "Aptos", "font_size": 10,
                                     "border": 1, "border_color": grid})
    date_input = workbook.add_format({"num_format": "dd.mm.yyyy", "font_name": "Aptos", "font_size": 10,
                                      "bg_color": yellow, "border": 1, "border_color": grid})
    money = workbook.add_format({"num_format": "#,##0.00", "font_name": "Aptos", "font_size": 10,
                                 "border": 1, "border_color": grid})
    note = workbook.add_format({"italic": True, "font_color": grey, "bg_color": pale,
                                "text_wrap": True, "valign": "vcenter"})
    service = workbook.add_format({"font_name": "Aptos", "font_size": 10, "font_color": grey,
                                   "bg_color": pale, "border": 1, "border_color": grid})
    card_formats = {color: workbook.add_format({"bold": True, "font_color": "#FFFFFF", "font_size": 12,
                                                 "bg_color": color, "align": "center", "valign": "vcenter",
                                                 "text_wrap": True})
                    for color in (purple, cyan, orange, green)}

    sheets = [
        ("00_Главная", blue), ("01_КП", purple), ("02_Библиотека_СР", cyan),
        ("03_СР_проекта", orange), ("04_Согласование_плана", cyan),
        ("05_График", purple), ("06_Контроль", orange), ("07_Акты", green),
        ("08_Внешние_согласования", cyan), ("99_Служебный", grey),
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
                    ("I4:L4", "СТОИМОСТЬ КП", orange), ("M4:P4", "СОСТОЯНИЕ", green)]
    for cell_range, label, color in card_headers:
        home.merge_range(cell_range, label, card_formats[color])
    for cell_range, color in (("A5:D6", purple), ("E5:H6", cyan), ("I5:L6", orange), ("M5:P6", green)):
        home.merge_range(cell_range, "", card_formats[color])
    home.write_formula("A5", "=IF('01_КП'!B5=\"\",\"не импортирован\",'01_КП'!B5)", card_formats[purple], "не импортирован")
    home.write_formula("E5", "=COUNTA('01_КП'!$C$5:$C$79)", card_formats[cyan], 0)
    home.write_formula("I5", "=SUM('01_КП'!$G$5:$G$79)", card_formats[orange], 0)
    home.write_formula("M5", "=IF(COUNTA('01_КП'!$C$5:$C$79)=0,\"ожидается импорт КП\",\"КП импортирован\")", card_formats[green], "ожидается импорт КП")
    home.merge_range("A8:P10", "1. Создайте отдельную книгу проекта из этого шаблона.\n"
                                  "2. Импортируйте исходный КП после предпросмотра.\n"
                                  "3. Затем выберите сопутствующие работы и согласуйте внутренний график.", note)
    home.set_row(7, 42)

    # 01_КП: 75 reserved logical rows, not a filled demonstration plan.
    kp = ws["01_КП"]
    kp.set_column("A:A", 31); kp.set_column("B:B", 42); kp.set_column("C:C", 13)
    kp.set_column("D:D", 56); kp.set_column("E:F", 15); kp.set_column("G:G", 23)
    kp.merge_range("A1:G1", "Утверждённый календарный план", title)
    kp.merge_range("A2:G2", "Эти семь колонок заполняются только импортом из исходного КП. "
                             "Мастер не содержит этапов проекта.", note)
    kp_headers = ["Номер договора / Номер Наряд-Заказа", "Наименование проекта", "№ Этапа",
                  "Наименование этапа", "Начало работ", "Окончание работ",
                  "Стоимость этапа, руб. без НДС"]
    kp.write_row("A4", kp_headers, header); kp.set_row(3, 34)
    for row in range(4, 79):
        for col in range(7):
            fmt = date_cell if col in (4, 5) else money if col == 6 else text
            kp.write_blank(row, col, None, fmt)
        kp.set_row(row, 24)
    kp.autofilter(3, 0, 78, 6); kp.freeze_panes(4, 0)

    # 02_Библиотека_СР: standard library. It contains no project-specific selection.
    library = ws["02_Библиотека_СР"]
    for col, width in enumerate((13, 44, 20, 34, 16, 15, 24, 12, 14, 20, 36)):
        library.set_column(col, col, width)
    library.merge_range("A1:K1", "Библиотека типовых сопутствующих работ", title)
    library.merge_range("A2:K2", "Начальная библиотека из 81 типовой СР. Сроки можно менять только в книге проекта; «уточнить» не является утверждённым нормативом.", note)
    library_headers = ["Код СР", "Наименование", "Группа", "Ожидаемый результат", "Длительность",
                       "Тип дней", "Применимость", "ИД", "Источник срока", "Типовая связь", "Примечание"]
    library.write_row("A4", library_headers, header); library.set_row(3, 34)
    for row, item in enumerate(LIBRARY_ROWS, 4):
        values = [item["code"], item["name"], item["group"], item["result"], item["duration"], item["day_type"],
                  item["applicability"], item["needs_id"], item["source"], item["dependency"], item["note"]]
        for col, value in enumerate(values):
            library.write_blank(row, col, None, text) if value is None else library.write(row, col, value, text)
        library.set_row(row, 30)
    library.autofilter(3, 0, 84, 10); library.freeze_panes(4, 0)

    # 03_СР_проекта: space for a later, per-project working schedule (300 records).
    sr = ws["03_СР_проекта"]
    widths = (20, 13, 38, 16, 16, 13, 15, 16, 16, 16, 16, 16, 17, 28)
    for col, width in enumerate(widths): sr.set_column(col, col, width)
    sr.merge_range("A1:N1", "План сопутствующих работ проекта", title)
    sr.merge_range("A2:N2", "Жёлтые ячейки предназначены для ввода после выбора СР. В шаблоне нет демонстрационных работ.", note)
    sr_headers = ["ID экземпляра", "Код СР", "Наименование", "Этап КП", "План начала", "Длительность",
                  "Тип дней", "Расчётное окончание", "Факт. начало", "Факт. окончание", "Ожид. окончание",
                  "Прогноз окончания", "Статус ИД", "Комментарий"]
    sr.write_row("A4", sr_headers, header); sr.set_row(3, 34)
    for row in range(4, 304):
        for col in range(14):
            fmt = date_input if col in (4, 8, 9, 10) else date_cell if col in (7, 11) else input_cell
            sr.write_blank(row, col, None, fmt)
        excel_row = row + 1
        sr.write_formula(row, 2, f'=IF(B{excel_row}="","",IFERROR(VLOOKUP(B{excel_row},\'02_Библиотека_СР\'!$A$5:$B$85,2,FALSE),""))', input_cell)
        sr.write_formula(row, 5, f'=IF(B{excel_row}="","",IFERROR(VLOOKUP(B{excel_row},\'02_Библиотека_СР\'!$A$5:$F$85,5,FALSE),""))', input_cell)
        sr.write_formula(row, 6, f'=IF(B{excel_row}="","",IFERROR(VLOOKUP(B{excel_row},\'02_Библиотека_СР\'!$A$5:$F$85,6,FALSE),""))', input_cell)
        sr.set_row(row, 23)
    kp_stage_list = "=OFFSET('01_КП'!$C$5,0,0,MAX(1,COUNTA('01_КП'!$C$5:$C$79)),1)"
    sr.data_validation("D5:D304", {"validate": "list", "source": kp_stage_list})
    sr.data_validation("B5:B304", {"validate": "list", "source": "='02_Библиотека_СР'!$A$5:$A$85", "error_type": "warning"})
    sr.autofilter(3, 0, 303, 13); sr.freeze_panes(4, 0)

    agreement = ws["04_Согласование_плана"]
    agreement.set_column("A:A", 16); agreement.set_column("B:B", 18); agreement.set_column("C:C", 17)
    agreement.set_column("D:D", 19); agreement.set_column("E:E", 25); agreement.set_column("F:F", 42)
    agreement.merge_range("A1:F1", "Согласование внутреннего графика", title)
    agreement.merge_range("A2:F2", "Согласованная версия фиксируется отдельно от текущего прогноза. "
                                   "Последующий импорт КП не должен менять её без новой ревизии.", note)
    agreement.write_row("A4", ["Номер ревизии", "Дата фиксации", "Согласовано", "Основание",
                                "Статус", "Комментарий"], header)
    for row in range(4, 24):
        for col in range(6): agreement.write_blank(row, col, None, date_input if col == 1 else input_cell)
    agreement.freeze_panes(4, 0)

    graph = ws["05_График"]
    graph.set_column("A:A", 16); graph.set_column("B:B", 50); graph.set_column("C:F", 16)
    graph.merge_range("A1:F1", "Диаграмма сроков КП и СР", title)
    graph.merge_range("A2:F2", "На этом шаге диаграмма показывает только реальные этапы КП. Согласованный план, факт и прогноз СР будут добавлены после настройки СР.", note)
    graph.write_row("A4", ["Тип", "Наименование", "Начало", "Длительность", "КП — срок", "СР — срок"], header)
    for row in range(4, 79):
        excel_row = row + 1
        graph.write_formula(row, 0, f"=IF('01_КП'!$C{excel_row}=\"\",\"\",\"КП\")", text)
        graph.write_formula(row, 1, f"=IF($A{excel_row}=\"\",\"\",'01_КП'!$D{excel_row})", text)
        graph.write_formula(row, 2, f"=IF($A{excel_row}=\"\",\"\",'01_КП'!$E{excel_row})", date_cell)
        graph.write_formula(row, 3, f"=IF($A{excel_row}=\"\",\"\",'01_КП'!$F{excel_row}-'01_КП'!$E{excel_row}+1)", text)
        graph.write_formula(row, 4, f"=IF($A{excel_row}=\"\",NA(),$D{excel_row})", text)
        graph.write_formula(row, 5, "=NA()", text)
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
    chart.set_size({"width": 1220, "height": 650})
    graph.insert_chart("A82", chart)
    graph.freeze_panes(4, 0)

    control = ws["06_Контроль"]
    control.set_column("A:A", 18); control.set_column("B:B", 42); control.set_column("C:C", 18); control.set_column("D:D", 48)
    control.merge_range("A1:D1", "Контроль сроков и данных", title)
    control.merge_range("A2:D2", "Сигналы появятся после импорта КП и выбора СР. Неполные сведения не отмечаются как нормальный прогноз.", note)
    control.write_row("A4", ["Категория", "Объект", "Состояние", "Рекомендация"], header)
    control.freeze_panes(4, 0)

    acts = ws["07_Акты"]
    for col, width in enumerate((18, 16, 18, 16, 18, 22, 22, 46)): acts.set_column(col, col, width)
    acts.merge_range("A1:H1", "Акты по этапам КП", title)
    acts.merge_range("A2:H2", "Разрешено несколько актов на этап. В стоимость входят только акты со статусом «Подписан».", note)
    acts.write_row("A4", ["ID акта", "Этап КП", "Номер акта", "Дата акта", "Сумма без НДС",
                            "Статус", "Дата подписания", "Комментарий"], header)
    for row in range(4, 304):
        for col in range(8):
            fmt = money if col == 4 else date_input if col in (3, 6) else input_cell
            acts.write_blank(row, col, None, fmt)
    acts.data_validation("B5:B304", {"validate": "list", "source": kp_stage_list})
    acts.autofilter(3, 0, 303, 7); acts.freeze_panes(4, 0)

    ext = ws["08_Внешние_согласования"]
    ext.set_column("A:A", 18); ext.set_column("B:B", 45); ext.set_column("C:D", 16); ext.set_column("E:E", 22); ext.set_column("F:F", 45)
    ext.merge_range("A1:F1", "Внешние согласования", title)
    ext.merge_range("A2:F2", "Краткий перечень отправленных материалов и ответов. Это не реестр документооборота.", note)
    ext.write_row("A4", ["ID записи", "Материал / вопрос", "Дата направления", "Срок ответа", "Статус", "Комментарий"], header)
    for row in range(4, 54):
        for col in range(6): ext.write_blank(row, col, None, date_input if col in (2, 3) else input_cell)
    ext.freeze_panes(4, 0)

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
        ("schema_version", "1"),
    ]
    for row, values in enumerate(service_rows, 3): service_ws.write_row(row, 0, values, service)
    service_ws.hide()

    workbook.close()
    print(args.output)


if __name__ == "__main__":
    main()
