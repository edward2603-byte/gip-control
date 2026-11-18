# GIP Navigator

Excel-инструмент ГИП для ведения календарного плана, сопутствующих работ, согласований, актов и контроля сроков.

## Состав

- `templates/GIP_Navigator_STAGE0_PROTOTYPE.xlsx` — единая рабочая книга.
- `templates/Стрижов Луг 11.xlsx` — пример исходного КП для импорта.
- `tools/Import-Kp.ps1` — проверка и импорт КП в единую книгу.
- `tools/build_gip_navigator.py` — сборка книги со всеми листами и диаграммой Ганта.
- `tools/gipnav_import.py` — проверка семи полей КП перед импортом.

## Импорт КП

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\Import-Kp.ps1 -SourcePath .\templates\Стрижов Луг 11.xlsx -Apply
```

Скрипт проверяет структуру исходного файла и пересобирает только `GIP_Navigator_STAGE0_PROTOTYPE.xlsx`. Исходный КП не изменяется.

## Работа в книге

На листе `03_СР_проекта` жёлтые ячейки предназначены для ввода этапа КП, плановой даты, длительности, фактической и ожидаемой дат. После заполнения обновляются листы `05_График` и `06_Контроль`.

Для установки зависимостей:

```powershell
python -m pip install -r requirements.txt
```
