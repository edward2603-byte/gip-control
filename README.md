# GIP Navigator

Универсальный Excel-инструмент для календарного плана, сопутствующих работ (СР), согласования сроков, контроля и актирования. Рабочая книга создаётся отдельно для каждого проекта; мастер-шаблон не содержит проектных данных.

## Состав

- `templates/GIP_Navigator_TEMPLATE.xlsx` — пустой неизменяемый мастер-шаблон.
- `tools/New-ProjectFromTemplate.ps1` — безопасное создание отдельной книги проекта.
- `tools/Import-Kp.ps1` — предпросмотр исходного КП. Обновление рабочей книги будет включено только с безопасным построчным импортом.
- `tests/fixtures/` — исходные КП для тестов.
- `examples/` — демонстрационные книги, не используемые как мастер.

## Создание книги проекта

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\New-ProjectFromTemplate.ps1 `
  -ProjectFilePath 'C:\Проекты\Проект_001\GIP_Navigator_Проект_001.xlsx'
```

Команда откажется заменять существующий файл. Для намеренной замены нужен ключ `-ReplaceExisting`; перед ней создаётся резервная копия рядом с файлом.

## Предпросмотр КП

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\Import-Kp.ps1 `
  -SourcePath '.\tests\fixtures\Стрижов Луг 11.xlsx' -PreviewOnly
```

Исходный файл КП только читается. На текущем шаге ключ `-Apply` намеренно не изменяет книгу: полная пересборка запрещена, поскольку может уничтожить введённые ГИПом СР, факты, акты и согласованный план.

## Проверка шаблона

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\Test-Template.ps1
```
