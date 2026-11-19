"""Pure-Python regression model for the independent SR network.

The production calculation is in the workbook VBA. All durations and lags are
calendar days; the model has no workday mode.
"""
from __future__ import annotations

from collections import defaultdict, deque
from datetime import date, timedelta
from typing import Any


def add_days(value: date, day_count: int) -> date:
    return value + timedelta(days=day_count)


def calculate(items: list[dict[str, Any]], project_start: date | None = None) -> dict[str, Any]:
    """Calculate independent SR instances and return dates plus visible issues.

    Roots without an own start use ``project_start``. A later own start is a
    deliberate constraint; a date before project_start is an error. FS is the
    day after predecessor end, SS is predecessor start, and durations are inclusive.
    """
    active = [item for item in items if str(item.get("id") or "").strip()]
    ids = [str(item["id"]).strip() for item in active]
    issues: list[dict[str, str]] = []
    duplicates = {item_id for item_id in ids if ids.count(item_id) > 1}
    for item_id in sorted(duplicates):
        issues.append({"id": item_id, "kind": "duplicate_id", "message": "Повторяется ID экземпляра СР."})
    by_id = {str(item["id"]).strip(): item for item in active if str(item["id"]).strip() not in duplicates}
    incoming = {item_id: 0 for item_id in by_id}
    successors: dict[str, list[str]] = defaultdict(list)
    valid_links: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item_id, item in by_id.items():
        for link in item.get("predecessors") or []:
            predecessor = str(link.get("id") or "").strip()
            relation = str(link.get("relation") or "FS").upper()
            lag = link.get("lag", 0)
            if predecessor not in by_id:
                issues.append({"id": item_id, "kind": "unknown_predecessor", "message": f"Не выбрана или не найдена предшествующая СР: {predecessor or 'не указана'}."})
                continue
            if relation not in {"FS", "SS"}:
                issues.append({"id": item_id, "kind": "unsupported_relation", "message": f"Неподдерживаемый тип связи: {relation}."})
                continue
            if not isinstance(lag, int):
                issues.append({"id": item_id, "kind": "invalid_lag", "message": "Лаг связи должен быть целым числом."})
                continue
            valid_links[item_id].append({"id": predecessor, "relation": relation, "lag": lag})
            incoming[item_id] += 1
            successors[predecessor].append(item_id)
    queue = deque(item_id for item_id, degree in incoming.items() if degree == 0)
    order: list[str] = []
    while queue:
        current = queue.popleft()
        order.append(current)
        for successor in successors[current]:
            incoming[successor] -= 1
            if incoming[successor] == 0:
                queue.append(successor)
    if len(order) != len(by_id):
        for item_id in sorted(set(by_id) - set(order)):
            issues.append({"id": item_id, "kind": "cycle", "message": "Обнаружен цикл зависимостей СР."})
    result: dict[str, dict[str, Any]] = {}
    for item_id in order:
        item = by_id[item_id]
        duration = item.get("duration")
        if not isinstance(duration, int) or duration < 0:
            issues.append({"id": item_id, "kind": "missing_duration", "message": "Не задана длительность СР."})
            continue
        own_start = item.get("planned_start")
        if isinstance(own_start, date) and project_start and own_start < project_start:
            issues.append({"id": item_id, "kind": "before_project_start", "message": "Плановое начало раньше начала проекта."})
            continue
        candidates: list[date] = [own_start] if isinstance(own_start, date) else []
        if not valid_links[item_id] and not candidates and project_start:
            candidates.append(project_start)
        for link in valid_links[item_id]:
            predecessor = link["id"]
            if predecessor not in result:
                issues.append({"id": item_id, "kind": "missing_predecessor_date", "message": f"Нет срока предшествующей СР: {predecessor}."})
                continue
            pred = result[predecessor]
            if link["relation"] == "FS":
                candidates.append(add_days(pred["forecast_end"], link["lag"] + 1))
            else:
                candidates.append(add_days(pred["forecast_start"], link["lag"]))
        if not candidates:
            issues.append({"id": item_id, "kind": "missing_start", "message": "Не задано начало СР и нет рассчитываемой связи."})
            continue
        forecast_start = max(candidates)
        calculated_end = add_days(forecast_start, duration - 1)
        result[item_id] = {"forecast_start": forecast_start, "forecast_end": calculated_end, "calculated_end": calculated_end}
    return {"items": result, "issues": issues, "order": order}
