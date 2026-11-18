"""Deterministic dependency-network calculation for project accompanying works.

The workbook stores inputs. This module makes the network rules testable before
they are connected to the safe in-place Excel importer.
"""
from __future__ import annotations

from collections import defaultdict, deque
from datetime import date, timedelta
from typing import Any


def add_days(start: date, duration: int, day_type: str) -> date:
    """Return inclusive finish date. Zero duration is a milestone on start."""
    if duration < 0:
        raise ValueError("Длительность не может быть отрицательной.")
    if duration == 0:
        return start
    if day_type == "календарные":
        return start + timedelta(days=duration - 1)
    if day_type != "рабочие":
        raise ValueError("Тип дней должен быть «календарные» или «рабочие».")
    current, remaining = start, duration - 1
    while remaining:
        current += timedelta(days=1)
        if current.weekday() < 5:
            remaining -= 1
    return current


def next_day(value: date, day_type: str) -> date:
    return add_days(value + timedelta(days=1), 1, day_type)


def calculate(items: list[dict[str, Any]], stage_ends: dict[str, date] | None = None) -> dict[str, Any]:
    """Calculate FS/SS dependencies and return forecasts plus visible issues.

    An item is active when it has a non-empty stable ``id``.  Predecessors are
    dictionaries with ``id``, ``relation`` (FS/SS) and optional integer ``lag``.
    Facts and expected dates are never written back to a baseline input.
    """
    stage_ends = stage_ends or {}
    active = [item for item in items if str(item.get("id") or "").strip()]
    ids = [str(item["id"]).strip() for item in active]
    issues: list[dict[str, str]] = []
    duplicates = {item_id for item_id in ids if ids.count(item_id) > 1}
    for item_id in sorted(duplicates):
        issues.append({"id": item_id, "kind": "duplicate_id", "message": "Повторяется ID экземпляра СР."})
    by_id = {str(item["id"]).strip(): item for item in active if str(item["id"]).strip() not in duplicates}
    incoming: dict[str, int] = {item_id: 0 for item_id in by_id}
    successors: dict[str, list[str]] = defaultdict(list)
    for item_id, item in by_id.items():
        for link in item.get("predecessors") or []:
            predecessor = str(link.get("id") or "").strip()
            relation = str(link.get("relation") or "FS").upper()
            if predecessor not in by_id:
                issues.append({"id": item_id, "kind": "unknown_predecessor", "message": f"Не выбрана или не найдена предшествующая СР: {predecessor or 'не указана'}."})
                continue
            if relation not in {"FS", "SS"}:
                issues.append({"id": item_id, "kind": "unsupported_relation", "message": f"Неподдерживаемый тип связи: {relation}."})
                continue
            incoming[item_id] += 1; successors[predecessor].append(item_id)
    queue = deque(item_id for item_id, degree in incoming.items() if degree == 0)
    order: list[str] = []
    while queue:
        current = queue.popleft(); order.append(current)
        for successor in successors[current]:
            incoming[successor] -= 1
            if incoming[successor] == 0:
                queue.append(successor)
    if len(order) != len(by_id):
        cycle_ids = sorted(set(by_id) - set(order))
        for item_id in cycle_ids:
            issues.append({"id": item_id, "kind": "cycle", "message": "Обнаружен цикл зависимостей СР."})
    result: dict[str, dict[str, Any]] = {}
    for item_id in order:
        item = by_id[item_id]
        duration = item.get("duration")
        if not isinstance(duration, int):
            issues.append({"id": item_id, "kind": "missing_duration", "message": "Не задана длительность СР."})
            continue
        day_type = item.get("day_type", "календарные")
        own_start = item.get("fact_start") or item.get("planned_start")
        candidates = [own_start] if isinstance(own_start, date) else []
        for link in item.get("predecessors") or []:
            predecessor = str(link.get("id") or "").strip()
            if predecessor not in result:
                continue
            lag = link.get("lag", 0)
            if not isinstance(lag, int):
                issues.append({"id": item_id, "kind": "invalid_lag", "message": "Лаг связи должен быть целым числом."})
                continue
            pred = result[predecessor]
            relation = str(link.get("relation") or "FS").upper()
            base = pred["forecast_end"] if relation == "FS" else pred["forecast_start"]
            candidates.append(next_day(base, day_type) + timedelta(days=lag) if relation == "FS" else base + timedelta(days=lag))
        if not candidates:
            issues.append({"id": item_id, "kind": "missing_start", "message": "Не задано начало СР и нет рассчитываемой связи."})
            continue
        forecast_start = max(candidates)
        calculated_end = add_days(forecast_start, duration, day_type)
        fact_end, expected_end = item.get("fact_end"), item.get("expected_end")
        forecast_end = fact_end or expected_end or calculated_end
        if isinstance(forecast_end, date) and forecast_end < forecast_start:
            issues.append({"id": item_id, "kind": "impossible_dates", "message": "Окончание СР раньше расчётного начала."})
            continue
        result[item_id] = {"forecast_start": forecast_start, "forecast_end": forecast_end,
                           "calculated_end": calculated_end, "baseline_start": item.get("planned_start")}
        stage = str(item.get("stage") or "").strip()
        if stage and stage in stage_ends and forecast_end > stage_ends[stage]:
            issues.append({"id": item_id, "kind": "cp_breach", "message": f"Прогноз СР позднее окончания этапа КП {stage}."})
    return {"items": result, "issues": issues, "order": order}
