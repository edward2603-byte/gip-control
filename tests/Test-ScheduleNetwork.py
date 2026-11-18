from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1] / "tools"))
from schedule_network import calculate


def main() -> None:
    chain = [
        {"id": "ID-GR", "planned_start": date(2026, 1, 1), "duration": 1, "day_type": "календарные", "expected_end": date(2026, 1, 10)},
        {"id": "GR", "planned_start": date(2026, 1, 1), "duration": 3, "day_type": "календарные", "stage": "A", "predecessors": [{"id": "ID-GR", "relation": "FS", "lag": 0}]},
    ]
    delayed = calculate(chain, {"A": date(2026, 1, 11)})
    assert delayed["items"]["GR"]["forecast_start"] == date(2026, 1, 11)
    assert any(issue["kind"] == "cp_breach" for issue in delayed["issues"])

    parallel = [
        {"id": "FIELD", "planned_start": date(2026, 2, 1), "duration": 30, "day_type": "календарные"},
        {"id": "M20", "duration": 0, "day_type": "календарные", "predecessors": [{"id": "FIELD", "relation": "SS", "lag": 19}]},
        {"id": "PLACE", "duration": 20, "day_type": "календарные", "predecessors": [{"id": "M20", "relation": "FS", "lag": 0}]},
    ]
    planned = calculate(parallel)
    assert planned["items"]["M20"]["forecast_start"] == date(2026, 2, 20)
    assert planned["items"]["PLACE"]["forecast_start"] == date(2026, 2, 21)
    assert planned["items"]["PLACE"]["forecast_start"] < planned["items"]["FIELD"]["forecast_end"]

    broken = calculate([{"id": "A", "duration": 1, "day_type": "календарные", "predecessors": [{"id": "B", "relation": "FS"}]},
                        {"id": "B", "duration": 1, "day_type": "календарные", "predecessors": [{"id": "A", "relation": "FS"}]}])
    assert {issue["kind"] for issue in broken["issues"]} == {"cycle"}
    print("PASS: сеть СР рассчитала задержку ИД, параллельную веху 20-го дня и цикл.")


if __name__ == "__main__":
    main()
