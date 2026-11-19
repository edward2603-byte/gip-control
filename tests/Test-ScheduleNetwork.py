from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parents[1] / "tools"))
from schedule_network import calculate


def main() -> None:
    project_start = min(date(2027, 3, 10), date(2027, 3, 1), date(2027, 4, 1))
    network = [
        {"id": "ROOT", "duration": 1},
        {"id": "FS", "duration": 3, "predecessors": [{"id": "ROOT", "relation": "FS", "lag": 0}]},
        {"id": "SS", "duration": 1, "predecessors": [{"id": "ROOT", "relation": "SS", "lag": 2}]},
    ]
    planned = calculate(network, project_start)
    assert planned["items"]["ROOT"]["forecast_start"] == date(2027, 3, 1)
    assert planned["items"]["FS"]["forecast_start"] == date(2027, 3, 2)
    assert planned["items"]["FS"]["forecast_end"] == date(2027, 3, 4)
    assert planned["items"]["SS"]["forecast_start"] == date(2027, 3, 3)

    weekend = calculate([{"id": "W", "duration": 2}], date(2027, 5, 1))
    assert weekend["items"]["W"]["forecast_start"] == date(2027, 5, 1)
    assert weekend["items"]["W"]["forecast_end"] == date(2027, 5, 2)

    broken = calculate([{"id": "A", "duration": 1, "predecessors": [{"id": "B", "relation": "FS"}]},
                        {"id": "B", "duration": 1, "predecessors": [{"id": "A", "relation": "FS"}]}], project_start)
    assert {issue["kind"] for issue in broken["issues"]} == {"cycle"}
    missing = calculate([{"id": "A", "duration": 1, "predecessors": [{"id": "MISSING", "relation": "FS"}]}], project_start)
    assert any(issue["kind"] == "unknown_predecessor" for issue in missing["issues"])
    print("PASS: независимая сеть СР рассчитала project_start, FS/SS, календарные дни, цикл и отсутствующую СР.")


if __name__ == "__main__":
    main()
