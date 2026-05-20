"""Generate task timeline and bandwidth utilization charts."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config import (
    DATA_FILE, WEEKLY_CAPACITY_HOURS, TEAM_MEMBERS,
    PAST_WEEKS, FUTURE_WEEKS, OUTPUT_DIR,
)
from src.data_reader import read_all_tasks, read_updates, read_allocations, resolve_task_statuses
from src.chart_timeline import generate_gantt
from src.chart_bandwidth import generate_bandwidth


def main():
    print(f"Reading data from {DATA_FILE}...")
    tasks_df = read_all_tasks(DATA_FILE)
    updates_df = read_updates(DATA_FILE)
    allocations_df = read_allocations(DATA_FILE)

    tasks_df = resolve_task_statuses(tasks_df, updates_df)

    print(f"  Tasks: {len(tasks_df)} rows")
    print(f"  Updates: {len(updates_df)} rows")
    print(f"  Allocations: {len(allocations_df)} rows")
    print()

    gantt_path = str(OUTPUT_DIR / "task_timeline.png")
    generate_gantt(tasks_df, gantt_path)

    bw_path = str(OUTPUT_DIR / "bandwidth.png")
    generate_bandwidth(
        updates_df, allocations_df,
        TEAM_MEMBERS, WEEKLY_CAPACITY_HOURS,
        PAST_WEEKS, FUTURE_WEEKS,
        bw_path,
    )

    print("\nDone! Charts are in the output/ folder.")


if __name__ == "__main__":
    main()
