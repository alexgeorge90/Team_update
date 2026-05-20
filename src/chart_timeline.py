from datetime import datetime

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd

STATUS_COLORS = {
    "On Track": "#4A90D9",
    "In Progress": "#E8943A",
    "At Risk": "#D94A4A",
    "Blocked": "#8B1A1A",
    "Completed": "#5CB85C",
    "Closed": "#5CB85C",
    "Not Started": "#B0B0B0",
}

DEFAULT_COLOR = "#999999"


def generate_gantt(tasks_df: pd.DataFrame, output_path: str) -> None:
    tasks = tasks_df.sort_values("Start Date", ascending=False).reset_index(drop=True)
    n = len(tasks)
    if n == 0:
        print("No tasks to chart.")
        return

    fig, ax = plt.subplots(figsize=(14, max(3, n * 0.9 + 1.5)))

    for i, row in tasks.iterrows():
        start = row["Start Date"]
        end = row["Target End Date"]
        duration = (end - start).days
        status = row.get("Status", "")
        color = STATUS_COLORS.get(status, DEFAULT_COLOR)

        ax.barh(
            i, duration, left=start, height=0.5,
            color=color, edgecolor="white", linewidth=0.5,
            label=status if status not in [t.get_label() for t in ax.patches] else "",
        )

        members = str(row.get("Team Members", "")).replace(";", ",")
        ax.text(
            end + pd.Timedelta(days=3), i, members.strip(),
            va="center", fontsize=8, color="#555555",
        )

    today = pd.Timestamp(datetime.now().date())
    ax.axvline(today, color="#333333", linestyle="--", linewidth=1, alpha=0.8)
    ax.text(
        today, n - 0.1, " Today", fontsize=8, color="#333333",
        va="bottom", ha="left",
    )

    ax.set_yticks(range(n))
    ax.set_yticklabels(tasks["Task Name"], fontsize=10)
    ax.invert_yaxis()

    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    plt.xticks(fontsize=9, rotation=30, ha="right")

    all_starts = tasks["Start Date"]
    all_ends = tasks["Target End Date"]
    margin = pd.Timedelta(days=15)
    ax.set_xlim(all_starts.min() - margin, all_ends.max() + pd.Timedelta(days=60))

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_visible(False)
    ax.grid(axis="x", alpha=0.3, linestyle="--")

    handles = []
    seen = set()
    for status, color in STATUS_COLORS.items():
        if status in tasks["Status"].values and status not in seen:
            handles.append(plt.Rectangle((0, 0), 1, 1, fc=color, edgecolor="white"))
            seen.add(status)
    if handles:
        ax.legend(
            handles, list(seen),
            loc="lower right", fontsize=8, framealpha=0.9,
        )

    ax.set_title("Task Timeline", fontsize=14, fontweight="bold", pad=15)
    fig.tight_layout()
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"Gantt chart saved to {output_path}")
