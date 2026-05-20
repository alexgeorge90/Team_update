from datetime import datetime, timedelta

import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import numpy as np
import pandas as pd


def _monday_of(dt):
    return dt - timedelta(days=dt.weekday())


def generate_bandwidth(
    updates_df: pd.DataFrame,
    allocations_df: pd.DataFrame,
    team_members: list[str],
    capacity: int,
    past_weeks: int,
    future_weeks: int,
    output_path: str,
) -> None:
    today = datetime.now().date()
    current_monday = _monday_of(today)

    all_weeks = [
        current_monday - timedelta(weeks=past_weeks - i)
        for i in range(past_weeks + future_weeks + 1)
    ]
    all_weeks_ts = [pd.Timestamp(w) for w in all_weeks]
    current_monday_ts = pd.Timestamp(current_monday)
    n_weeks = len(all_weeks)

    all_tasks = sorted(
        set(updates_df["Task"].unique()) | set(allocations_df["Task"].unique())
    )

    rows = []
    row_labels = []
    row_types = []  # "task" or "total"
    member_boundaries = []  # y-positions where member groups start

    for member in team_members:
        member_boundaries.append(len(rows))
        member_updates = updates_df[updates_df["Team Member"] == member]
        member_allocs = allocations_df[allocations_df["Team Member"] == member]

        member_tasks = sorted(
            set(member_updates["Task"].unique()) | set(member_allocs["Task"].unique())
        )

        for task in member_tasks:
            row = np.zeros(n_weeks)
            task_updates = member_updates[member_updates["Task"] == task]
            for _, r in task_updates.iterrows():
                if r["Week"] in all_weeks_ts:
                    idx = all_weeks_ts.index(r["Week"])
                    row[idx] += r["Hours Spent"]

            task_allocs = member_allocs[member_allocs["Task"] == task]
            for _, r in task_allocs.iterrows():
                if r["Week Starting"] in all_weeks_ts:
                    idx = all_weeks_ts.index(r["Week Starting"])
                    row[idx] += r["Planned Hours"]

            rows.append(row)
            row_labels.append(f"  {task}")
            row_types.append("task")

        total_row = np.zeros(n_weeks)
        for r in rows[member_boundaries[-1]:]:
            total_row += r
        rows.append(total_row)
        row_labels.append(f"{member} TOTAL")
        row_types.append("total")

    data = np.array(rows)
    n_rows = len(rows)

    # Separate color maps for actual (past/current) and planned (future) weeks
    current_idx = all_weeks_ts.index(current_monday_ts) if current_monday_ts in all_weeks_ts else past_weeks

    actual_cmap = mcolors.LinearSegmentedColormap.from_list(
        "actual", ["#F0F4F8", "#2B6CB0"]
    )
    planned_cmap = mcolors.LinearSegmentedColormap.from_list(
        "planned", ["#FFF8F0", "#DD6B20"]
    )

    max_hours = max(data.max(), 1)

    fig_height = max(4, n_rows * 0.45 + 2)
    fig, ax = plt.subplots(figsize=(max(10, n_weeks * 0.85 + 3), fig_height))

    for r_idx in range(n_rows):
        for c_idx in range(n_weeks):
            val = data[r_idx, c_idx]
            is_planned = c_idx > current_idx
            is_total = row_types[r_idx] == "total"

            if is_planned:
                norm_val = val / max_hours if max_hours > 0 else 0
                bg = planned_cmap(norm_val * 0.85)
            else:
                norm_val = val / max_hours if max_hours > 0 else 0
                bg = actual_cmap(norm_val * 0.85)

            if is_total and val > capacity:
                bg = (0.95, 0.3, 0.3, 0.6)

            rect = plt.Rectangle(
                (c_idx, n_rows - r_idx - 1), 1, 1,
                facecolor=bg,
                edgecolor="white",
                linewidth=1.5 if is_total else 0.8,
            )
            ax.add_patch(rect)

            if val > 0:
                text_color = "white" if norm_val > 0.55 else "#333333"
                if is_total and val > capacity:
                    text_color = "white"
                fontweight = "bold" if is_total else "normal"
                fontsize = 8.5 if is_total else 7.5

                ax.text(
                    c_idx + 0.5, n_rows - r_idx - 0.5,
                    f"{val:.0f}",
                    ha="center", va="center",
                    fontsize=fontsize, fontweight=fontweight,
                    color=text_color,
                )

    # Horizontal lines to separate member groups
    for boundary_idx in member_boundaries[1:]:
        total_offset = sum(1 for t in row_types[:boundary_idx] if t == "total")
        y = n_rows - boundary_idx
        ax.axhline(y, color="#333333", linewidth=1.5, zorder=3)

    # Vertical line for "today"
    if current_monday_ts in all_weeks_ts:
        today_x = current_idx + 1
        ax.axvline(today_x, color="#333333", linestyle="--", linewidth=1.2, alpha=0.7, zorder=3)
        ax.text(
            today_x, n_rows + 0.3, "Today",
            ha="center", fontsize=7, color="#333333",
        )

    ax.set_xlim(0, n_weeks)
    ax.set_ylim(0, n_rows)

    ax.set_yticks([n_rows - i - 0.5 for i in range(n_rows)])
    ax.set_yticklabels(row_labels, fontsize=8)

    for i, label in enumerate(row_labels):
        if row_types[i] == "total":
            tick = ax.get_yticklabels()[i]
            tick.set_fontweight("bold")
            tick.set_fontsize(9)

    week_labels = []
    for w in all_weeks:
        week_labels.append(pd.Timestamp(w).strftime("%b %d"))
    ax.set_xticks([i + 0.5 for i in range(n_weeks)])
    ax.set_xticklabels(week_labels, fontsize=7.5, rotation=45, ha="right")

    ax.tick_params(axis="both", which="both", length=0)
    ax.spines[:].set_visible(False)

    # Legend
    from matplotlib.patches import Patch
    legend_handles = [
        Patch(facecolor="#2B6CB0", alpha=0.7, label="Actual hours"),
        Patch(facecolor="#DD6B20", alpha=0.7, label="Planned hours"),
        Patch(facecolor=(0.95, 0.3, 0.3, 0.6), label=f"Over {capacity}h capacity"),
    ]
    ax.legend(
        handles=legend_handles, loc="upper left",
        bbox_to_anchor=(0, -0.12), ncol=3, fontsize=8, framealpha=0.9,
    )

    ax.set_title(
        f"Weekly Bandwidth Utilization (hours/week, capacity = {capacity}h)",
        fontsize=13, fontweight="bold", pad=15,
    )

    fig.tight_layout()
    fig.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"Bandwidth chart saved to {output_path}")
