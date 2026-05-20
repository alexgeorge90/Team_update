import pandas as pd


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace from column names."""
    df.columns = df.columns.str.strip()
    return df


def read_tasks(path, sheet_name="Tasks") -> pd.DataFrame:
    df = _clean_columns(pd.read_excel(path, sheet_name=sheet_name))
    df = df.dropna(subset=["Task Name"])
    df["Start Date"] = pd.to_datetime(df["Start Date"], dayfirst=True)
    df["Target End Date"] = pd.to_datetime(df["Target End Date"], dayfirst=True)
    return df


def read_updates(path, sheet_name="Weekly Updates") -> pd.DataFrame:
    df = _clean_columns(pd.read_excel(path, sheet_name=sheet_name))
    df = df.dropna(subset=["Task"])
    df["Update Date"] = pd.to_datetime(df["Update Date"], dayfirst=True)
    df["Team Member"] = df["Team Member"].str.strip()
    df["Task"] = df["Task"].str.strip()
    df["Hours Spent"] = pd.to_numeric(df["Hours Spent"], errors="coerce").fillna(0)
    # Week column = Monday of the update week
    df["Week"] = df["Update Date"] - pd.to_timedelta(df["Update Date"].dt.weekday, unit="D")
    return df


def read_allocations(path, sheet_name="Allocations") -> pd.DataFrame:
    df = _clean_columns(pd.read_excel(path, sheet_name=sheet_name))
    df = df.dropna(subset=["Task"])
    df["Week Starting"] = pd.to_datetime(df["Week Starting"], dayfirst=True)
    df["Team Member"] = df["Team Member"].str.strip()
    df["Task"] = df["Task"].str.strip()
    df["Planned Hours"] = pd.to_numeric(df["Planned Hours"], errors="coerce").fillna(0)
    return df


def resolve_task_statuses(tasks_df: pd.DataFrame, updates_df: pd.DataFrame) -> pd.DataFrame:
    """
    Fill in Status for tasks by picking the latest Weekly Update status.
    Handles the case where Excel formulas aren't evaluated (openpyxl reads formula text).
    """
    if updates_df.empty:
        tasks_df["Status"] = tasks_df["Status"].fillna("Not Started")
        return tasks_df

    latest = (
        updates_df
        .sort_values("Update Date")
        .groupby("Task")["Status"]
        .last()
    )

    def pick_status(row):
        existing = row.get("Status")
        if pd.notna(existing) and isinstance(existing, str) and not existing.startswith("="):
            return existing
        task_name = row["Task Name"]
        return latest.get(task_name, "Not Started")

    tasks_df["Status"] = tasks_df.apply(pick_status, axis=1)
    return tasks_df
