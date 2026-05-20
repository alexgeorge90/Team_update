# Task Reporting Chart Generator

Generates two charts for the weekly leadership email:
- **Task Timeline** (Gantt) -- each task's duration across months/quarters, color-coded by status
- **Bandwidth Utilization** -- actual hours (past) and planned allocations (future) per team member

## Quick Start

```powershell
# Activate the virtual environment
.\.venv\Scripts\Activate.ps1

# Generate charts from sample data
python src/main.py
```

Charts are saved to `output/task_timeline.png` and `output/bandwidth.png`.

## Configuration

Copy `.env.example` to `.env` and update the path to your real Excel export:

```
DATA_FILE=C:\Users\alexg\OneDrive\Reports\team_data.xlsx
WEEKLY_CAPACITY_HOURS=40
TEAM_MEMBERS=Alex,Bob,Carol
PAST_WEEKS=8
FUTURE_WEEKS=4
```

The Excel file has 3 tabs: **Tasks**, **Weekly Updates**, **Allocations**.

## SharePoint List Setup

Create 3 SharePoint Lists on your site. Export all 3 to the same Excel workbook (or maintain a single workbook with 3 tabs that mirrors them). Below are the columns for each.

### List 1: Tasks

| Column | Type | Notes |
|--------|------|-------|
| Task Name | Single line of text | Required, enforce unique |
| Description | Multiple lines of text | |
| Team Members | Person or Group | Allow multiple selections |
| Start Date | Date | Required |
| Target End Date | Date | Required |
| Status | Choice | Not Started, In Progress, On Track, At Risk, Blocked, Completed |
| Dependency | Single line of text | |
| Next Milestone | Single line of text | |

### List 2: Weekly Updates

| Column | Type | Notes |
|--------|------|-------|
| Update Date | Date | Default = today |
| Team Member | Person | Single person |
| Task | Lookup | From "Tasks" list, showing "Task Name" |
| Status | Choice | On Track, At Risk, Blocked, Completed |
| Progress | Multiple lines of text | |
| Blockers | Multiple lines of text | Optional |
| Hours Spent | Number | Min 0, max 80 |

### List 3: Allocations

| Column | Type | Notes |
|--------|------|-------|
| Week Starting | Date | Always use a Monday |
| Team Member | Person | Single person |
| Task | Lookup | From "Tasks" list, showing "Task Name" |
| Planned Hours | Number | Min 0, max 60 |

### Exporting to Excel

You can either:
1. **Maintain a single Excel workbook** with 3 tabs (Tasks, Weekly Updates, Allocations) that you keep in sync with SharePoint manually or via the Power App
2. **Export each list separately** to Excel and combine into one workbook with 3 tabs

Before running the script, refresh the data if using SharePoint-linked workbooks (open, Refresh All, save).

## Power App Setup

Create a Canvas App in Power Apps with two screens:

### Screen 1: Team Update

- Connect to "Tasks" and "Weekly Updates" SharePoint Lists
- Auto-detect user: `User().Email`
- Gallery showing tasks where current user is in Team Members
- Each card: Status dropdown, Progress text, Blockers text, Hours number
- Submit button patches to "Weekly Updates" list

### Screen 2: PM Planning

- Connect to "Tasks" and "Allocations" SharePoint Lists
- Date picker for target week (snapped to Monday)
- Grid: rows = team members x tasks, cells = planned hours
- Totals per person per week, red if > 40h
- Save button patches to "Allocations" list

Publish the app and add it as a Teams tab. Set up a recurring Friday Teams message with the app link.

## Weekly Workflow

1. **Friday**: Team fills updates in Power App (~2 min each)
2. **Monday**: PM sets allocations in Power App, refreshes Excel exports, runs `python src/main.py`
3. **Monday**: Copilot Agent drafts email, PM attaches the 2 chart PNGs, sends to leadership
