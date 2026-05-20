# Power Apps Task Tracker -- Setup Guide

Complete step-by-step guide to build your team task management system using **Power Apps + SharePoint Lists + Python charts**.

| Scope | Count |
|-------|-------|
| SharePoint Lists | 3 |
| Power App Screens | 3 |
| Team Members | 15 |
| Sub-Units | 2-3 |

**Estimated total setup time: 1-2 hours**

---

## Architecture

```
Team (15 people)  -->  Power Apps  -->  SharePoint Lists (source of truth)
                                              |
PM (you)          -->  Power Apps  -->        |
                                              v
                                      OneDrive sync / Export to Excel
                                              |
                                              v
                                      Python scripts --> chart PNGs
                                              |
                                              v
                                      Copilot Agent --> leadership email
```

---

## Phase 1: Create SharePoint Lists (10-15 min)

Go to your SharePoint site > **Site Contents** > **New** > **List** > **Blank list**. Create these 3 lists.

### List 1: Tasks

> **How to create:** Site Contents > New > List > Blank list > Name it "Tasks"

| Column Name | Type | Required | Settings |
|-------------|------|----------|----------|
| Title | Single line (rename to **Task Name**) | Yes | List Settings > click Title > change to "Task Name" |
| Description | Multiple lines of text | No | Plain text, 6 lines |
| Sub_Unit | Choice | Yes | Choices: `Unit A, Unit B, Unit C` (customize to your sub-units) |
| Team_Members | Multiple lines of text | Yes | Use text field (e.g. "Alex; Srikanth"). Power Apps will handle the picker UI |
| Start_Date | Date | Yes | Date only, no time. Friendly format |
| Target_End_Date | Date | Yes | Date only, no time |
| Status | Choice | Yes | Choices: `Not Started, In Progress, On Track, At Risk, Blocked, Closed`. Default: Not Started |
| Priority | Choice | No | Choices: `High, Medium, Low`. Default: Medium |
| Task_ID | Calculated (see below) | Auto | Formula gives DRM001, DRM002, etc. |

> **Calculated column for Task_ID:**
> After creating all other columns: List Settings > Create column > Type: **Calculated** > Formula:
> ```
> ="DRM"&RIGHT("000"&[ID],3)
> ```
> Return type: Single line of text

### List 2: Weekly Updates

| Column Name | Type | Required | Settings |
|-------------|------|----------|----------|
| Title | Single line (rename to **Team_Member**) | Yes | Rename default Title column to Team_Member |
| Task | Lookup | Yes | Get information from: **Tasks** list, column: **Task Name** (Title) |
| Status | Choice | Yes | Choices: `On Track, At Risk, Blocked, Closed` |
| Progress_Notes | Multiple lines of text | No | Plain text |
| Dependency | Single line of text | No | |
| Planned_Next_Action | Multiple lines of text | No | Plain text |
| Hours_Spent | Number | Yes | Min: 0, Max: 80, Decimal places: 0 |
| Update_Date | Date | Yes | Date only. Default: Today's date |
| Sub_Unit | Lookup | No | Get from: Tasks list, column: Sub_Unit (auto-fills via Power Apps) |

> **Lookup column setup:** For the Task column: Add column > Lookup > Select "Tasks" as source list > Select "Task Name" as the column to show. This creates the relationship between updates and tasks.

### List 3: Allocations

| Column Name | Type | Required | Settings |
|-------------|------|----------|----------|
| Title | Single line (rename to **Team_Member**) | Yes | Rename default Title column |
| Task | Lookup | Yes | Get from: Tasks list, column: Task Name (Title) |
| Week_Starting | Date | Yes | Date only. Always use Monday of the week |
| Planned_Hours | Number | Yes | Min: 0, Max: 60, Decimal places: 0 |
| Sub_Unit | Lookup | No | Get from: Tasks list, column: Sub_Unit |

### SharePoint Views

Create these views on the **Tasks** list for quick filtering:

| View Name | Filter | Sort |
|-----------|--------|------|
| Active Tasks | Status is not equal to Closed | Start_Date ascending |
| My Tasks (per person) | Team_Members contains [Current User] | Status |
| By Sub-Unit | Group by Sub_Unit | Status, then Start_Date |
| Archived / Closed | Status equals Closed | Target_End_Date descending |

> List Settings > Views > Create View > Standard View

---

## Phase 2: Build Power App (30-45 min)

Go to **make.powerapps.com** > Create > Blank app > **Blank canvas app** > **Tablet layout**

### Step 1: Connect Data Sources

In Power Apps Studio: **Data** (left panel) > Add data > SharePoint > Select your site > Check all 3 lists (Tasks, Weekly Updates, Allocations) > Connect

### Step 2: Screen 1 -- Team Update (for team members)

**Screen name:** `TeamUpdateScreen`
**Layout:** Simple form, 30-second entry

| Control | Type | Property / Formula |
|---------|------|--------------------|
| Team Member Dropdown | Dropdown | `Items: ["Alex","Srikanth","Enith", ... your 15 members]` |
| Task Dropdown | Dropdown | `Items: Filter(Tasks, Status <> "Closed").Task_Name` |
| Status Dropdown | Dropdown | `Items: ["On Track","At Risk","Blocked","Closed"]` |
| Progress Notes | Text Input | Mode: MultiLine, HintText: "What did you work on?" |
| Dependency | Text Input | HintText: "Any blockers or dependencies?" |
| Planned Next Action | Text Input | Mode: MultiLine, HintText: "What's next?" |
| Hours Spent | Text Input | Format: Number, Min: 0, Max: 80 |
| Submit Button | Button | OnSelect formula below |

**Submit Button -- OnSelect formula:**

```
Patch(
  'Weekly Updates',
  Defaults('Weekly Updates'),
  {
    Team_Member: ddTeamMember.Selected.Value,
    Task: {
      Id: LookUp(Tasks,
        Task_Name = ddTask.Selected.Value
      ).ID,
      Value: ddTask.Selected.Value
    },
    Status: {Value: ddStatus.Selected.Value},
    Progress_Notes: txtProgressNotes.Text,
    Dependency: txtDependency.Text,
    Planned_Next_Action: txtNextAction.Text,
    Hours_Spent: Value(txtHours.Text),
    Update_Date: Today()
  }
);

// Auto-update Task status in Tasks list
Patch(
  Tasks,
  LookUp(Tasks,
    Task_Name = ddTask.Selected.Value
  ),
  {
    Status: {Value: ddStatus.Selected.Value}
  }
);

// Reset form
Reset(ddTask);
Reset(ddStatus);
Reset(txtProgressNotes);
Reset(txtDependency);
Reset(txtNextAction);
Reset(txtHours);

Notify("Update submitted!", NotificationType.Success);
```

> **Key: Filtered Task dropdown** -- The formula `Filter(Tasks, Status <> "Closed")` ensures closed tasks never appear in the dropdown. This solves the Excel limitation.

### Step 3: Screen 2 -- PM Task Board (for PM)

**Screen name:** `TaskBoardScreen`
**Layout:** Gallery view of all tasks + edit form

| Control | Type | Property / Formula |
|---------|------|--------------------|
| Sub-Unit Filter | Dropdown | `Items: ["All","Unit A","Unit B","Unit C"]` |
| Status Filter | Dropdown | `Items: ["All","Active Only","Closed Only"]` |
| Task Gallery | Vertical Gallery | Items formula below |
| Add Task Button | Button | Navigate to NewTaskForm |
| Edit Form | Edit Form | DataSource: Tasks, Item: Gallery.Selected |

**Task Gallery -- Items formula (with filters):**

```
SortByColumns(
  Filter(
    Tasks,
    // Sub-unit filter
    ddSubUnit.Selected.Value = "All"
      || Sub_Unit.Value = ddSubUnit.Selected.Value,
    // Status filter
    ddStatusFilter.Selected.Value = "All"
      || (ddStatusFilter.Selected.Value = "Active Only"
          && Status.Value <> "Closed")
      || (ddStatusFilter.Selected.Value = "Closed Only"
          && Status.Value = "Closed")
  ),
  "Start_Date", SortOrder.Ascending
)
```

**Gallery card layout (per row):**

| Element | Shows |
|---------|-------|
| Title label | `ThisItem.Task_Name` |
| Status pill | `ThisItem.Status.Value` (color-coded) |
| Team label | `ThisItem.Team_Members` |
| Date range | `Text(ThisItem.Start_Date,"mmm dd") & " - " & Text(ThisItem.Target_End_Date,"mmm dd")` |
| Task ID | `ThisItem.Task_ID` |

### Step 4: Screen 3 -- Allocation Planner (for PM)

**Screen name:** `AllocationScreen`
**Layout:** Week selector + member/task grid for planning hours

| Control | Type | Property / Formula |
|---------|------|--------------------|
| Week Picker | Date Picker | DefaultDate: Today(), start of week logic |
| Member Filter | Dropdown | `Items: ["All","Alex","Srikanth","Enith", ...]` |
| Allocation Gallery | Vertical Gallery | `Items: Filter(Allocations, Week_Starting = selectedWeek)` |
| Add Allocation | Button | Opens form to add: Member + Task + Planned Hours for the selected week |
| Total Hours Label | Label | `Sum(Filter(Allocations, Week_Starting=selectedWeek && Team_Member="Alex"), Planned_Hours)` |

**Add Allocation -- OnSelect formula:**

```
Patch(
  Allocations,
  Defaults(Allocations),
  {
    Team_Member: ddAllocMember.Selected.Value,
    Task: {
      Id: LookUp(Tasks,
        Task_Name = ddAllocTask.Selected.Value
      ).ID,
      Value: ddAllocTask.Selected.Value
    },
    Week_Starting: dpWeekPicker.SelectedDate,
    Planned_Hours: Value(txtPlannedHours.Text)
  }
);

Notify("Allocation saved!", NotificationType.Success);
```

---

## Phase 3: Python Charts Integration

Export data from SharePoint Lists, run Python scripts to generate chart PNGs.

### Export workflow (weekly, takes 2 min)

| Step | Action | Result |
|------|--------|--------|
| 1 | Open Tasks list in SharePoint > Export > Export to Excel | Downloads Tasks.xlsx |
| 2 | Open Weekly Updates list > Export > Export to Excel | Downloads Weekly Updates.xlsx |
| 3 | Open Allocations list > Export > Export to Excel | Downloads Allocations.xlsx |
| 4 | Move all 3 files to your project's `exports/` folder | Ready for Python |
| 5 | Run: `python src/main.py` | Generates task_timeline.png + bandwidth.png |

> **Simpler alternative: OneDrive sync** -- Instead of manual exports, sync the SharePoint document library via OneDrive. The synced file updates automatically on your local machine. Point the `.env` to the synced path.

### Python .env configuration

```ini
# Option A: 3 separate exported files (SharePoint mode)
DATA_MODE=sharepoint
SP_TASKS_FILE=exports/Tasks.xlsx
SP_UPDATES_FILE=exports/Weekly Updates.xlsx
SP_ALLOCATIONS_FILE=exports/Allocations.xlsx

# Option B: Single workbook (current Excel setup)
DATA_MODE=workbook
DATA_FILE=sample_data/team_data.xlsx

# Common settings
WEEKLY_CAPACITY_HOURS=40
TEAM_MEMBERS=Alex,Srikanth,Enith
PAST_WEEKS=8
FUTURE_WEEKS=4
```

---

## Phase 4: Finishing Touches

### Archiving closed tasks

With SharePoint Lists + Power Apps, archiving is handled natively. No separate script needed.

| Approach | How |
|----------|-----|
| **Views (recommended)** | Active Tasks view filters Status != Closed. Closed tasks stay in the same list but are hidden from default view. |
| **Power Apps dropdown** | `Filter(Tasks, Status <> "Closed")` already excludes closed tasks from the Team Update form. |
| **Separate list (optional)** | If you want physical separation, create an "Archived Tasks" list and use a Power Automate flow (SharePoint trigger) to move items when Status = Closed. |

### Notifications (if Power Automate allows)

Check if your DLP policy allows the SharePoint + Outlook connectors together. If yes:

| Trigger | Action | Purpose |
|---------|--------|---------|
| When item modified (Tasks list) | Send email to PM | Alert when a task status changes |
| Recurrence (every Monday 9am) | Send email to team | Weekly reminder to submit updates |
| When item created (Weekly Updates) | Update Tasks list status | Auto-sync status back to Tasks (alternative to the Patch in Power Apps) |

### Power Apps sharing

| Step | Action |
|------|--------|
| 1 | In Power Apps Studio: File > Save > Publish |
| 2 | Share: Apps panel > ... > Share > Enter team members' emails > Assign "User" role |
| 3 | Team accesses via: make.powerapps.com or Power Apps mobile app |
| 4 | Optional: Pin to Teams channel as a tab for easy access |

---

## Complete Workflow Summary

| Who | What | Where | When |
|-----|------|-------|------|
| Team (15 people) | Submit weekly updates | Power Apps (Team Update screen) | Weekly or on progress |
| PM (you) | Add/manage tasks | Power Apps (Task Board screen) or SharePoint | As needed |
| PM (you) | Plan allocations | Power Apps (Allocation screen) | Weekly planning |
| PM (you) | Generate charts | Export lists > `python src/main.py` | Before leadership email |
| Copilot Agent | Draft leadership email | Reads from SharePoint + attaches chart PNGs | Weekly |

### What this gives you over Excel

- Native filtered dropdowns (no closed tasks in picker)
- Mobile access for the team
- Proper multi-user support without file locking
- Sub-unit grouping
- Version history on every record
- Professional app interface your team and leadership will appreciate
