# Task Reporting Chart Generator

Generates two charts for the weekly leadership email:
- **Task Timeline** (Gantt) -- each task's duration across months/quarters, color-coded by status
- **Bandwidth Utilization** -- actual hours (past) and planned allocations (future) per team member

---

## Setup on a New Laptop (Step by Step)

### Step 1: Install Python

1. Go to https://www.python.org/downloads/
2. Click the big **"Download Python 3.12.x"** button
3. Run the installer
4. **IMPORTANT:** Check the box that says **"Add python.exe to PATH"** at the bottom of the first screen
5. Click **Install Now**
6. Once done, close the installer

**Verify:** Open a new PowerShell window (search "PowerShell" in Start menu) and type:
```
python --version
```
You should see something like `Python 3.12.10`. If you see "Python was not found", restart your laptop and try again.

### Step 2: Download the Repository

1. Go to https://github.com/alexgeorge90/Team_update
2. Click the green **"Code"** button
3. Click **"Download ZIP"**
4. Extract the ZIP to a folder, e.g. `C:\Users\YourName\Team_update`

Or if you have Git installed:
```
git clone https://github.com/alexgeorge90/Team_update.git
```

### Step 3: Open in VS Code

1. Open VS Code
2. File > Open Folder > select the `Team_update` folder
3. Open a terminal in VS Code: **Terminal > New Terminal** (or press Ctrl+`)

### Step 4: Create Virtual Environment and Install Dependencies

Run these commands one by one in the VS Code terminal:

```powershell
python -m venv .venv
```

```powershell
.\.venv\Scripts\Activate.ps1
```

If you get a red error about "execution policy", run this first and then retry:
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

```powershell
pip install -r requirements.txt
```

Wait for it to finish (it will download pandas, matplotlib, etc.).

### Step 5: Create Your Config File

1. In the project folder, find the file `.env.example`
2. Copy it and rename the copy to `.env`
3. Open `.env` and update if needed:

```
DATA_FILE=sample_data/team_data.xlsx
WEEKLY_CAPACITY_HOURS=40
TEAM_MEMBERS=Alex,Srikanth,Enith
PAST_WEEKS=8
FUTURE_WEEKS=4
```

For now, leave it as-is to test with sample data. Later, change `DATA_FILE` to point to your real Excel file.

### Step 6: Generate Charts

```powershell
python src/main.py
```

You should see:
```
Reading data from ...\sample_data\team_data.xlsx...
  Tasks: 5 rows
  Updates: 42 rows
  Allocations: 27 rows

Gantt chart saved to ...\output\task_timeline.png
Bandwidth chart saved to ...\output\bandwidth.png

Done! Charts are in the output/ folder.
```

Open the `output/` folder to see your two chart images.

---

## Using Your Real Data

### Option A: Edit the Excel file directly

1. Open `sample_data/team_data.xlsx` in Excel
2. Replace the sample data with your real tasks, updates, and allocations
3. The file has 3 tabs: **Tasks**, **Weekly Updates**, **Allocations**
4. Dropdowns and validations are built in -- just use them
5. Save and run `python src/main.py`

### Option B: Point to a different Excel file

1. Put your Excel file anywhere (e.g., OneDrive folder)
2. Edit `.env` and set `DATA_FILE` to the full path:
   ```
   DATA_FILE=C:\Users\YourName\OneDrive\Reports\team_data.xlsx
   ```
3. Run `python src/main.py`

### Regenerating the Template

If you need a fresh Excel template with all validations and sample data:
```powershell
python create_sample_data.py
```

---

## Excel File Structure

The workbook has 4 tabs:

### Tasks (managed by PM)
| Column | Description |
|--------|-------------|
| Task ID | Auto-generated (T001, T002, ...) |
| Task Name | Name of the task |
| Description | What the task involves |
| Team Members | Dropdown: pick 1, 2, or all 3 members |
| Start Date | When work began |
| Target End Date | Expected completion |
| Status | Auto-filled from latest Weekly Update entry |

### Weekly Updates (filled by team)
| Column | Description |
|--------|-------------|
| Update Date | Date of the update |
| Team Member | Dropdown: Alex, Srikanth, or Enith |
| Task | Dropdown: picks from Tasks tab |
| Status | Dropdown: On Track, At Risk, Blocked, Closed |
| Progress | What was accomplished |
| Blockers | Any blockers (optional) |
| Hours Spent | Hours this week (0-80) |

### Allocations (filled by PM)
| Column | Description |
|--------|-------------|
| Week Starting | Monday of the target week |
| Team Member | Dropdown |
| Task | Dropdown: picks from Tasks tab |
| Planned Hours | Target hours (0-60) |

### Lookups (reference data, do not edit)
Contains dropdown lists used by the other tabs.

---

## Weekly Workflow

1. **Friday**: Team fills updates in Power App or directly in the Excel file (~2 min each)
2. **Monday**: PM sets allocations for upcoming weeks
3. **Monday**: Run `python src/main.py` to generate charts
4. **Monday**: Copilot Agent drafts email, PM attaches the 2 chart PNGs, sends to leadership
