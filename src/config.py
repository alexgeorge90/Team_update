import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

PROJECT_ROOT = Path(__file__).resolve().parent.parent

DATA_FILE = PROJECT_ROOT / os.getenv("DATA_FILE", "sample_data/team_data.xlsx")

WEEKLY_CAPACITY_HOURS = int(os.getenv("WEEKLY_CAPACITY_HOURS", "40"))
TEAM_MEMBERS = [m.strip() for m in os.getenv("TEAM_MEMBERS", "Alex,Srikanth,Enith").split(",")]

PAST_WEEKS = int(os.getenv("PAST_WEEKS", "8"))
FUTURE_WEEKS = int(os.getenv("FUTURE_WEEKS", "4"))

OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
