from pathlib import Path
import os

BASE_DIR = Path(__file__).resolve().parents[2]

DATA_DIR = BASE_DIR / "data" / "overdue"
DEBUG_DIR = DATA_DIR / "debug"
NETWORK_DIR = DATA_DIR / "network"
RESPONSES_DIR = DEBUG_DIR / "responses"

PARSED_SUMMARY_FILE = DATA_DIR / "parsed_summary.json"
REPORT_FILE = DATA_DIR / "report.txt"
DATA_FILE = DATA_DIR / "final_result.json"

DASHBOARD_URL = os.getenv(
    "DASHBOARD_URL",
    "https://datalens.yandex/hdxgldxnx8ui1?state=f6c37bf4134",
)

PLAYWRIGHT_PROFILE_DIR = Path(
    os.getenv("PLAYWRIGHT_PROFILE_DIR", str(DATA_DIR / "browser-profile"))
)

DEBUG_ENV_DIR = Path(
    os.getenv("DEBUG_DIR", str(DEBUG_DIR))
)

HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
MAX_WAIT_SECONDS = int(os.getenv("MAX_WAIT_SECONDS", "180"))
IDLE_AFTER_DATA_SECONDS = int(os.getenv("IDLE_AFTER_DATA_SECONDS", "8"))