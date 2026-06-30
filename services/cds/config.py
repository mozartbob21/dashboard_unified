import os
from pathlib import Path
from dotenv import load_dotenv

# Явно указываем путь к .env от корня проекта
_project_root = Path(__file__).resolve().parent.parent.parent
load_dotenv(_project_root / ".env")

CDS_BASE_URL = os.getenv("CDS_BASE_URL", "http://176.100.216.181:63871/CDS/ru/")
CDS_HTTP_USER = os.getenv("CDS_HTTP_USER", "")
CDS_HTTP_PASSWORD = os.getenv("CDS_HTTP_PASSWORD", "")
CDS_1C_USER = os.getenv("CDS_1C_USER", "").strip()
CDS_1C_PASSWORD = os.getenv("CDS_1C_PASSWORD", "").strip()
SCREENSHOT_DIR = Path("data/cds")
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
