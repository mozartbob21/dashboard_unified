import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

CDS_BASE_URL = os.getenv("CDS_BASE_URL", "http://176.100.216.181:63871/CDS/ru/")
CDS_HTTP_USER = os.getenv("CDS_HTTP_USER", "")
CDS_HTTP_PASSWORD = os.getenv("CDS_HTTP_PASSWORD", "")
CDS_1C_USER = os.getenv("CDS_1C_USER", "")
CDS_1C_PASSWORD = os.getenv("CDS_1C_PASSWORD", "")
SCREENSHOT_DIR = Path("data/cds")
SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)