from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = BASE_DIR / "data" / "water_dashboard"
DEBUG_DIR = DATA_DIR / "debug"
SNAPSHOT_FILE = DATA_DIR / "snapshot.json"
PLAYWRIGHT_PROFILE_DIR = DATA_DIR / "playwright_profile"

HEADLESS = False
PAGE_WAIT_SECONDS = 6

SOURCES = [
    {"id": "valves",   "name": "Замена задвижек (ZULUGIS)",    "url": "https://datalens.yandex/e2q0obgt7xsex"},
    {"id": "flush",    "name": "Промывки сетей и РЧВ",         "url": "https://datalens.yandex/j9dqqujx03qa3"},
    {"id": "tasks",    "name": "Просроченные задачи ОМСУ",     "url": "https://datalens.yandex/hdxgldxnx8ui1"},
    {"id": "sys_vs",   "name": "Системные адреса (ВС)",        "url": "https://datalens.yandex/6k9dbjyurmu0q"},
    {"id": "edo_rso",  "name": "Переход РСО на ЭДО",           "url": "https://datalens.yandex/f5wqqij889haz"},
    {"id": "nvos",     "name": "Плата за НВОС",                "url": "https://datalens.yandex/62l1qih3msocq?tab=KD"},
    {"id": "meetings", "name": "Присутствие на совещаниях",    "url": "https://datalens.yandex/5f1g88ft7v1up"},
    {"id": "sys_kr",   "name": "Системные адреса (капремонт)", "url": "https://datalens.yandex/6rxd41nckzkep"},
]