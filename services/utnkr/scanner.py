import asyncio
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[2]

BASE_DIR = PROJECT_ROOT
DATA_DIR = PROJECT_ROOT / "data" / "utnkr"
ENV_PATH = PROJECT_ROOT / ".env"

try:
    from dotenv import load_dotenv

    if ENV_PATH.exists():
        load_dotenv(dotenv_path=ENV_PATH, override=True)
        print("[scanner] .env загружен из:", str(ENV_PATH))
    else:
        load_dotenv(override=True)
        print("[scanner] .env в корне проекта не найден, пробую стандартную загрузку")
except Exception as error:
    print("[scanner] Не удалось загрузить .env:", repr(error))


DEBUG_ROWS_PATH = DATA_DIR / "debug_collected_planfix_rows.json"
VIOLATORS_PATH = DATA_DIR / "violators.json"
VIOLATORS_META_PATH = DATA_DIR / "violators_meta.json"

MAX_REASONABLE_OVERDUE_DAYS = int(os.getenv("MAX_REASONABLE_OVERDUE_DAYS", "30000"))

SITE_LOGIN_URL = os.getenv("SITE_LOGIN_URL", "").strip().strip('"').strip("'")
SITE_TABLE_URL = os.getenv("SITE_TABLE_URL", "").strip().strip('"').strip("'")
SITE_USERNAME = os.getenv("SITE_USERNAME", "").strip().strip('"').strip("'")
SITE_PASSWORD = os.getenv("SITE_PASSWORD", "").strip().strip('"').strip("'")

LOGIN_SELECTOR = os.getenv("LOGIN_SELECTOR", "form#lform input:visible").strip().strip('"').strip("'")
PASSWORD_SELECTOR = os.getenv("PASSWORD_SELECTOR", "form#lform input[type='password']:visible").strip().strip('"').strip("'")
SUBMIT_SELECTOR = os.getenv("SUBMIT_SELECTOR", "form#lform a.lf-btn:has-text('Войти')").strip().strip('"').strip("'")
TABLE_SELECTOR = os.getenv("TABLE_SELECTOR", "table").strip().strip('"').strip("'")

HEADLESS = os.getenv("HEADLESS", "false").strip().strip('"').strip("'").lower() in {"1", "true", "yes", "y"}
SLOW_MO_MS = int(os.getenv("SLOW_MO_MS", "300"))
SCAN_SCROLL_STEPS = int(os.getenv("MAX_TABLE_SCROLL_STEPS", "250"))
SCAN_SCROLL_DELAY_MS = int(os.getenv("SCAN_SCROLL_DELAY_MS", "500"))
PLANFIX_WAIT_AFTER_OPEN_MS = int(os.getenv("PLANFIX_WAIT_AFTER_OPEN_MS", "5000"))
PLAYWRIGHT_STORAGE_STATE_RAW = os.getenv("PLAYWRIGHT_STORAGE_STATE", "").strip().strip('"').strip("'")

if PLAYWRIGHT_STORAGE_STATE_RAW:
    PLAYWRIGHT_STORAGE_STATE = str((PROJECT_ROOT / PLAYWRIGHT_STORAGE_STATE_RAW).resolve())
else:
    PLAYWRIGHT_STORAGE_STATE = str(DATA_DIR / "storage_state.json")

print("[scanner] SITE_LOGIN_URL =", repr(SITE_LOGIN_URL))
print("[scanner] SITE_TABLE_URL =", repr(SITE_TABLE_URL))
print("[scanner] SITE_USERNAME =", repr(SITE_USERNAME))
print("[scanner] HEADLESS =", HEADLESS)
print("[scanner] SLOW_MO_MS =", SLOW_MO_MS)
print("[scanner] LOGIN_SELECTOR =", repr(LOGIN_SELECTOR))
print("[scanner] PASSWORD_SELECTOR =", repr(PASSWORD_SELECTOR))
print("[scanner] SUBMIT_SELECTOR =", repr(SUBMIT_SELECTOR))
print("[scanner] TABLE_SELECTOR =", repr(TABLE_SELECTOR))

STATUS: dict[str, Any] = {
    "running": False,
    "last_started_at": None,
    "last_finished_at": None,
    "last_error": None,
    "last_rows_count": 0,
    "last_violators_count": 0,
}

STOP_SECTION_MARKERS = [
    "отклонения от плановой динамики",
    "cравнение динамики",
    "сравнение динамики",
    "ход выполнения смр",
]

RELEVANT_OVERDUE_SECTION_MARKERS = [
    "просрочена дата начала работ по гпр",
    "просрочена дата окончания работ по гпр",
]

UTNKR_COLOR_MAP = {
    "red": {
        "code": "red",
        "ui": "critical",
        "name": "Красный",
        "priority": 5,
    },
    "yellow": {
        "code": "yellow",
        "ui": "warning",
        "name": "Желтый",
        "priority": 4,
    },
    "green": {
        "code": "green",
        "ui": "ok",
        "name": "Зеленый",
        "priority": 3,
    },
    "blue": {
        "code": "blue",
        "ui": "info",
        "name": "Голубой",
        "priority": 2,
    },
    "violet": {
        "code": "violet",
        "ui": "violet",
        "name": "Фиолетовый",
        "priority": 1,
    },
    "white": {
        "code": "white",
        "ui": "neutral",
        "name": "Белый",
        "priority": 0,
    },
}

UTNKR_RULES = [
    {
        "code": "red",
        "label": "Грубое нарушение технологии производства работ",
        "markers": [
            "грубое нарушение технологии производства работ",
        ],
    },
    {
        "code": "red",
        "label": "Отсутствие рабочих на объекте",
        "markers": [
            "отсутствие рабочих на объекте",
        ],
    },
    {
        "code": "red",
        "label": "Работы приостановлены",
        "markers": [
            "работы приостановлены",
            "работы приостановлена",
            "приостановлены работы",
        ],
    },
    {
        "code": "red",
        "label": "Срыв сроков",
        "markers": [
            "срыв сроков",
            "срыв срока",
        ],
    },
    {
        "code": "yellow",
        "label": "Нарушения технологии производства работ",
        "markers": [
            "нарушения технологии производства работ",
            "нарушение технологии производства работ",
        ],
    },
    {
        "code": "yellow",
        "label": "Недостаточное количество рабочих",
        "markers": [
            "недостаточное количество рабочих",
        ],
    },
    {
        "code": "yellow",
        "label": "Не начаты работы по ГПР более 7 дней",
        "markers": [
            "не начаты работы по гпр более 7 дней",
            "не начаты работы по гпр",
        ],
    },
    {
        "code": "yellow",
        "label": "Отсутствие ГПР",
        "markers": [
            "отсутствие гпр",
        ],
    },
    {
        "code": "yellow",
        "label": "Отсутствие ИД на выполняемые виды работ",
        "markers": [
            "отсутствие ид на выполняемые виды работ",
            "отсутствие ид",
        ],
    },
    {
        "code": "yellow",
        "label": "Выполнение работ без РД",
        "markers": [
            "выполнение работ без рд",
            "без рд",
        ],
    },
    {
        "code": "yellow",
        "label": "Отсутствие ОЖР",
        "markers": [
            "отсутствие ожр",
        ],
    },
    {
        "code": "yellow",
        "label": "Применение материалов с отступлением от ПСД",
        "markers": [
            "применение материалов с отступлением от псд",
        ],
    },
    {
        "code": "yellow",
        "label": "Систематическая несвоевременная уборка и вывоз строительного мусора",
        "markers": [
            "систематическая несвоевременная уборка и вывоз строительного мусора",
        ],
    },
    {
        "code": "yellow",
        "label": "Систематическое нарушение правил складирования материалов",
        "markers": [
            "систематическое нарушение правил складирования материалов",
        ],
    },
    {
        "code": "yellow",
        "label": "Систематическое неустранение выявленных нарушений",
        "markers": [
            "систематическое неустранение выявленных нарушений",
        ],
    },
    {
        "code": "yellow",
        "label": "Частично функционирует",
        "markers": [
            "частично функционирует",
        ],
    },
    {
        "code": "yellow",
        "label": "ГПР без разбивки",
        "markers": [
            "гпр без разбивки",
        ],
    },
    {
        "code": "yellow",
        "label": "Выполнение работ с отступлением от ПСД",
        "markers": [
            "выполнение работ с отступлением от псд",
        ],
    },
    {
        "code": "yellow",
        "label": "ГПР без объемов",
        "markers": [
            "гпр без объемов",
        ],
    },
    {
        "code": "green",
        "label": "Проблематика отсутствует",
        "markers": [
            "проблематика отсутствует",
        ],
    },
    {
        "code": "green",
        "label": "Несоблюдение техники безопасности при производстве работ",
        "markers": [
            "несоблюдение техники безопасности при производстве работ",
            "несоблюдение техники безопасности",
        ],
    },
    {
        "code": "blue",
        "label": "Объект введен в эксплуатацию с условием доработки",
        "markers": [
            "объект введен в эксплуатацию с условием доработки",
            "объект введён в эксплуатацию с условием доработки",
        ],
    },
    {
        "code": "violet",
        "label": "Объект введен в эксплуатацию",
        "markers": [
            "объект введен в эксплуатацию",
            "объект введён в эксплуатацию",
        ],
    },
    {
        "code": "white",
        "label": "Объект функционирует",
        "markers": [
            "объект функционирует",
        ],
    },
]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def save_json(path: Path, data: Any) -> None:
    ensure_data_dir()
    with path.open("w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


def load_json(path: Path, default: Any = None) -> Any:
    if default is None:
        default = {}

    if not path.exists():
        return default

    try:
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return default


def clean_text(text: Any) -> str:
    if text is None:
        return ""

    text = str(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_text(text: Any) -> str:
    text = clean_text(text).lower()
    text = text.replace("ё", "е")
    text = text.replace("cравнение", "сравнение")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def get_utnkr_color(code: str) -> dict[str, Any]:
    return UTNKR_COLOR_MAP.get(code, UTNKR_COLOR_MAP["red"])


def classify_utnkr_overdue(
    overdue_days: int = 0,
    section: str = "",
) -> dict[str, Any]:
    normalized_section = clean_text(section).lower()

    if overdue_days >= 60:
        return {
            **get_utnkr_color("red"),
            "label": "Отставание от ГПР 60 и более дней",
        }

    if overdue_days >= 30:
        return {
            **get_utnkr_color("red"),
            "label": "Отставание от ГПР 30 до 59 дней",
        }

    if overdue_days >= 14:
        return {
            **get_utnkr_color("red"),
            "label": "Отставание от ГПР 14 до 29 дней",
        }

    if overdue_days >= 7:
        if normalized_section == "start":
            return {
                **get_utnkr_color("yellow"),
                "label": "Не начаты работы по ГПР более 7 дней",
            }

        return {
            **get_utnkr_color("yellow"),
            "label": "Отставание от ГПР 7 до 14 дней",
        }

    if overdue_days > 0:
        return {
            **get_utnkr_color("yellow"),
            "label": f"Просрочка менее 7 дней ({overdue_days} дн.)",
        }

    return {
        **get_utnkr_color("white"),
        "label": "Проблематика отсутствует",
    }


def split_non_empty_lines(text: str) -> list[str]:
    text = clean_text(text)
    if not text:
        return []

    lines: list[str] = []
    for line in text.split("\n"):
        line = clean_text(line)
        if line:
            lines.append(line)

    return lines


def extract_task_id_from_text(text: str) -> str:
    text = clean_text(text)

    patterns = [
        r"/task(?:s)?/(\d+)",
        r"task(?:Id|_id|=|/)(\d+)",
        r"\bID[:\s#№-]*(\d{5,})\b",
        r"\bзадач[аи]\s*№?\s*(\d{5,})\b",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1)

    return ""


def parse_planfix_row_identity(row_text: str) -> dict[str, str]:
    lines = split_non_empty_lines(row_text)

    result = {
        "grbs": "",
        "name": "",
        "municipality": "",
    }

    if len(lines) >= 1:
        result["grbs"] = lines[0]

    if len(lines) >= 2:
        result["name"] = lines[1]

    if len(lines) >= 3:
        third_line = lines[2]

        municipality_match = re.match(
            r"^(?P<municipality>.+?)\s+(?:Заключ[её]н|Не\s+заключ[её]н)\b",
            third_line,
            flags=re.IGNORECASE,
        )

        if municipality_match:
            result["municipality"] = clean_text(municipality_match.group("municipality"))
        else:
            cleaned_third_line = clean_text(third_line)

            service_words = [
                "Заключен",
                "Заключён",
                "Не заключен",
                "Не заключён",
                "ПИР",
                "СМР",
                "Да",
                "Нет",
            ]

            municipality = cleaned_third_line

            for service_word in service_words:
                index = municipality.lower().find(service_word.lower())
                if index > 0:
                    municipality = municipality[:index]
                    break

            result["municipality"] = clean_text(municipality)

    return result


def extract_overdue_relevant_text(text: str) -> str:
    text = clean_text(text)
    if not text:
        return ""

    lines = text.split("\n")
    result_lines: list[str] = []
    inside_relevant_block = False

    for line in lines:
        cleaned_line = clean_text(line)
        normalized_line = normalize_text(cleaned_line)

        if not normalized_line:
            if inside_relevant_block:
                result_lines.append("")
            continue

        if any(marker in normalized_line for marker in RELEVANT_OVERDUE_SECTION_MARKERS):
            inside_relevant_block = True
            result_lines.append(cleaned_line)
            continue

        if inside_relevant_block and any(marker in normalized_line for marker in STOP_SECTION_MARKERS):
            inside_relevant_block = False
            continue

        if inside_relevant_block:
            result_lines.append(cleaned_line)

    return clean_text("\n".join(result_lines))


def is_header_like_row(row: dict[str, Any]) -> bool:
    text = normalize_text(
        row.get("raw_row_text")
        or row.get("text")
        or row.get("situation")
        or ""
    )

    if not text:
        return False

    markers = [
        "грбс (гз)",
        "название",
        "муниципальное образование",
        "ситуация на объекте new",
        "дата начала исполнения контракта",
        "дата окончания исполнения контракта",
    ]

    matched = sum(1 for marker in markers if marker in text)
    return matched >= 3


def contains_overdue_keywords(text: str) -> bool:
    normalized = normalize_text(text)

    overdue_keywords = [
        "просроч",
        "отставан",
        "задержк",
        "нарушение срока",
        "нарушения срока",
        "срыв срока",
        "срыв сроков",
        "срок наруш",
    ]

    return any(keyword in normalized for keyword in overdue_keywords)


def extract_overdue_details_from_situation(text: str) -> list[dict[str, Any]]:
    text = clean_text(text)
    if not text:
        return []

    relevant_text = extract_overdue_relevant_text(text)
    if not relevant_text:
        relevant_text = text

    lines = split_non_empty_lines(relevant_text)
    details: list[dict[str, Any]] = []

    current_section = ""

    section_start_markers = {
        "start": "просрочена дата начала работ по гпр",
        "finish": "просрочена дата окончания работ по гпр",
    }

    for line in lines:
        normalized_line = normalize_text(line)

        if section_start_markers["start"] in normalized_line:
            current_section = "start"
            continue

        if section_start_markers["finish"] in normalized_line:
            current_section = "finish"
            continue

        if any(marker in normalized_line for marker in STOP_SECTION_MARKERS):
            current_section = ""
            continue

        if not current_section:
            continue

        patterns = [
            r"(?P<work>.+?),\s*начало\s+работ\s+.+?,\s*отставани[ея]\s*дн(?:ей|я|\.?)?\s*(?P<days>\d{1,6})",
            r"(?P<work>.+?),\s*начало\s+работ\s+.+?,\s*отставание\s*дней\s*(?P<days>\d{1,6})",
            r"(?P<work>.+?),\s*сг\s*.+?,\s*просрочен[оаы]?\s*дн(?:ей|я|\.?)?\s*(?P<days>\d{1,6})",
            r"(?P<work>.+?),\s*сг\s*.+?,\s*просрочено\s*дней\s*(?P<days>\d{1,6})",
            r"(?P<work>.+?),\s*.+?,\s*отставани[ея]\s*дн(?:ей|я|\.?)?\s*(?P<days>\d{1,6})",
            r"(?P<work>.+?),\s*.+?,\s*просрочен[оаы]?\s*дн(?:ей|я|\.?)?\s*(?P<days>\d{1,6})",
        ]

        matched = False

        for pattern in patterns:
            match = re.search(pattern, line, flags=re.IGNORECASE)
            if not match:
                continue

            matched = True

            work_name = clean_text(match.group("work"))
            try:
                days = int(match.group("days"))
            except ValueError:
                continue

            if days <= 0:
                continue

            if days > MAX_REASONABLE_OVERDUE_DAYS:
                print(
                    "Пропускаю подозрительно большую просрочку:",
                    days,
                    "дней.",
                    "Строка:",
                    line,
                )
                continue

            details.append(
                {
                    "days": days,
                    "work_name": work_name,
                    "fragment": line,
                    "match": match.group(0),
                    "text": line,
                    "section": current_section,
                }
            )
            break

        if matched:
            continue

        fallback_match = re.search(
            r"(?P<days>\d{1,6})\s*дн(?:ей|я|\.?)?",
            line,
            flags=re.IGNORECASE,
        )

        if fallback_match:
            try:
                days = int(fallback_match.group("days"))
            except ValueError:
                continue

            if 0 < days <= MAX_REASONABLE_OVERDUE_DAYS:
                work_name = clean_text(re.split(r",", line, maxsplit=1)[0])

                details.append(
                    {
                        "days": days,
                        "work_name": work_name,
                        "fragment": line,
                        "match": fallback_match.group(0),
                        "text": line,
                        "section": current_section,
                    }
                )

    unique: dict[tuple[int, str, str], dict[str, Any]] = {}
    for item in details:
        key = (
            int(item.get("days", 0)),
            clean_text(item.get("work_name", "")),
            clean_text(item.get("section", "")),
        )
        if key not in unique:
            unique[key] = item

    result = list(unique.values())
    result.sort(key=lambda item: int(item.get("days", 0)), reverse=True)
    return result


def find_best_overdue_in_texts(texts: list[str]) -> dict[str, Any] | None:
    best: dict[str, Any] | None = None

    for text in texts:
        text = clean_text(text)
        if not text:
            continue

        details = extract_overdue_details_from_situation(text)
        if details:
            candidate = details[0]
            if best is None or int(candidate.get("days", 0)) > int(best.get("days", 0)):
                best = candidate

    return best


def extract_all_overdue_details_from_texts(texts: list[str]) -> list[dict[str, Any]]:
    all_details: list[dict[str, Any]] = []

    for text in texts:
        text = clean_text(text)
        if not text:
            continue
        all_details.extend(extract_overdue_details_from_situation(text))

    unique: dict[tuple[int, str, str], dict[str, Any]] = {}
    for item in all_details:
        key = (
            int(item.get("days", 0)),
            clean_text(item.get("work_name", "")),
            clean_text(item.get("fragment", ""))[:120],
        )
        if key not in unique:
            unique[key] = item

    result = list(unique.values())
    result.sort(key=lambda item: int(item.get("days", 0)), reverse=True)
    return result


def extract_classification_texts(situation: str, raw_row_text: str) -> list[str]:
    texts: list[str] = []

    relevant_situation_text = extract_overdue_relevant_text(situation)
    relevant_raw_row_text = extract_overdue_relevant_text(raw_row_text)

    if relevant_situation_text:
        texts.append(relevant_situation_text)

    if relevant_raw_row_text and relevant_raw_row_text != relevant_situation_text:
        texts.append(relevant_raw_row_text)

    if situation:
        texts.append(clean_text(situation))

    if raw_row_text and raw_row_text != situation:
        texts.append(clean_text(raw_row_text))

    unique: list[str] = []
    seen: set[str] = set()

    for text in texts:
        key = normalize_text(text)
        if not key or key in seen:
            continue
        seen.add(key)
        unique.append(text)

    return unique


def detect_utnkr_rules_in_text(problem_text: str) -> list[dict[str, Any]]:
    normalized = normalize_text(problem_text)
    matches: list[dict[str, Any]] = []

    if not normalized:
        return matches

    for rule in UTNKR_RULES:
        for marker in rule.get("markers", []):
            normalized_marker = normalize_text(marker)
            if normalized_marker and normalized_marker in normalized:
                color = get_utnkr_color(rule["code"])
                matches.append(
                    {
                        **color,
                        "label": rule["label"],
                        "marker": marker,
                    }
                )
                break

    unique: dict[tuple[str, str], dict[str, Any]] = {}
    for item in matches:
        key = (clean_text(item.get("code", "")), clean_text(item.get("label", "")))
        if key not in unique:
            unique[key] = item

    return list(unique.values())


def classify_utnkr_traffic_light(
    problem_text: str,
    overdue_days: int = 0,
    section: str = "",
) -> dict[str, Any]:
    rule_matches = detect_utnkr_rules_in_text(problem_text)

    best_match: dict[str, Any] | None = None
    for match in rule_matches:
        if best_match is None or int(match.get("priority", 0)) > int(best_match.get("priority", 0)):
            best_match = match

    overdue_match = None
    if overdue_days > 0:
        overdue_match = classify_utnkr_overdue(
            overdue_days=overdue_days,
            section=section,
        )

    if best_match and overdue_match:
        if int(best_match.get("priority", 0)) >= int(overdue_match.get("priority", 0)):
            return {
                **best_match,
                "source": "rule",
                "all_matches": rule_matches,
            }
        return {
            **overdue_match,
            "source": "overdue",
            "all_matches": rule_matches,
        }

    if best_match:
        return {
            **best_match,
            "source": "rule",
            "all_matches": rule_matches,
        }

    if overdue_match:
        return {
            **overdue_match,
            "source": "overdue",
            "all_matches": rule_matches,
        }

    return {
        **get_utnkr_color("white"),
        "label": "Не классифицировано",
        "source": "none",
        "all_matches": [],
    }


def should_include_row_as_violator(
    classification: dict[str, Any],
    overdue_days: int,
    situation: str,
    raw_row_text: str,
) -> bool:
    if overdue_days > 0:
        return True

    color_code = clean_text(classification.get("code", ""))
    label = clean_text(classification.get("label", ""))

    if color_code in {"red", "yellow", "green", "blue", "violet"}:
        return True

    normalized_combined = normalize_text(" ".join([situation or "", raw_row_text or ""]))

    if label == "Объект функционирует":
        return True

    if "проблематика отсутствует" in normalized_combined:
        return True

    return False


def normalize_collected_row(row: dict[str, Any]) -> dict[str, Any]:
    raw_row_text = clean_text(
        row.get("raw_row_text")
        or row.get("row_text")
        or row.get("text")
        or row.get("situation")
        or ""
    )

    if not raw_row_text:
        cells = row.get("cells") or []
        if isinstance(cells, list):
            raw_row_text = clean_text("\n".join(clean_text(cell) for cell in cells if clean_text(cell)))

    task_id = clean_text(row.get("task_id") or extract_task_id_from_text(raw_row_text))
    identity = parse_planfix_row_identity(raw_row_text)

    grbs = clean_text(row.get("grbs")) or identity.get("grbs", "")
    name = clean_text(row.get("name")) or identity.get("name", "")
    municipality = clean_text(row.get("municipality")) or identity.get("municipality", "")

    if name == grbs and identity.get("name"):
        name = identity["name"]

    if municipality == name and identity.get("municipality"):
        municipality = identity["municipality"]

    situation = clean_text(row.get("situation") or raw_row_text)

    return {
        "task_id": task_id,
        "grbs": grbs,
        "name": name,
        "municipality": municipality,
        "situation": situation,
        "raw_row_text": raw_row_text,
    }


def build_item_title(
    max_overdue_days: int,
    grbs: str,
    municipality: str,
    name: str,
    reason: str,
) -> str:
    parts: list[str] = []

    if max_overdue_days > 0:
        parts.append(f"{max_overdue_days} дн.")

    if grbs:
        parts.append(grbs)
    if municipality:
        parts.append(municipality)
    if name:
        parts.append(name)
    if reason:
        parts.append(reason)

    return " | ".join(parts)


def build_reason_from_overdue(best_overdue_data: dict[str, Any] | None) -> tuple[int, str]:
    if not best_overdue_data:
        return 0, ""

    max_overdue_days = int(best_overdue_data.get("days", 0))
    work_name = clean_text(best_overdue_data.get("work_name", ""))
    section = clean_text(best_overdue_data.get("section", ""))

    section_label = ""
    if section == "start":
        section_label = "просрочка начала работ"
    elif section == "finish":
        section_label = "просрочка окончания работ"

    reason_parts: list[str] = []
    if work_name:
        reason_parts.append(work_name)
    if section_label:
        reason_parts.append(section_label)
    if max_overdue_days > 0:
        reason_parts.append(f"{max_overdue_days} дн.")

    return max_overdue_days, ", ".join(reason_parts)


def analyze_rows_by_traffic_light(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    violators: list[dict[str, Any]] = []

    for row_index, row in enumerate(rows, start=1):
        normalized_row = normalize_collected_row(row)

        if is_header_like_row(normalized_row):
            continue

        task_id = normalized_row.get("task_id", "")
        grbs = normalized_row.get("grbs", "")
        name = normalized_row.get("name", "")
        municipality = normalized_row.get("municipality", "")
        situation = normalized_row.get("situation", "")
        raw_row_text = normalized_row.get("raw_row_text", "")

        texts_for_check = extract_classification_texts(situation, raw_row_text)

        best_overdue_data = find_best_overdue_in_texts(texts_for_check)
        all_overdue_details = extract_all_overdue_details_from_texts(texts_for_check)

        max_overdue_days, overdue_reason = build_reason_from_overdue(best_overdue_data)

        fragment = clean_text(best_overdue_data.get("fragment", "")) if best_overdue_data else ""
        work_name = clean_text(best_overdue_data.get("work_name", "")) if best_overdue_data else ""
        overdue_section = clean_text(best_overdue_data.get("section", "")) if best_overdue_data else ""

        combined_problem_text = "\n".join(texts_for_check)

        classification = classify_utnkr_traffic_light(
            problem_text=combined_problem_text,
            overdue_days=max_overdue_days,
            section=overdue_section,
        )

        if not should_include_row_as_violator(
            classification=classification,
            overdue_days=max_overdue_days,
            situation=situation,
            raw_row_text=raw_row_text,
        ):
            continue

        traffic_light_label = clean_text(classification.get("label", ""))

        overdue_text = ""
        if max_overdue_days > 0:
            overdue_text = f"{max_overdue_days} дней"
            if work_name:
                overdue_text = f"{max_overdue_days} дней — {work_name}"
            elif fragment:
                overdue_text = f"{max_overdue_days} дней — {fragment}"

        reason_parts: list[str] = []

        if traffic_light_label:
            reason_parts.append(traffic_light_label)

        if overdue_reason and overdue_reason not in reason_parts:
            reason_parts.append(overdue_reason)

        reason = " | ".join(reason_parts) if reason_parts else (overdue_reason or traffic_light_label or "Найдена проблематика")

        title = build_item_title(
            max_overdue_days=max_overdue_days,
            grbs=grbs,
            municipality=municipality,
            name=name,
            reason=reason,
        )

        problem = reason or overdue_text or traffic_light_label or "Найдена проблематика"

        violator = {
            "row_number": row_index,
            "task_id": task_id or f"row_{row_index}",
            "column": "Ситуация на объекте NEW (автоматическая)",
            "column_index": 0,
            "value": problem,
            "grbs": grbs,
            "name": name,
            "municipality": municipality,
            "max_overdue_days": max_overdue_days,
            "overdue_days": max_overdue_days,
            "days": max_overdue_days,
            "overdue_details": all_overdue_details,
            "details": all_overdue_details,
            "problem": problem,
            "reason": reason,
            "fragment": fragment,
            "work_name": work_name,
            "raw_row_text": raw_row_text,
            "situation": situation,
            "color": classification["ui"],
            "color_code": classification["code"],
            "color_name": classification["name"],
            "traffic_light_label": classification["label"],
            "classification_source": classification.get("source", ""),
            "classification_matches": classification.get("all_matches", []),
            "priority": max(max_overdue_days, int(classification.get("priority", 0))),
            "title": title,
            "overdue": overdue_text,
            "detected_at": now_str(),
        }

        violators.append(violator)

    violators.sort(
        key=lambda item: (
            int(item.get("priority", 0)),
            int(item.get("max_overdue_days", 0)),
        ),
        reverse=True,
    )
    return violators


def build_scan_result(violators: list[dict[str, Any]], rows_count: int = 0) -> dict[str, Any]:
    return {
        "count": len(violators),
        "updated_at": now_str(),
        "scanned_columns": [
            "ГРБС (ГЗ)",
            "Название",
            "Муниципальное образование",
            "Ситуация на объекте NEW (автоматическая)",
        ],
        "items": violators,
        "rows_count": rows_count,
        "violators_count": len(violators),
        "ok": True,
        "status": "success",
        "violators": violators,
        "results": violators,
        "data": violators,
    }


def save_violators_payload(payload: dict[str, Any]) -> None:
    ensure_data_dir()
    save_json(VIOLATORS_PATH, payload)

    meta = {
        "updated_at": now_str(),
        "count": len(payload.get("items", [])),
        "max_reasonable_overdue_days": MAX_REASONABLE_OVERDUE_DAYS,
    }
    save_json(VIOLATORS_META_PATH, meta)
    print("[scanner] violators сохранён:", VIOLATORS_PATH)
    print("[scanner] violators count:", len(payload.get("items", [])))
    print("[scanner] meta сохранён:", VIOLATORS_META_PATH)


def deduplicate_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[str] = set()

    for row in rows:
        raw_row_text = clean_text(row.get("raw_row_text") or row.get("text") or "")
        task_id = clean_text(row.get("task_id"))

        if task_id:
            key = f"task:{task_id}"
        else:
            key = f"text:{normalize_text(raw_row_text)[:500]}"

        if key in seen:
            continue

        seen.add(key)
        result.append(row)

    return result


def value_from_row_by_key_or_fallback(
    row: dict[str, Any],
    key: str | None,
    fallback_indexes: list[int],
) -> str:
    cells_by_key = row.get("cells_by_key", {}) or {}
    cells_by_index = row.get("cells_by_index", []) or []

    if key and cells_by_key.get(key):
        return clean_text(cells_by_key.get(key, ""))

    for index in fallback_indexes:
        if 0 <= index < len(cells_by_index):
            value = clean_text(cells_by_index[index])
            if value:
                return value

    return ""


def get_all_candidate_texts_from_row(row: dict[str, Any]) -> list[str]:
    candidate_texts = []

    cells_by_key = row.get("cells_by_key", {}) or {}
    cells_by_index = row.get("cells_by_index", []) or []
    row_text = clean_text(row.get("row_text", ""))

    for value in cells_by_key.values():
        value = clean_text(value)
        if value:
            candidate_texts.append(value)

    for value in cells_by_index:
        value = clean_text(value)
        if value:
            candidate_texts.append(value)

    if row_text:
        candidate_texts.append(row_text)

    unique_texts = []
    seen = set()

    for text in candidate_texts:
        key = normalize_text(text)

        if not key or key in seen:
            continue

        seen.add(key)
        unique_texts.append(text)

    return unique_texts


def pick_situation_text_from_row(row: dict[str, Any], situation_key: str | None) -> str:
    situation = value_from_row_by_key_or_fallback(
        row,
        situation_key,
        fallback_indexes=[7, 8, 6, 9, 10, 11, 12, 13, 14, 15],
    )

    candidate_texts = get_all_candidate_texts_from_row(row)

    direct_problem_candidates = []
    for candidate in candidate_texts:
        normalized_candidate = normalize_text(candidate)

        if contains_overdue_keywords(candidate):
            direct_problem_candidates.append(candidate)
            continue

        for rule in UTNKR_RULES:
            if any(normalize_text(marker) in normalized_candidate for marker in rule.get("markers", [])):
                direct_problem_candidates.append(candidate)
                break

    if direct_problem_candidates:
        direct_problem_candidates.sort(
            key=lambda value: (len(value), normalize_text(value)),
            reverse=True,
        )
        return direct_problem_candidates[0]

    if situation and contains_overdue_keywords(situation):
        return situation

    if situation:
        return situation

    return clean_text(row.get("row_text", ""))


def merge_text_value(existing_value: str, incoming_value: str, prefer_longer: bool = False) -> str:
    existing_value = clean_text(existing_value)
    incoming_value = clean_text(incoming_value)

    if not existing_value:
        return incoming_value

    if not incoming_value:
        return existing_value

    existing_has_problem = contains_overdue_keywords(existing_value) or bool(detect_utnkr_rules_in_text(existing_value))
    incoming_has_problem = contains_overdue_keywords(incoming_value) or bool(detect_utnkr_rules_in_text(incoming_value))

    if incoming_has_problem and not existing_has_problem:
        return incoming_value

    if prefer_longer and len(incoming_value) > len(existing_value):
        return incoming_value

    return existing_value


def merge_row_data(existing: dict[str, Any], incoming: dict[str, Any]) -> dict[str, Any]:
    text_fields_prefer_longer = {
        "situation",
        "raw_row_text",
    }

    for key, value in incoming.items():
        if isinstance(value, str):
            existing[key] = merge_text_value(
                str(existing.get(key, "")),
                value,
                prefer_longer=key in text_fields_prefer_longer,
            )
        elif isinstance(value, dict):
            existing.setdefault(key, {})
            existing[key].update(value)
        elif isinstance(value, list):
            existing.setdefault(key, [])

            for item in value:
                if item not in existing[key]:
                    existing[key].append(item)
        elif value is not None and not existing.get(key):
            existing[key] = value

    return existing


async def get_planfix_scroll_state(page: Any) -> dict[str, Any]:
    return await page.evaluate(
        """
        () => {
            function hasRows(element) {
                if (!element || !element.querySelector) {
                    return false;
                }

                return Boolean(
                    element.querySelector(".tr-common-item")
                    || element.querySelector("tr.tr-common-item")
                    || element.querySelector("[class*='tr-taskid-']")
                );
            }

            function findVerticalScroller() {
                const elements = Array.from(document.querySelectorAll("*"));

                const candidates = elements.filter(element => {
                    const style = window.getComputedStyle(element);
                    const overflowY = style.overflowY;

                    return hasRows(element)
                        && element.scrollHeight > element.clientHeight + 80
                        && ["auto", "scroll", "overlay"].includes(overflowY);
                });

                candidates.sort((a, b) => {
                    const aDelta = a.scrollHeight - a.clientHeight;
                    const bDelta = b.scrollHeight - b.clientHeight;
                    return bDelta - aDelta;
                });

                return candidates[0] || document.scrollingElement || document.documentElement;
            }

            function findHorizontalScroller() {
                const elements = Array.from(document.querySelectorAll("*"));

                const candidates = elements.filter(element => {
                    const style = window.getComputedStyle(element);
                    const overflowX = style.overflowX;

                    return hasRows(element)
                        && element.scrollWidth > element.clientWidth + 80
                        && ["auto", "scroll", "overlay"].includes(overflowX);
                });

                candidates.sort((a, b) => {
                    const aDelta = a.scrollWidth - a.clientWidth;
                    const bDelta = b.scrollWidth - b.clientWidth;
                    return bDelta - aDelta;
                });

                return candidates[0] || document.scrollingElement || document.documentElement;
            }

            const verticalScroller = findVerticalScroller();
            const horizontalScroller = findHorizontalScroller();

            const isDocumentVertical = verticalScroller === document.scrollingElement || verticalScroller === document.documentElement;
            const isDocumentHorizontal = horizontalScroller === document.scrollingElement || horizontalScroller === document.documentElement;

            return {
                top: isDocumentVertical ? window.scrollY : verticalScroller.scrollTop,
                left: isDocumentHorizontal ? window.scrollX : horizontalScroller.scrollLeft,
                clientHeight: isDocumentVertical ? window.innerHeight : verticalScroller.clientHeight,
                clientWidth: isDocumentHorizontal ? window.innerWidth : horizontalScroller.clientWidth,
                scrollHeight: isDocumentVertical ? document.documentElement.scrollHeight : verticalScroller.scrollHeight,
                scrollWidth: isDocumentHorizontal ? document.documentElement.scrollWidth : horizontalScroller.scrollWidth
            };
        }
        """
    )


async def set_planfix_scroll(page: Any, top: int | None = None, left: int | None = None) -> None:
    await page.evaluate(
        """
        ({ top, left }) => {
            function hasRows(element) {
                if (!element || !element.querySelector) {
                    return false;
                }

                return Boolean(
                    element.querySelector(".tr-common-item")
                    || element.querySelector("tr.tr-common-item")
                    || element.querySelector("[class*='tr-taskid-']")
                );
            }

            function findVerticalScroller() {
                const elements = Array.from(document.querySelectorAll("*"));

                const candidates = elements.filter(element => {
                    const style = window.getComputedStyle(element);
                    const overflowY = style.overflowY;

                    return hasRows(element)
                        && element.scrollHeight > element.clientHeight + 80
                        && ["auto", "scroll", "overlay"].includes(overflowY);
                });

                candidates.sort((a, b) => {
                    const aDelta = a.scrollHeight - a.clientHeight;
                    const bDelta = b.scrollHeight - b.clientHeight;
                    return bDelta - aDelta;
                });

                return candidates[0] || document.scrollingElement || document.documentElement;
            }

            function findHorizontalScroller() {
                const elements = Array.from(document.querySelectorAll("*"));

                const candidates = elements.filter(element => {
                    const style = window.getComputedStyle(element);
                    const overflowX = style.overflowX;

                    return hasRows(element)
                        && element.scrollWidth > element.clientWidth + 80
                        && ["auto", "scroll", "overlay"].includes(overflowX);
                });

                candidates.sort((a, b) => {
                    const aDelta = a.scrollWidth - a.clientWidth;
                    const bDelta = b.scrollWidth - b.clientWidth;
                    return bDelta - aDelta;
                });

                return candidates[0] || document.scrollingElement || document.documentElement;
            }

            const verticalScroller = findVerticalScroller();
            const horizontalScroller = findHorizontalScroller();

            const isDocumentVertical = verticalScroller === document.scrollingElement || verticalScroller === document.documentElement;
            const isDocumentHorizontal = horizontalScroller === document.scrollingElement || horizontalScroller === document.documentElement;

            if (typeof top === "number") {
                if (isDocumentVertical) {
                    window.scrollTo(window.scrollX, top);
                } else {
                    verticalScroller.scrollTop = top;
                }
            }

            if (typeof left === "number") {
                if (isDocumentHorizontal) {
                    window.scrollTo(left, window.scrollY);
                } else {
                    horizontalScroller.scrollLeft = left;
                }
            }
        }
        """,
        {
            "top": top,
            "left": left,
        },
    )


async def extract_visible_planfix_rows(page: Any) -> dict[str, Any]:
    return await page.evaluate(
        """
        () => {
            function clean(value) {
                return String(value || "")
                    .replace(/\\u00a0/g, " ")
                    .replace(/[ \\t]+/g, " ")
                    .replace(/\\n{3,}/g, "\\n\\n")
                    .trim();
            }

            function norm(value) {
                return clean(value).toLowerCase();
            }

            function getClassText(element) {
                return element ? String(element.getAttribute("class") || "") : "";
            }

            function getQeKey(element) {
                const classText = getClassText(element);
                const match = classText.match(/td-item-qe-([0-9]+)/);

                if (!match) {
                    return null;
                }

                return `qe_${match[1]}`;
            }

            function isProbablyHeaderCell(element) {
                if (!element) {
                    return false;
                }

                if (element.closest(".tr-common-item") || element.closest("[class*='tr-taskid-']")) {
                    return false;
                }

                const text = clean(element.innerText || element.textContent || "");

                if (!text) {
                    return false;
                }

                if (text.length > 280) {
                    return false;
                }

                return true;
            }

            function getTaskId(row, rowIndex) {
                const direct = row.getAttribute("data-taskid")
                    || row.getAttribute("data-task-id")
                    || row.getAttribute("taskid");

                if (direct) {
                    return direct;
                }

                const classText = getClassText(row);
                const match = classText.match(/tr-taskid-([0-9]+)/);

                if (match) {
                    return match[1];
                }

                const taskIdFromChild = row.querySelector("[data-taskid], [data-task-id]");

                if (taskIdFromChild) {
                    return taskIdFromChild.getAttribute("data-taskid")
                        || taskIdFromChild.getAttribute("data-task-id");
                }

                return `visible_${rowIndex}_${clean(row.innerText || "").slice(0, 140)}`;
            }

            function findColumnKeyByNames(columnMap, names) {
                const normalizedNames = names.map(norm);

                for (const [key, title] of Object.entries(columnMap)) {
                    const normalizedTitle = norm(title);

                    if (!normalizedTitle) {
                        continue;
                    }

                    for (const wantedName of normalizedNames) {
                        if (
                            normalizedTitle === wantedName
                            || normalizedTitle.includes(wantedName)
                            || wantedName.includes(normalizedTitle)
                        ) {
                            return key;
                        }
                    }
                }

                return null;
            }

            const columnMap = {};

            const headerCandidates = Array.from(document.querySelectorAll("[class*='td-item-qe-']"));

            for (const element of headerCandidates) {
                const key = getQeKey(element);

                if (!key) {
                    continue;
                }

                if (!isProbablyHeaderCell(element)) {
                    continue;
                }

                const text = clean(element.innerText || element.textContent || "");

                if (!text) {
                    continue;
                }

                if (!columnMap[key] || text.length > columnMap[key].length) {
                    columnMap[key] = text;
                }
            }

            const rowElements = Array.from(
                document.querySelectorAll(".tr-common-item, tr.tr-common-item, [class*='tr-taskid-']")
            );

            const rows = rowElements.map((row, rowIndex) => {
                const cellElementsRaw = Array.from(
                    row.querySelectorAll("[class*='td-item-qe-'], td, .td-item")
                );

                const cellElements = Array.from(new Set(cellElementsRaw));

                const cellsByKey = {};
                const cellsByIndex = [];

                cellElements.forEach((cell, cellIndex) => {
                    const text = clean(cell.innerText || cell.textContent || "");

                    const qeKey = getQeKey(cell);
                    const indexKey = `idx_${cellIndex}`;

                    cellsByIndex.push(text);

                    if (qeKey) {
                        cellsByKey[qeKey] = text;
                    } else {
                        cellsByKey[indexKey] = text;
                    }
                });

                const rowText = clean(row.innerText || row.textContent || "");

                return {
                    task_id: getTaskId(row, rowIndex),
                    row_text: rowText,
                    cells_by_key: cellsByKey,
                    cells_by_index: cellsByIndex
                };
            });

            const grbsKey = findColumnKeyByNames(columnMap, [
                "ГРБС (ГЗ)",
                "ГРБС",
                "ГЗ"
            ]);

            const nameKey = findColumnKeyByNames(columnMap, [
                "Название",
                "Наименование",
                "Наименование объекта"
            ]);

            const municipalityKey = findColumnKeyByNames(columnMap, [
                "Муниципальное образование",
                "Муниципалитет",
                "МО"
            ]);

            const situationKey = findColumnKeyByNames(columnMap, [
                "Ситуация на объекте NEW (автоматическая)",
                "Ситуация на объекте NEW",
                "Ситуация на объекте",
                "автоматическая"
            ]);

            return {
                column_map: columnMap,
                keys: {
                    grbs: grbsKey,
                    name: nameKey,
                    municipality: municipalityKey,
                    situation: situationKey
                },
                rows
            };
        }
        """
    )


async def extract_planfix_table_all_rows(page: Any) -> list[dict[str, Any]]:
    print("Жду строки таблицы Planfix")

    await page.wait_for_selector(
        ".tr-common-item, tr.tr-common-item, [class*='tr-taskid-']",
        timeout=60000,
    )

    await page.wait_for_timeout(3000)

    max_scroll_steps = int(os.getenv("MAX_TABLE_SCROLL_STEPS", "250"))
    horizontal_positions_limit = int(os.getenv("MAX_HORIZONTAL_SCROLL_POSITIONS", "5"))

    collected: dict[str, dict[str, Any]] = {}
    column_debug_snapshots = []

    last_top = -1
    stale_steps = 0

    for step in range(max_scroll_steps):
        scroll_state = await get_planfix_scroll_state(page)

        max_left = max(
            0,
            int(scroll_state.get("scrollWidth", 0)) - int(scroll_state.get("clientWidth", 0)),
        )

        if max_left <= 0:
            horizontal_positions = [0]
        else:
            raw_positions = [
                0,
                int(max_left * 0.25),
                int(max_left * 0.5),
                int(max_left * 0.75),
                max_left,
            ]

            horizontal_positions = []

            for position in raw_positions:
                if position not in horizontal_positions:
                    horizontal_positions.append(position)

            horizontal_positions = horizontal_positions[:horizontal_positions_limit]

        before_count = len(collected)

        for left in horizontal_positions:
            await set_planfix_scroll(page, left=left)
            await page.wait_for_timeout(500)

            visible_data = await extract_visible_planfix_rows(page)

            column_map = visible_data.get("column_map", {}) or {}
            column_keys = visible_data.get("keys", {}) or {}
            rows = visible_data.get("rows", []) or []

            column_debug_snapshots.append(
                {
                    "step": step + 1,
                    "left": left,
                    "column_map": column_map,
                    "keys": column_keys,
                }
            )

            for visible_index, row in enumerate(rows):
                task_id = str(row.get("task_id") or "").strip()
                row_text = clean_text(row.get("row_text", ""))

                grbs = value_from_row_by_key_or_fallback(
                    row,
                    column_keys.get("grbs"),
                    fallback_indexes=[1, 0, 2],
                )

                name = value_from_row_by_key_or_fallback(
                    row,
                    column_keys.get("name"),
                    fallback_indexes=[2, 1, 3],
                )

                municipality = value_from_row_by_key_or_fallback(
                    row,
                    column_keys.get("municipality"),
                    fallback_indexes=[3, 2, 4],
                )

                situation = pick_situation_text_from_row(
                    row,
                    column_keys.get("situation"),
                )

                if not task_id:
                    task_id = normalize_text(
                        f"{grbs}|{name}|{municipality}|{visible_index}|{row_text[:80]}"
                    )

                if not task_id:
                    continue

                incoming = {
                    "task_id": task_id,
                    "grbs": grbs,
                    "name": name,
                    "municipality": municipality,
                    "situation": situation,
                    "raw_row_text": row_text,
                    "detected_column_keys": column_keys,
                }

                if task_id not in collected:
                    collected[task_id] = incoming
                else:
                    collected[task_id] = merge_row_data(collected[task_id], incoming)

        after_count = len(collected)
        new_count = after_count - before_count

        print(
            "Скролл таблицы:",
            f"step={step + 1}/{max_scroll_steps}",
            f"top={int(scroll_state.get('top', 0))}",
            f"rows_total={after_count}",
            f"new={new_count}",
        )

        current_top = int(scroll_state.get("top", 0))
        client_height = int(scroll_state.get("clientHeight", 0))
        scroll_height = int(scroll_state.get("scrollHeight", 0))

        is_bottom = current_top + client_height >= scroll_height - 10

        if is_bottom:
            print("Достигнут низ таблицы")
            break

        if current_top == last_top and new_count == 0:
            stale_steps += 1
        else:
            stale_steps = 0

        if stale_steps >= 5:
            print("Таблица больше не прокручивается или новые строки не появляются")
            break

        last_top = current_top

        next_top = current_top + max(300, int(client_height * 0.75))
        await set_planfix_scroll(page, top=next_top)
        await page.wait_for_timeout(900)

    rows_list = list(collected.values())

    rows_list = [normalize_collected_row(row) for row in rows_list]
    rows_list = [row for row in rows_list if not is_header_like_row(row)]
    rows_list = deduplicate_rows(rows_list)

    print("Всего уникальных строк собрано:", len(rows_list))

    for row in rows_list:
        row.pop("detected_column_keys", None)

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    debug_payload = {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "rows_count": len(rows_list),
        "column_debug_snapshots": column_debug_snapshots[-30:],
        "rows": rows_list,
    }

    with open(DEBUG_ROWS_PATH, "w", encoding="utf-8") as file:
        json.dump(debug_payload, file, ensure_ascii=False, indent=2)

    print("Debug собранных строк сохранён:", DEBUG_ROWS_PATH)

    return rows_list

async def find_first_visible_locator(page, selectors, timeout_per_selector=2000):
    for selector in selectors:
        locator = page.locator(selector).first()

        try:
            await locator.wait_for(state="visible", timeout=timeout_per_selector)
            print(f"[scanner] Найден элемент по селектору: {selector}")
            return locator
        except Exception:
            continue

    return None


async def try_login_planfix(page: Any) -> None:
    if not SITE_LOGIN_URL:
        raise RuntimeError("Не задан SITE_LOGIN_URL в .env")

    if not SITE_USERNAME:
        raise RuntimeError("Не задан SITE_USERNAME в .env")

    if SITE_PASSWORD is None:
        raise RuntimeError("Не задан SITE_PASSWORD в .env")

    print("[scanner] Открываю страницу входа:", SITE_LOGIN_URL)

    response = await page.goto(
        SITE_LOGIN_URL,
        wait_until="domcontentloaded",
        timeout=60000,
    )

    print("[scanner] Текущий URL после открытия login page:", page.url)

    if response:
        print("[scanner] HTTP status login page:", response.status)

    await page.wait_for_timeout(3000)

    try:
        already_logged_in = await page.evaluate(
            """
            () => {
                const hasPasswordInput = Array.from(document.querySelectorAll("input[type='password']")).some(el => {
                    const rect = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    return style.display !== 'none'
                        && style.visibility !== 'hidden'
                        && rect.width > 0
                        && rect.height > 0;
                });

                const hasTaskRows =
                    !!document.querySelector(".tr-common-item") ||
                    !!document.querySelector("tr.tr-common-item") ||
                    !!document.querySelector("[class*='tr-taskid-']");

                const hasAnyTable = !!document.querySelector("table");

                const urlLooksAuthenticated =
                    window.location.href.includes("action=tasks") ||
                    window.location.href.includes("filter=") ||
                    window.location.href.includes("clb");

                return !hasPasswordInput && (hasTaskRows || hasAnyTable || urlLooksAuthenticated);
            }
            """
        )
    except Exception:
        already_logged_in = False

    if already_logged_in:
        print("[scanner] Уже авторизованы, повторный логин не нужен")
        return

    print("[scanner] Автовход не обнаружен, ищу форму логина")
    print("[scanner] LOGIN_SELECTOR =", repr(LOGIN_SELECTOR))
    print("[scanner] PASSWORD_SELECTOR =", repr(PASSWORD_SELECTOR))
    print("[scanner] SUBMIT_SELECTOR =", repr(SUBMIT_SELECTOR))

    login_locator = page.locator(LOGIN_SELECTOR).first
    password_locator = page.locator(PASSWORD_SELECTOR).first

    login_found = False
    password_found = False

    try:
        await login_locator.wait_for(state="visible", timeout=15000)
        login_found = True
    except Exception:
        login_found = False

    try:
        await password_locator.wait_for(state="visible", timeout=15000)
        password_found = True
    except Exception:
        password_found = False

    if not login_found or not password_found:
        try:
            current_html = await page.content()
            current_html = current_html[:4000]
        except Exception:
            current_html = "<не удалось получить HTML>"

        raise RuntimeError(
            "Не удалось найти форму логина, и при этом автологин не подтверждён.\n"
            f"URL: {page.url}\n"
            f"LOGIN_SELECTOR: {LOGIN_SELECTOR}\n"
            f"PASSWORD_SELECTOR: {PASSWORD_SELECTOR}\n"
            f"SUBMIT_SELECTOR: {SUBMIT_SELECTOR}\n"
            f"HTML preview:\n{current_html}"
        )

    print("[scanner] Заполняю логин")
    await login_locator.fill(SITE_USERNAME)

    print("[scanner] Заполняю пароль")
    await password_locator.fill(SITE_PASSWORD)

    submit_locator = page.locator(SUBMIT_SELECTOR).first
    clicked = False

    try:
        await submit_locator.wait_for(state="visible", timeout=10000)
        await submit_locator.click()
        clicked = True
        print("[scanner] Кнопка входа нажата через SUBMIT_SELECTOR")
    except Exception as error:
        print("[scanner] SUBMIT_SELECTOR не сработал:", repr(error))

    if not clicked:
        fallback_buttons = [
            page.locator("a.lf-btn:has-text('Войти')").first,
            page.locator("button:has-text('Войти')").first,
            page.locator("input[type='submit']").first,
            page.locator("a:has-text('Войти')").first,
        ]

        for button in fallback_buttons:
            try:
                await button.wait_for(state="visible", timeout=3000)
                await button.click()
                clicked = True
                print("[scanner] Кнопка входа нажата через fallback")
                break
            except Exception:
                continue

    if not clicked:
        raise RuntimeError("Не удалось нажать кнопку входа")

    await page.wait_for_timeout(5000)

    print("[scanner] URL после клика входа:", page.url)

    try:
        login_success = await page.evaluate(
            """
            () => {
                const hasPasswordInput = Array.from(document.querySelectorAll("input[type='password']")).some(el => {
                    const rect = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    return style.display !== 'none'
                        && style.visibility !== 'hidden'
                        && rect.width > 0
                        && rect.height > 0;
                });

                const hasTaskRows =
                    !!document.querySelector(".tr-common-item") ||
                    !!document.querySelector("tr.tr-common-item") ||
                    !!document.querySelector("[class*='tr-taskid-']");

                const hasAnyTable = !!document.querySelector("table");

                const urlLooksAuthenticated =
                    window.location.href.includes("action=tasks") ||
                    window.location.href.includes("filter=") ||
                    window.location.href.includes("clb");

                return !hasPasswordInput && (hasTaskRows || hasAnyTable || urlLooksAuthenticated);
            }
            """
        )
    except Exception:
        login_success = False

    if not login_success:
        raise RuntimeError(
            "После попытки входа авторизация не подтверждена: форма могла остаться или таблица не появилась"
        )

    print("[scanner] Авторизация подтверждена")


async def open_planfix_page(page: Any, url: str | None = None) -> None:
    table_url = clean_text(url or SITE_TABLE_URL)

    print("[scanner] open_planfix_page()")
    print("[scanner] SITE_LOGIN_URL =", repr(SITE_LOGIN_URL))
    print("[scanner] SITE_TABLE_URL =", repr(table_url))

    if not table_url:
        raise RuntimeError("Не указан SITE_TABLE_URL в .env")

    await try_login_planfix(page)

    if table_url and table_url not in page.url:
        print("[scanner] Перехожу на страницу таблицы:", table_url)

        response = await page.goto(
            table_url,
            wait_until="domcontentloaded",
            timeout=60000,
        )

        print("[scanner] URL страницы таблицы:", page.url)

        if response:
            print("[scanner] HTTP status таблицы:", response.status)

        await page.wait_for_timeout(8000)
    else:
        print("[scanner] Уже на странице таблицы после логина")

    row_selectors = ".tr-common-item, tr.tr-common-item, [class*='tr-taskid-']"

    try:
        print("[scanner] Жду строки таблицы:", row_selectors)
        await page.wait_for_selector(
            row_selectors,
            timeout=60000,
        )
        print("[scanner] Строки таблицы подтверждены")
    except Exception as error:
        print("[scanner] Не удалось дождаться строк таблицы:", repr(error))
        print("[scanner] Делаю fallback: дополнительное ожидание и продолжаю")

        await page.wait_for_timeout(5000)


async def scan_current_page(page: Any) -> dict[str, Any]:
    rows = await extract_planfix_table_all_rows(page)
    violators = analyze_rows_by_traffic_light(rows)
    payload = build_scan_result(violators=violators, rows_count=len(rows))
    save_violators_payload(payload)
    return payload


def load_violators() -> dict[str, Any]:
    data = load_json(VIOLATORS_PATH, default={})
    if isinstance(data, dict):
        return data

    if isinstance(data, list):
        return build_scan_result(data, rows_count=0)

    return build_scan_result([], rows_count=0)


def build_violators_from_debug_rows_file(
    input_path: Path | str = DEBUG_ROWS_PATH,
    output_path: Path | str = VIOLATORS_PATH,
) -> dict[str, Any]:
    input_path = Path(input_path)
    output_path = Path(output_path)

    payload = load_json(input_path, default={})
    rows = payload.get("rows", [])

    if not isinstance(rows, list):
        rows = []

    normalized_rows = [normalize_collected_row(row) for row in rows]
    normalized_rows = [row for row in normalized_rows if not is_header_like_row(row)]
    normalized_rows = deduplicate_rows(normalized_rows)

    violators = analyze_rows_by_traffic_light(normalized_rows)
    result_payload = build_scan_result(violators=violators, rows_count=len(normalized_rows))

    save_json(output_path, result_payload)

    meta = {
        "updated_at": now_str(),
        "source": str(input_path),
        "count": len(violators),
        "max_reasonable_overdue_days": MAX_REASONABLE_OVERDUE_DAYS,
    }
    save_json(VIOLATORS_META_PATH, meta)

    return result_payload


async def scan_planfix(url: str | None = None) -> dict[str, Any]:
    STATUS["running"] = True
    STATUS["last_started_at"] = now_str()
    STATUS["last_error"] = None

    print("[scanner] scan_planfix() старт")
    print("[scanner] url arg =", repr(url))
    print("[scanner] SITE_LOGIN_URL =", repr(SITE_LOGIN_URL))
    print("[scanner] SITE_TABLE_URL =", repr(SITE_TABLE_URL))
    print("[scanner] HEADLESS =", HEADLESS)
    print("[scanner] SLOW_MO_MS =", SLOW_MO_MS)

    try:
        from playwright.async_api import async_playwright
    except Exception as error:
        STATUS["running"] = False
        STATUS["last_error"] = f"Playwright не установлен или не запускается: {error}"
        raise RuntimeError(
            "Playwright не установлен. Установи зависимости:\n"
            "pip install playwright\n"
            "playwright install chromium"
        ) from error

    try:
        async with async_playwright() as playwright:
            print("[scanner] Playwright импортирован, запускаю browser.chromium.launch()")

            browser = await playwright.chromium.launch(
                headless=HEADLESS,
                slow_mo=SLOW_MO_MS,
            )

            print("[scanner] Браузер запущен")

            context_kwargs: dict[str, Any] = {
                "viewport": {
                    "width": 1600,
                    "height": 1000,
                },
                "ignore_https_errors": True,
            }

            storage_state_path = Path(PLAYWRIGHT_STORAGE_STATE)
            if storage_state_path.exists():
                context_kwargs["storage_state"] = str(storage_state_path)
                print("[scanner] Использую storage_state:", str(storage_state_path))

            context = await browser.new_context(**context_kwargs)
            page = await context.new_page()

            print("[scanner] Открываю страницу")
            await open_planfix_page(page, url=url)
            print("[scanner] Страница открыта, начинаю scan_current_page()")

            result_payload = await scan_current_page(page)

            print("[scanner] scan_current_page() завершён")

            try:
                storage_state_path.parent.mkdir(parents=True, exist_ok=True)
                await context.storage_state(path=str(storage_state_path))
                print("[scanner] storage_state сохранён:", str(storage_state_path))
            except Exception as error:
                print("[scanner] Не удалось сохранить storage_state:", repr(error))

            await context.close()
            await browser.close()

            print("[scanner] Браузер закрыт")

        rows = result_payload.get("rows_count", 0)
        items = result_payload.get("items", [])

        STATUS["last_rows_count"] = int(rows or 0)
        STATUS["last_violators_count"] = len(items) if isinstance(items, list) else 0
        STATUS["last_finished_at"] = now_str()
        STATUS["running"] = False

        print("[scanner] scan_planfix() успешно завершён")
        return result_payload

    except Exception as error:
        STATUS["running"] = False
        STATUS["last_error"] = repr(error)
        STATUS["last_finished_at"] = now_str()
        print("[scanner] Ошибка в scan_planfix():", repr(error))
        raise


async def scan_table_by_traffic_light(rows: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    print("[scanner] scan_table_by_traffic_light() вызвана")
    print("[scanner] SITE_TABLE_URL =", repr(SITE_TABLE_URL))
    print("[scanner] HEADLESS =", HEADLESS)

    if rows is not None:
        print("[scanner] Режим: анализ переданных rows")
        normalized_rows = [normalize_collected_row(row) for row in rows]
        normalized_rows = [row for row in normalized_rows if not is_header_like_row(row)]
        normalized_rows = deduplicate_rows(normalized_rows)

        violators = analyze_rows_by_traffic_light(normalized_rows)
        payload = build_scan_result(violators=violators, rows_count=len(normalized_rows))
        save_violators_payload(payload)
        return payload

    if SITE_TABLE_URL:
        print("[scanner] Режим: запуск реального сканирования через Playwright")
        return await scan_planfix()

    print("[scanner] Режим: SITE_TABLE_URL пустой, читаю debug json")
    return build_violators_from_debug_rows_file()


async def run_scan(url: str | None = None) -> dict[str, Any]:
    if url:
        return await scan_planfix(url=url)
    return await scan_table_by_traffic_light()


async def scan(url: str | None = None) -> dict[str, Any]:
    if url:
        return await scan_planfix(url=url)
    return await scan_table_by_traffic_light()


async def start_scan(url: str | None = None) -> dict[str, Any]:
    if url:
        return await scan_planfix(url=url)
    return await scan_table_by_traffic_light()


async def main() -> None:
    if SITE_TABLE_URL:
        result = await scan_planfix()
    else:
        result = build_violators_from_debug_rows_file()

    items = result.get("items", [])
    print(f"Найдено нарушителей: {len(items)}")

    for index, item in enumerate(items[:20], start=1):
        print(
            f"{index}. "
            f"{item.get('max_overdue_days')} дн. | "
            f"{item.get('municipality')} | "
            f"{item.get('name')} | "
            f"{item.get('color_name')} | "
            f"{item.get('traffic_light_label')}"
        )


if __name__ == "__main__":
    asyncio.run(main())