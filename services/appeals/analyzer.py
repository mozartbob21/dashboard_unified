import re
from datetime import datetime


EMOTION_DEFINITIONS = {
    "angry": {
        "label": "Возмущение / раздражение",
        "keywords": [
            "возмущен",
            "возмущена",
            "возмущены",
            "возмущение",
            "безобразие",
            "ужас",
            "кошмар",
            "надоело",
            "сколько можно",
            "жалоба",
            "жалуемся",
            "не работает",
            "никто не реагирует",
            "бездействие",
            "прошу принять меры",
            "требую",
            "отписки",
            "хамство",
            "невозможно",
            "отвратительно",
        ],
        "weight": 1.0,
    },
    "anxious": {
        "label": "Тревога / обеспокоенность",
        "keywords": [
            "опасно",
            "страшно",
            "угроза",
            "угрожает",
            "риск",
            "аварийный",
            "аварийная",
            "пожар",
            "задымление",
            "дети",
            "ребенок",
            "ребёнок",
            "школа",
            "садик",
            "переживаем",
            "обеспокоены",
            "боюсь",
            "опасность",
        ],
        "weight": 1.15,
    },
    "sad": {
        "label": "Беспомощность / разочарование",
        "keywords": [
            "просим помочь",
            "помогите",
            "не знаем куда обращаться",
            "никто не помогает",
            "остались без",
            "тяжело",
            "невозможно жить",
            "много раз обращались",
            "без результата",
        ],
        "weight": 0.85,
    },
    "urgent": {
        "label": "Срочность",
        "keywords": [
            "срочно",
            "незамедлительно",
            "немедленно",
            "в кратчайшие сроки",
            "сегодня",
            "прямо сейчас",
            "оперативно",
            "безотлагательно",
        ],
        "weight": 1.25,
    },
    "thankful": {
        "label": "Благодарность",
        "keywords": [
            "спасибо",
            "благодарю",
            "выражаю благодарность",
            "признательность",
            "благодарность",
        ],
        "weight": 0.35,
    },
    "neutral": {
        "label": "Нейтральное обращение",
        "keywords": [],
        "weight": 0.2,
    },
}


CRITICALITY_KEYWORDS = {
    "maximum": [
        "угроза жизни",
        "угроза здоровью",
        "пожар",
        "обрушение",
        "затопление",
        "авария",
        "газ",
        "электричество искрит",
        "открытый люк",
        "провал",
        "опасно для детей",
    ],
    "high": [
        "антисанитария",
        "мусор",
        "не вывозится",
        "крысы",
        "нет отопления",
        "нет воды",
        "нет горячей воды",
        "сломано",
        "аварийный",
        "опасно",
        "не работает освещение",
        "гололед",
        "гололёд",
    ],
    "medium": [
        "неудобство",
        "нарушение",
        "проблема",
        "требуется ремонт",
        "просьба разобраться",
        "примите меры",
        "проверить",
    ],
}


ADDRESS_PATTERNS = [
    r"((?:г\.?|город)\s+[А-Яа-яЁёA-Za-z\-\s]+,\s*(?:ул\.?|улица|проспект|пр-т|шоссе|проезд|переулок|пер\.?)\s+[А-Яа-яЁёA-Za-z0-9\-\s]+,\s*(?:д\.?|дом)\s*\d+[А-Яа-яA-Za-z0-9\-\/]*)",
    r"((?:ул\.?|улица|проспект|пр-т|шоссе|проезд|переулок|пер\.?)\s+[А-Яа-яЁёA-Za-z0-9\-\s]+,\s*(?:д\.?|дом)\s*\d+[А-Яа-яA-Za-z0-9\-\/]*)",
    r"((?:мкр\.?|микрорайон)\s+[А-Яа-яЁёA-Za-z0-9\-\s]+,\s*(?:д\.?|дом)\s*\d+[А-Яа-яA-Za-z0-9\-\/]*)",
]


def normalize_text(text):
    text = str(text or "").replace("\r", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def find_keywords(text, keywords):
    lowered = str(text or "").lower()
    found = []

    for keyword in keywords:
        if keyword.lower() in lowered:
            found.append(keyword)

    return found


def detect_emotion(text):
    text = normalize_text(text)
    scores = {}

    for emotion_code, definition in EMOTION_DEFINITIONS.items():
        if emotion_code == "neutral":
            continue

        found = find_keywords(text, definition["keywords"])

        if found:
            raw_score = min(1.0, len(found) / 5)
            scores[emotion_code] = {
                "score": round(raw_score * definition["weight"], 3),
                "found": found,
                "label": definition["label"],
            }

    if not scores:
        return {
            "emotion_class": "neutral",
            "emotion_label": EMOTION_DEFINITIONS["neutral"]["label"],
            "emotion_score": 0.2,
            "matched_words": [],
        }

    emotion_class = max(scores.items(), key=lambda item: item[1]["score"])[0]
    data = scores[emotion_class]

    return {
        "emotion_class": emotion_class,
        "emotion_label": data["label"],
        "emotion_score": min(1.0, data["score"]),
        "matched_words": data["found"],
    }


def extract_duration_days(text):
    lowered = normalize_text(text).lower()
    max_days = 0

    for match in re.finditer(r"(\d+)\s*(?:день|дня|дней|сутки|суток)", lowered):
        max_days = max(max_days, int(match.group(1)))

    for match in re.finditer(r"(\d+)\s*(?:неделя|недели|недель)", lowered):
        max_days = max(max_days, int(match.group(1)) * 7)

    for match in re.finditer(r"(\d+)\s*(?:месяц|месяца|месяцев)", lowered):
        max_days = max(max_days, int(match.group(1)) * 30)

    return max_days


def extract_address(text):
    text = normalize_text(text)

    for pattern in ADDRESS_PATTERNS:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            return normalize_text(match.group(1))

    return ""


def detect_collective(text):
    lowered = normalize_text(text).lower()

    collective_markers = [
        "мы, жители",
        "жители дома",
        "жители района",
        "коллективное",
        "от лица жителей",
        "просим принять меры",
        "подписи",
        "инициативная группа",
    ]

    is_collective = any(marker in lowered for marker in collective_markers)

    n_signers = 1
    signer_match = re.search(r"(\d+)\s*(?:подписей|подписи|подписантов|жителей)", lowered)

    if signer_match:
        n_signers = max(1, int(signer_match.group(1)))
        is_collective = n_signers > 1

    return is_collective, n_signers


def detect_criticality(text):
    lowered = normalize_text(text).lower()

    maximum_found = find_keywords(lowered, CRITICALITY_KEYWORDS["maximum"])
    high_found = find_keywords(lowered, CRITICALITY_KEYWORDS["high"])
    medium_found = find_keywords(lowered, CRITICALITY_KEYWORDS["medium"])

    if maximum_found:
        return {
            "criticality_level": "МАКСИМАЛЬНЫЙ",
            "criticality_label": "Максимальная",
            "criticality_score": 1.0,
            "criticality_reasons": maximum_found,
        }

    if high_found:
        return {
            "criticality_level": "ВЫСОКИЙ",
            "criticality_label": "Высокая",
            "criticality_score": 0.78,
            "criticality_reasons": high_found,
        }

    if medium_found:
        return {
            "criticality_level": "СРЕДНИЙ",
            "criticality_label": "Средняя",
            "criticality_score": 0.48,
            "criticality_reasons": medium_found,
        }

    return {
        "criticality_level": "НИЗКИЙ",
        "criticality_label": "Низкая",
        "criticality_score": 0.2,
        "criticality_reasons": [],
    }


def calculate_index(emotion_score, criticality_score, duration_days, is_collective):
    duration_score = min(1.0, float(duration_days or 0) / 30.0)
    collective_score = 1.0 if is_collective else 0.0

    index = (
        0.25 * float(emotion_score or 0) +
        0.45 * float(criticality_score or 0) +
        0.20 * duration_score +
        0.10 * collective_score
    )

    return round(max(0.0, min(1.0, index)), 3)


def get_priority(index_value):
    try:
        value = float(index_value)
    except Exception:
        value = 0.0

    if value >= 0.85:
        return "urgent", "Срочно"

    if value >= 0.70:
        return "high", "Высокий приоритет"

    if value >= 0.40:
        return "medium", "Средний приоритет"

    return "planned", "Плановый"


def analyze_appeal(subject, text):
    subject = subject or ""
    text = text or ""
    full_text = f"{subject}\n{text}"

    emotion = detect_emotion(full_text)
    criticality = detect_criticality(full_text)
    duration_days = extract_duration_days(full_text)
    address = extract_address(full_text)
    is_collective, n_signers = detect_collective(full_text)

    final_index = calculate_index(
        emotion_score=emotion["emotion_score"],
        criticality_score=criticality["criticality_score"],
        duration_days=duration_days,
        is_collective=is_collective,
    )

    priority_level, priority_label = get_priority(final_index)

    return {
        "created_by": "rule_based_analyzer",
        "analyzed_at": datetime.now().isoformat(timespec="seconds"),

        "emotion_class": emotion["emotion_class"],
        "emotion_label": emotion["emotion_label"],
        "emotion_score": emotion["emotion_score"],
        "matched_words": emotion["matched_words"],

        "criticality_level": criticality["criticality_level"],
        "criticality_label": criticality["criticality_label"],
        "criticality_score": criticality["criticality_score"],
        "criticality_reasons": criticality["criticality_reasons"],

        "duration_days": duration_days,
        "duration_score": round(min(1.0, float(duration_days or 0) / 30.0), 3),

        "address": address,
        "isCollective": is_collective,
        "n_signers": n_signers,
        "collective_score": 1.0 if is_collective else 0.0,

        "index_final": final_index,
        "priority_level": priority_level,
        "priority_label": priority_label,
    }
