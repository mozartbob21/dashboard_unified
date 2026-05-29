import re
from datetime import datetime


EMOTION_DEFINITIONS = {
    # Важно:
    # Это не медицинская/психологическая диагностика человека.
    # Это прикладная оценка эмоциональной тональности обращения
    # для приоритизации и выбора корректной служебной реакции.

    "angry": {
        "label": "Возмущение / раздражение",
        "description": "Признаки недовольства, обвинения, фрустрации из-за бездействия служб.",
        "base_score": 0.0,
        "weight": 1.0,
        "patterns": [
            # Сильное эмоциональное недовольство
            {"pattern": r"\bвозмущ[её]н\w*\b", "weight": 0.32, "marker": "возмущение"},
            {"pattern": r"\bбезобрази[ея]\b", "weight": 0.35, "marker": "безобразие"},
            {"pattern": r"\bбеспредел\b", "weight": 0.40, "marker": "беспредел"},
            {"pattern": r"\bужас\b|\bужасно\b|\bужасный\b", "weight": 0.28, "marker": "ужас"},
            {"pattern": r"\bкошмар\b|\bкошмарн\w*\b", "weight": 0.30, "marker": "кошмар"},
            {"pattern": r"\bотвратительн\w*\b", "weight": 0.30, "marker": "отвратительно"},
            {"pattern": r"\bневозможно\b", "weight": 0.22, "marker": "невозможно"},
            {"pattern": r"\bнадоело\b", "weight": 0.35, "marker": "надоело"},
            {"pattern": r"сколько\s+можно", "weight": 0.42, "marker": "сколько можно"},
            {"pattern": r"сил\s+больше\s+нет", "weight": 0.38, "marker": "сил больше нет"},

            # Претензии к бездействию
            {"pattern": r"никто\s+не\s+реагирует", "weight": 0.42, "marker": "никто не реагирует"},
            {"pattern": r"никто\s+не\s+принимает\s+мер", "weight": 0.42, "marker": "никто не принимает мер"},
            {"pattern": r"\bбездействи[ея]\b", "weight": 0.38, "marker": "бездействие"},
            {"pattern": r"меры\s+не\s+принимаются", "weight": 0.34, "marker": "меры не принимаются"},
            {"pattern": r"проблема\s+не\s+решается", "weight": 0.30, "marker": "проблема не решается"},
            {"pattern": r"ничего\s+не\s+меняется", "weight": 0.28, "marker": "ничего не меняется"},
            {"pattern": r"одни\s+отписки", "weight": 0.45, "marker": "одни отписки"},
            {"pattern": r"\bотписк[аи]\w*\b", "weight": 0.32, "marker": "отписки"},

            # Требовательная позиция
            {"pattern": r"\bтребую\b|\bтребуем\b", "weight": 0.35, "marker": "требую"},
            {"pattern": r"прошу\s+принять\s+меры", "weight": 0.18, "marker": "прошу принять меры"},
            {"pattern": r"примите\s+меры", "weight": 0.22, "marker": "примите меры"},
            {"pattern": r"будем\s+обращаться\s+в\s+прокуратуру", "weight": 0.40, "marker": "угроза обращения в прокуратуру"},
            {"pattern": r"обратимся\s+в\s+прокуратуру", "weight": 0.40, "marker": "обращение в прокуратуру"},
            {"pattern": r"жалоб[ауы]\s+в\s+прокуратуру", "weight": 0.42, "marker": "жалоба в прокуратуру"},
            {"pattern": r"вынужден[аы]?\s+обратиться", "weight": 0.25, "marker": "вынужден обратиться"},

            # ЖКХ-контекст, усиливающий раздражение
            {"pattern": r"мусор\s+не\s+вывоз[ия]т", "weight": 0.24, "marker": "мусор не вывозится"},
            {"pattern": r"контейнер\w*\s+переполнен\w*", "weight": 0.22, "marker": "переполненные контейнеры"},
            {"pattern": r"управляющ\w+\s+компани\w+\s+не\s+реагирует", "weight": 0.38, "marker": "УК не реагирует"},
            {"pattern": r"аварийн\w+\s+служб\w+\s+не\s+приехал\w*", "weight": 0.36, "marker": "аварийная служба не приехала"},
        ],
    },

    "anxious": {
        "label": "Тревога / обеспокоенность",
        "description": "Признаки опасения за безопасность, здоровье, детей, имущество или санитарное состояние.",
        "base_score": 0.0,
        "weight": 1.12,
        "patterns": [
            # Прямые маркеры тревоги
            {"pattern": r"\bопасн\w*\b", "weight": 0.38, "marker": "опасно"},
            {"pattern": r"\bстрашно\b|\bбоюсь\b|\bбоимся\b", "weight": 0.35, "marker": "страх"},
            {"pattern": r"\bугроз[ауы]\b|\bугрожает\b", "weight": 0.42, "marker": "угроза"},
            {"pattern": r"\bриск\b|\bриски\b", "weight": 0.24, "marker": "риск"},
            {"pattern": r"\bопасность\b", "weight": 0.36, "marker": "опасность"},
            {"pattern": r"\bобеспокоен\w*\b|\bпереживаем\b|\bпереживаю\b", "weight": 0.28, "marker": "обеспокоенность"},

            # Безопасность детей
            {"pattern": r"опасно\s+для\s+детей", "weight": 0.55, "marker": "опасно для детей"},
            {"pattern": r"\bдети\b|\bреб[её]нок\b|\bребята\b", "weight": 0.18, "marker": "дети"},
            {"pattern": r"рядом\s+(?:школа|детский\s+сад|садик)", "weight": 0.28, "marker": "рядом школа/садик"},
            {"pattern": r"дети\s+могут\s+пострадать", "weight": 0.55, "marker": "дети могут пострадать"},

            # Аварийность / коммунальные риски
            {"pattern": r"\bаварийн\w*\b", "weight": 0.34, "marker": "аварийность"},
            {"pattern": r"\bпожар\b|\bвозгорани[ея]\b", "weight": 0.55, "marker": "пожарный риск"},
            {"pattern": r"\bзадымлени[ея]\b|\bдым\b", "weight": 0.44, "marker": "задымление"},
            {"pattern": r"\bгаз\b|\bзапах\s+газа\b", "weight": 0.60, "marker": "газ"},
            {"pattern": r"электричеств\w*\s+искрит|проводк\w*\s+искрит|искрит", "weight": 0.55, "marker": "искрение"},
            {"pattern": r"открыт\w+\s+люк", "weight": 0.50, "marker": "открытый люк"},
            {"pattern": r"\bпровал\b|\bяма\b|\bпровалил\w*\b", "weight": 0.35, "marker": "провал/яма"},
            {"pattern": r"\bобрушени[ея]\b|\bможет\s+обрушиться\b", "weight": 0.58, "marker": "риск обрушения"},
            {"pattern": r"\bзатоплени[ея]\b|\bтеч[её]т\b|\bпротечк[аи]\b", "weight": 0.34, "marker": "затопление/протечка"},

            # Санитарная тревога
            {"pattern": r"\bкрысы\b|\bкрыса\b|\bмыши\b", "weight": 0.38, "marker": "грызуны"},
            {"pattern": r"\bантисанитари[яи]\b", "weight": 0.38, "marker": "антисанитария"},
            {"pattern": r"неприятн\w+\s+запах", "weight": 0.22, "marker": "неприятный запах"},
            {"pattern": r"бездомн\w+\s+животн\w+", "weight": 0.22, "marker": "бездомные животные"},
        ],
    },

    "sad": {
        "label": "Беспомощность / разочарование",
        "description": "Признаки усталости, потери надежды, длительного отсутствия результата.",
        "base_score": 0.0,
        "weight": 0.95,
        "patterns": [
            {"pattern": r"просим\s+помочь", "weight": 0.32, "marker": "просим помочь"},
            {"pattern": r"\bпомогите\b", "weight": 0.35, "marker": "помогите"},
            {"pattern": r"не\s+знаем\s+куда\s+обращаться", "weight": 0.45, "marker": "не знаем куда обращаться"},
            {"pattern": r"куда\s+еще\s+обращаться", "weight": 0.40, "marker": "куда еще обращаться"},
            {"pattern": r"никто\s+не\s+помогает", "weight": 0.42, "marker": "никто не помогает"},
            {"pattern": r"остал[аи]сь\s+без", "weight": 0.30, "marker": "остались без"},
            {"pattern": r"\bтяжело\b|\bтяж[её]лая\s+ситуация\b", "weight": 0.25, "marker": "тяжело"},
            {"pattern": r"невозможно\s+жить", "weight": 0.42, "marker": "невозможно жить"},
            {"pattern": r"много\s+раз\s+обращал\w*", "weight": 0.38, "marker": "много раз обращались"},
            {"pattern": r"неоднократно\s+обращал\w*", "weight": 0.36, "marker": "неоднократно обращались"},
            {"pattern": r"без\s+результата", "weight": 0.38, "marker": "без результата"},
            {"pattern": r"результата\s+нет", "weight": 0.34, "marker": "результата нет"},
            {"pattern": r"никто\s+не\s+слышит", "weight": 0.36, "marker": "никто не слышит"},
            {"pattern": r"устал[аиы]?\s+обращаться", "weight": 0.36, "marker": "устали обращаться"},
        ],
    },

    "urgent": {
        "label": "Срочность",
        "description": "Явная просьба о немедленном или ускоренном реагировании.",
        "base_score": 0.0,
        "weight": 1.18,
        "patterns": [
            {"pattern": r"\bсрочно\b", "weight": 0.50, "marker": "срочно"},
            {"pattern": r"\bнезамедлительно\b", "weight": 0.55, "marker": "незамедлительно"},
            {"pattern": r"\bнемедленно\b", "weight": 0.55, "marker": "немедленно"},
            {"pattern": r"в\s+кратчайшие\s+сроки", "weight": 0.45, "marker": "в кратчайшие сроки"},
            {"pattern": r"\bсегодня\b", "weight": 0.30, "marker": "сегодня"},
            {"pattern": r"прямо\s+сейчас", "weight": 0.50, "marker": "прямо сейчас"},
            {"pattern": r"\bоперативно\b", "weight": 0.30, "marker": "оперативно"},
            {"pattern": r"\bбезотлагательно\b", "weight": 0.52, "marker": "безотлагательно"},
            {"pattern": r"требует\s+срочного\s+реагирования", "weight": 0.50, "marker": "требует срочного реагирования"},
            {"pattern": r"прошу\s+срочно", "weight": 0.48, "marker": "прошу срочно"},
        ],
    },

    "thankful": {
        "label": "Благодарность",
        "description": "Позитивная обратная связь, признательность, благодарственное обращение.",
        "base_score": 0.0,
        "weight": 0.75,
        "patterns": [
            {"pattern": r"\bспасибо\b", "weight": 0.45, "marker": "спасибо"},
            {"pattern": r"\bблагодарю\b", "weight": 0.50, "marker": "благодарю"},
            {"pattern": r"выражаю\s+благодарность", "weight": 0.70, "marker": "выражаю благодарность"},
            {"pattern": r"\bпризнательн\w*\b", "weight": 0.45, "marker": "признательность"},
            {"pattern": r"\bблагодарность\b", "weight": 0.50, "marker": "благодарность"},
            {"pattern": r"хочу\s+поблагодарить", "weight": 0.65, "marker": "хочу поблагодарить"},
            {"pattern": r"работ[ау]\s+выполнен[аы]?\s+качественно", "weight": 0.45, "marker": "работа выполнена качественно"},
            {"pattern": r"проблем[ау]\s+решили", "weight": 0.38, "marker": "проблему решили"},
        ],
    },

    "neutral": {
        "label": "Нейтральное обращение",
        "description": "Формальное обращение без выраженного эмоционального фона.",
        "base_score": 0.2,
        "weight": 0.2,
        "patterns": [],
    },
}


# Фразы, которые сами по себе не должны превращать обращение в злое.
# Например: "прошу организовать вывоз мусора" — это нормальная деловая просьба.
NEUTRAL_FORMAL_PATTERNS = [
    r"прошу\s+рассмотреть",
    r"прошу\s+организовать",
    r"прошу\s+провести",
    r"прошу\s+проверить",
    r"прошу\s+сообщить",
    r"направляю\s+обращение",
    r"сообщаю\s+о",
    r"информирую\s+о",
]


# Усилители эмоционального сигнала.
INTENSIFIERS = [
    r"\bочень\b",
    r"\bкрайне\b",
    r"\bсовершенно\b",
    r"\bабсолютно\b",
    r"\bпостоянно\b",
    r"\bрегулярно\b",
    r"\bежедневно\b",
    r"\bсистематически\b",
    r"\bдлительное\s+время\b",
    r"\bуже\s+давно\b",
]


# Маркеры повторности обращений.
REPEATED_REQUEST_PATTERNS = [
    r"повторно\s+обраща\w+",
    r"неоднократно\s+обраща\w+",
    r"много\s+раз\s+обраща\w+",
    r"ранее\s+обраща\w+",
    r"обраща\w+\s+уже\s+не\s+первый\s+раз",
]



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
    # дом № 18 по улице Центральная
    # дом 18 по ул. Центральная
    # дома №18 по улице Центральная
    r"""
    (?:
        дом[а]?|д\.
    )
    \s*
    №?
    \s*
    (?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    \s+
    (?:по|на)
    \s+
    (?P<street_type>
        ул\.?|улице|улица|
        проспекте|проспект|пр-т\.?|пр\.?|
        шоссе|
        проезде|проезд|
        переулке|переулок|пер\.?|
        бульваре|бульвар|бул\.?|
        набережной|набережная|наб\.?|
        площади|площадь|пл\.?
    )
    \s+
    (?P<street>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})
    (?=,|\.|;|:|\n|$)
    """,

    # по улице Центральная, дом № 18
    # по ул. Центральная, д. 18
    r"""
    (?:по|на)
    \s+
    (?P<street_type>
        ул\.?|улице|улица|
        проспекте|проспект|пр-т\.?|пр\.?|
        шоссе|
        проезде|проезд|
        переулке|переулок|пер\.?|
        бульваре|бульвар|бул\.?|
        набережной|набережная|наб\.?|
        площади|площадь|пл\.?
    )
    \s+
    (?P<street>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})
    \s*,?\s*
    (?:
        дом|д\.
    )
    \s*
    №?
    \s*
    (?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """,

    # ул. Центральная, д. 18
    # улица Центральная дом 18
    # г. Чехов, ул. Центральная, д. 18
    r"""
    (?:
        (?:
            г\.?|город|пос\.?|поселок|посёлок|д\.?|деревня|село|рп|пгт
        )
        \s+
        [А-Яа-яЁёA-Za-z0-9\-\s]+
        ,\s*
    )?
    (?P<street_type>
        ул\.?|улица|
        проспект|пр-т\.?|пр\.?|
        шоссе|
        проезд|
        переулок|пер\.?|
        бульвар|бул\.?|
        набережная|наб\.?|
        площадь|пл\.?
    )
    \s+
    (?P<street>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})
    \s*,?\s*
    (?:
        дом|д\.
    )
    \s*
    №?
    \s*
    (?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """,

    # мкр. Венюково, д. 5
    # микрорайон Венюково дом 5
    r"""
    (?P<micro_type>мкр\.?|микрорайон)
    \s+
    (?P<micro>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})
    \s*,?\s*
    (?:
        дом|д\.
    )
    \s*
    №?
    \s*
    (?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """,
]


STREET_TYPE_NORMALIZATION = {
    "ул": "ул.",
    "ул.": "ул.",
    "улица": "ул.",
    "улице": "ул.",

    "проспект": "проспект",
    "проспекте": "проспект",
    "пр-т": "проспект",
    "пр-т.": "проспект",
    "пр": "проспект",
    "пр.": "проспект",

    "шоссе": "шоссе",

    "проезд": "проезд",
    "проезде": "проезд",

    "переулок": "переулок",
    "переулке": "переулок",
    "пер": "переулок",
    "пер.": "переулок",

    "бульвар": "бульвар",
    "бульваре": "бульвар",
    "бул": "бульвар",
    "бул.": "бульвар",

    "набережная": "набережная",
    "набережной": "набережная",
    "наб": "набережная",
    "наб.": "набережная",

    "площадь": "площадь",
    "площади": "площадь",
    "пл": "площадь",
    "пл.": "площадь",
}


def normalize_text(text):
    text = str(text or "").replace("\r", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_street_type(value):
    value = normalize_text(value).lower().replace("ё", "е")
    return STREET_TYPE_NORMALIZATION.get(value, value)


def clean_address_piece(value):
    value = normalize_text(value)
    value = value.strip(" ,.;:")

    # Обрезаем хвосты, если регулярка захватила лишние слова.
    value = re.sub(
        r"\s+(обращаюсь|обращаемся|с\s+жалобой|жалобой|прошу|просим|находится|расположен[аоы]?|возле|около|рядом|напротив)\b.*$",
        "",
        value,
        flags=re.IGNORECASE,
    )

    return normalize_text(value).strip(" ,.;:")


def extract_address(text):
    text = normalize_text(text)

    if not text:
        return ""

    candidates = []

    for pattern in ADDRESS_PATTERNS:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE | re.VERBOSE):
            gd = match.groupdict()

            house = clean_address_piece(gd.get("house") or "")

            if gd.get("micro"):
                micro = clean_address_piece(gd.get("micro") or "")
                if micro and house:
                    candidates.append(f"мкр. {micro}, д. {house}")
                continue

            street = clean_address_piece(gd.get("street") or "")
            street_type = normalize_street_type(gd.get("street_type") or "")

            if street and house:
                candidates.append(f"{street_type} {street}, д. {house}")

    # Дополнительный простой fallback именно для кейса:
    # "дом № 18 по улице Центральная"
    if not candidates:
        simple = re.search(
            r"дом[а]?\s*№?\s*(\d+[А-Яа-яA-Za-z0-9\-\/]*)\s+по\s+улице\s+([А-Яа-яЁёA-Za-z0-9\-\s]+?)(?=,|\.|;|:|\n|$)",
            text,
            flags=re.IGNORECASE,
        )
        if simple:
            house = clean_address_piece(simple.group(1))
            street = clean_address_piece(simple.group(2))
            if street and house:
                candidates.append(f"ул. {street}, д. {house}")

    if not candidates:
        return ""

    unique = []
    seen = set()

    for item in candidates:
        item = clean_address_piece(item)
        key = item.lower().replace("ё", "е")

        if key not in seen:
            seen.add(key)
            unique.append(item)

    unique.sort(key=len)
    return unique[0]


def _regex_found(text, pattern):
    return re.search(pattern, text, flags=re.IGNORECASE) is not None


def _count_regex_matches(text, patterns):
    count = 0
    for pattern in patterns:
        if _regex_found(text, pattern):
            count += 1
    return count


def find_keywords(text, keywords):
    """
    Обратная совместимость:
    раньше keywords был списком строк.
    Теперь может быть список строк или список словарей:
    {"pattern": "...", "weight": 0.3, "marker": "..."}
    """
    lowered = str(text or "").lower()
    found = []

    for item in keywords:
        if isinstance(item, dict):
            pattern = item.get("pattern") or ""
            marker = item.get("marker") or pattern
            if pattern and _regex_found(lowered, pattern):
                found.append(marker)
        else:
            keyword = str(item or "").lower()
            if keyword and keyword in lowered:
                found.append(keyword)

    return found


def _score_emotion_by_patterns(text, emotion_code, definition):
    score = float(definition.get("base_score") or 0.0)
    found = []

    for item in definition.get("patterns", []):
        pattern = item.get("pattern") or ""
        weight = float(item.get("weight") or 0.0)
        marker = item.get("marker") or pattern

        if not pattern:
            continue

        if _regex_found(text, pattern):
            score += weight
            found.append(marker)

    # Усиление, если есть повторность обращений.
    repeated_count = _count_regex_matches(text, REPEATED_REQUEST_PATTERNS)
    if repeated_count:
        if emotion_code in {"angry", "sad"}:
            score += min(0.18, repeated_count * 0.09)
            found.append("повторность обращений")

    # Усиление, если есть общие интенсификаторы.
    intensifier_count = _count_regex_matches(text, INTENSIFIERS)
    if intensifier_count and emotion_code in {"angry", "anxious", "sad", "urgent"}:
        score += min(0.12, intensifier_count * 0.04)
        found.append("усиленная эмоциональная лексика")

    # Снижение злости для формально-делового текста.
    # Например: "прошу организовать вывоз мусора" — не обязательно возмущение.
    formal_count = _count_regex_matches(text, NEUTRAL_FORMAL_PATTERNS)
    if formal_count and emotion_code == "angry":
        score -= min(0.16, formal_count * 0.08)

    # Благодарность не должна автоматически становиться негативной из-за слова "проблема".
    if emotion_code in {"angry", "sad", "anxious"}:
        thankful_signals = _score_emotion_by_patterns(
            text,
            "thankful",
            EMOTION_DEFINITIONS["thankful"],
        )
        if thankful_signals["score"] >= 0.45:
            score -= 0.12

    score = max(0.0, min(1.0, score * float(definition.get("weight") or 1.0)))

    return {
        "score": round(score, 3),
        "found": found,
        "label": definition["label"],
    }


def detect_emotion(text):
    """
    Профессионализированная rule-based оценка эмоциональной тональности ЖКХ-обращения.

    Логика:
    - считаем не просто слова, а психолингвистические маркеры;
    - разделяем возмущение, тревогу, беспомощность, срочность и благодарность;
    - учитываем повторность обращений и усилители;
    - формальные просьбы не считаем автоматически злостью.
    """
    normalized = normalize_text(text).lower().replace("ё", "е")

    scores = {}

    for emotion_code, definition in EMOTION_DEFINITIONS.items():
        if emotion_code == "neutral":
            continue

        data = _score_emotion_by_patterns(normalized, emotion_code, definition)

        if data["score"] > 0 and data["found"]:
            scores[emotion_code] = data

    if not scores:
        return {
            "emotion_class": "neutral",
            "emotion_label": EMOTION_DEFINITIONS["neutral"]["label"],
            "emotion_score": 0.2,
            "matched_words": [],
            "emotion_scores": {},
        }

    # Особое правило:
    # Если есть срочность, но одновременно сильная тревога по безопасности,
    # основная эмоция должна быть "Тревога", а срочность останется в деталях.
    if (
        scores.get("anxious", {}).get("score", 0) >= 0.45
        and scores.get("urgent", {}).get("score", 0) >= 0.35
    ):
        scores["anxious"]["score"] = min(1.0, scores["anxious"]["score"] + 0.08)
        scores["anxious"]["found"].append("срочность на фоне риска")

    # Если благодарность явно доминирует, выбираем её.
    if scores.get("thankful", {}).get("score", 0) >= 0.65:
        negative_max = max(
            scores.get("angry", {}).get("score", 0),
            scores.get("anxious", {}).get("score", 0),
            scores.get("sad", {}).get("score", 0),
            scores.get("urgent", {}).get("score", 0),
        )
        if scores["thankful"]["score"] >= negative_max:
            emotion_class = "thankful"
        else:
            emotion_class = max(scores.items(), key=lambda item: item[1]["score"])[0]
    else:
        emotion_class = max(scores.items(), key=lambda item: item[1]["score"])[0]

    data = scores[emotion_class]

    return {
        "emotion_class": emotion_class,
        "emotion_label": data["label"],
        "emotion_score": min(1.0, data["score"]),
        "matched_words": data["found"][:12],
        "emotion_scores": {
            code: {
                "label": value["label"],
                "score": value["score"],
                "found": value["found"][:8],
            }
            for code, value in sorted(
                scores.items(),
                key=lambda item: item[1]["score"],
                reverse=True,
            )
        },
    }


def extract_duration_days(text):
    lowered = normalize_text(text).lower().replace("ё", "е")
    max_days = 0

    word_numbers = {
        "один": 1,
        "одна": 1,
        "одного": 1,
        "одной": 1,

        "два": 2,
        "две": 2,
        "двух": 2,

        "три": 3,
        "трех": 3,
        "трёх": 3,

        "четыре": 4,
        "четырех": 4,
        "четырёх": 4,

        "пять": 5,
        "пяти": 5,

        "шесть": 6,
        "шести": 6,

        "семь": 7,
        "семи": 7,

        "восемь": 8,
        "восьми": 8,

        "девять": 9,
        "девяти": 9,

        "десять": 10,
        "десяти": 10,
    }

    for match in re.finditer(r"(\d+)\s*(?:день|дня|дней|сутки|суток)", lowered):
        max_days = max(max_days, int(match.group(1)))

    for match in re.finditer(r"(\d+)\s*(?:неделя|недели|недель)", lowered):
        max_days = max(max_days, int(match.group(1)) * 7)

    for match in re.finditer(r"(\d+)\s*(?:месяц|месяца|месяцев)", lowered):
        max_days = max(max_days, int(match.group(1)) * 30)

    number_words_pattern = "|".join(word_numbers.keys())

    for match in re.finditer(
        rf"\b({number_words_pattern})\s+(?:день|дня|дней|сутки|суток)\b",
        lowered,
    ):
        max_days = max(max_days, word_numbers[match.group(1)])

    for match in re.finditer(
        rf"\b({number_words_pattern})\s+(?:неделя|недели|недель)\b",
        lowered,
    ):
        max_days = max(max_days, word_numbers[match.group(1)] * 7)

    for match in re.finditer(
        rf"\b({number_words_pattern})\s+(?:месяц|месяца|месяцев)\b",
        lowered,
    ):
        max_days = max(max_days, word_numbers[match.group(1)] * 30)

    fuzzy_patterns = [
        (r"несколько\s+дней", 3),
        (r"несколько\s+недель", 21),
        (r"несколько\s+месяцев", 90),
        (r"давно", 30),
        (r"длительное\s+время", 30),
        (r"уже\s+давно", 30),
        (r"с\s+выходных", 3),
        (r"после\s+выходных", 3),
        (r"каждый\s+день", 7),
        (r"ежедневно", 7),
        (r"постоянно", 14),
        (r"регулярно", 14),
    ]

    for pattern, days in fuzzy_patterns:
        if re.search(pattern, lowered, flags=re.IGNORECASE):
            max_days = max(max_days, days)

    return max_days



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



# Веса расчёта итогового индекса заявки.
# Формула соответствует блоку 13.4:
# emotion_adj -> criticality_adj -> index_raw -> index_final.
WEIGHT_EMOTIONAL = 0.30
WEIGHT_CRITICALITY = 0.30
WEIGHT_DURATION = 0.20
WEIGHT_SYSTEMATICITY = 0.10
WEIGHT_COLLECTIVE = 0.10

FACT_MAX = 5
CRITICAL_COEF = 0.40
SYSTEMATICITY_MAX_DAYS = 90


JKH_FACT_PATTERNS = [
    r"\bмусор\b",
    r"\bотход\w*\b",
    r"\bконтейнер\w*\b",
    r"\bбак\w*\b",
    r"\bплощадк\w*\b",
    r"\bантисанитари\w*\b",
    r"\bкрыс\w*\b",
    r"\bмыш\w*\b",
    r"\bзапах\b",
    r"\bотоплени\w*\b",
    r"\bвод\w*\b",
    r"\bгоряч\w+\s+вод\w*\b",
    r"\bхолодн\w+\s+вод\w*\b",
    r"\bканализаци\w*\b",
    r"\bзатоплени\w*\b",
    r"\bпротечк\w*\b",
    r"\bкрыша\b",
    r"\bподъезд\w*\b",
    r"\bдвор\w*\b",
    r"\bдорог\w*\b",
    r"\bтротуар\w*\b",
    r"\bосвещени\w*\b",
    r"\bфонар\w*\b",
    r"\bлюк\w*\b",
    r"\bгаз\b",
    r"\bпроводк\w*\b",
    r"\bэлектричеств\w*\b",
    r"\bуправляющ\w+\s+компани\w*\b",
    r"\bук\b",
    r"\bжкх\b",
]


REPEAT_FACT_PATTERNS = [
    r"\bповторно\b",
    r"\bнеоднократно\b",
    r"много\s+раз",
    r"несколько\s+раз",
    r"ранее\s+обращал\w*",
    r"уже\s+обращал\w*",
    r"обращал\w+\s+не\s+первый\s+раз",
    r"без\s+результата",
    r"результата\s+нет",
    r"ничего\s+не\s+изменилось",
    r"существенных\s+изменений\s+не\s+произошло",
]


def clamp(value, min_value=0.0, max_value=1.0):
    try:
        value = float(value)
    except Exception:
        value = 0.0

    return max(min_value, min(max_value, value))


def count_pattern_hits(text, patterns):
    lowered = normalize_text(text).lower().replace("ё", "е")
    count = 0

    for pattern in patterns:
        if re.search(pattern, lowered, flags=re.IGNORECASE):
            count += 1

    return count


def estimate_emotion_confidence(emotion):
    """
    conf_text ∈ [0..1] — коэффициент доверия к эмоциональной оценке текста.

    Чем больше маркеров и чем сильнее доминирует один эмоциональный класс,
    тем выше доверие.
    """
    emotion = emotion or {}

    emotion_class = emotion.get("emotion_class") or "neutral"
    matched_words = emotion.get("matched_words") or []
    emotion_scores = emotion.get("emotion_scores") or {}

    if emotion_class == "neutral":
        return 0.55

    confidence = 0.55
    confidence += min(0.30, len(matched_words) * 0.06)

    scored = []

    for _, data in emotion_scores.items():
        try:
            scored.append(float(data.get("score") or 0))
        except Exception:
            pass

    scored.sort(reverse=True)

    if len(scored) >= 2:
        gap = scored[0] - scored[1]

        if gap >= 0.25:
            confidence += 0.12
        elif gap >= 0.15:
            confidence += 0.08
        elif gap <= 0.05:
            confidence -= 0.08
    elif len(scored) == 1:
        confidence += 0.06

    return round(clamp(confidence), 3)


def calculate_fact_score(text, address, criticality_reasons, duration_days):
    """
    fact_score ∈ [0..FACT_MAX].

    Это не проверка истинности обращения, а количество формализуемых фактических опор:
    адрес, длительность, ЖКХ-сущности, критические маркеры, повторность.
    """
    text = normalize_text(text)
    score = 0
    reasons = []

    if address:
        score += 1
        reasons.append("адрес определён")

    if duration_days and duration_days > 0:
        score += 1
        reasons.append("указана длительность проблемы")

    criticality_reasons = criticality_reasons or []

    if criticality_reasons:
        add = min(2, len(criticality_reasons))
        score += add
        reasons.append(f"маркеры критичности: {', '.join(criticality_reasons[:5])}")

    jkh_hits = count_pattern_hits(text, JKH_FACT_PATTERNS)

    if jkh_hits:
        score += 1
        reasons.append("обнаружены ЖКХ-объекты/сущности")

    repeat_hits = count_pattern_hits(text, REPEAT_FACT_PATTERNS)

    if repeat_hits:
        score += 1
        reasons.append("есть признаки повторности/отсутствия результата")

    score = min(FACT_MAX, score)

    return score, reasons


def extract_systematicity_days_from_text(text):
    """
    Proxy для systematicity_days.

    В идеале systematicity_days надо считать по истории обращений:
    разница между первой и последней жалобой по адресу/теме.

    Пока истории нет — используем текстовые признаки повторности.
    """
    lowered = normalize_text(text).lower().replace("ё", "е")

    if re.search(r"неоднократно|много\s+раз|несколько\s+раз", lowered):
        return 30

    if re.search(r"повторно|ранее\s+обращал|уже\s+обращал", lowered):
        return 14

    if re.search(r"без\s+результата|результата\s+нет|ничего\s+не\s+изменилось|существенных\s+изменений\s+не\s+произошло", lowered):
        return 21

    return 0


def calculate_index(
    emotion_score,
    criticality_score,
    duration_days,
    is_collective,
    conf_text=1.0,
    fact_score=0,
    criticality_level="НИЗКИЙ",
    systematicity_days=0,
):
    """
    Расчёт итогового индекса заявки по формуле стандарта.

    1. emotion_adj = emotion_score * conf_text
    2. fact01 = clamp(fact_score / FACT_MAX, 0, 1)
    3. criticality_adj = clamp(CRITICAL_COEF * fact01 + (1 - CRITICAL_COEF) * criticality_score, 0, 1)
    4. index_raw = линейная сумма компонентов
    5. index_final = clamp(index_raw, 0, 1)
    """
    emotion_score = clamp(emotion_score)
    criticality_score = clamp(criticality_score)
    conf_text = clamp(conf_text)

    emotion_adj = clamp(emotion_score * conf_text)

    fact01 = clamp(float(fact_score or 0) / float(FACT_MAX))

    criticality_adj = clamp(
        CRITICAL_COEF * fact01 + (1.0 - CRITICAL_COEF) * criticality_score
    )

    duration_score = clamp(float(duration_days or 0) / 30.0)

    systematicity_days = int(systematicity_days or 0)
    systematicity_score = clamp(
        float(systematicity_days) / float(SYSTEMATICITY_MAX_DAYS)
    )

    collective_score = 1.0 if is_collective else 0.0

    index_raw = (
        WEIGHT_EMOTIONAL * emotion_adj +
        WEIGHT_CRITICALITY * criticality_adj +
        WEIGHT_DURATION * duration_score +
        WEIGHT_SYSTEMATICITY * systematicity_score +
        WEIGHT_COLLECTIVE * collective_score
    )

    index_final = clamp(index_raw)

    # Минимальный порог для опасных обращений.
    if criticality_level == "МАКСИМАЛЬНЫЙ":
        index_final = max(index_final, 0.70)

    if criticality_level == "ВЫСОКИЙ":
        index_final = max(index_final, 0.50)

    index_final = clamp(index_final)

    return {
        "emotion_adj": round(emotion_adj, 3),
        "conf_text": round(conf_text, 3),

        "fact_score": int(fact_score or 0),
        "fact_max": FACT_MAX,
        "fact01": round(fact01, 3),

        "criticality_adj": round(criticality_adj, 3),

        "duration_score": round(duration_score, 3),

        "systematicity_days": systematicity_days,
        "systematicity_score": round(systematicity_score, 3),

        "collective_score": round(collective_score, 3),

        "index_raw": round(index_raw, 3),
        "index_final": round(index_final, 3),

        "weights": {
            "emotional": WEIGHT_EMOTIONAL,
            "criticality": WEIGHT_CRITICALITY,
            "duration": WEIGHT_DURATION,
            "systematicity": WEIGHT_SYSTEMATICITY,
            "collective": WEIGHT_COLLECTIVE,
        },
    }


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
        "emotion_scores": emotion.get("emotion_scores", {}),

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
