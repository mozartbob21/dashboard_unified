import re
import math
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════
# КОНФИГУРАЦИЯ
# ═══════════════════════════════════════════════════════════════════════

CONFIG = {
    "WEIGHT_EMOTIONAL": 0.30,
    "WEIGHT_CRITICALITY": 0.30,
    "WEIGHT_DURATION": 0.20,
    "WEIGHT_SYSTEMATICITY": 0.10,
    "WEIGHT_COLLECTIVE": 0.10,

    "CONF_TEXT_THRESHOLD": 300,
    "CONF_TEXT_MIN": 0.20,

    "COLLECTIVE_MIN": 0.2,
    "COLLECTIVE_COEF": 15.0,

    "DURATION_COEF": 20.0,
    "SYSTEMATICITY_COEF": 90.0,

    "CRITICAL_COEF": 0.60,
    "W_LOW": 0.25,
    "W_MED": 0.50,
    "W_HIGH": 0.75,
    "W_EXT": 1.00,

    "MAX_PUNCT_EXCLAM": 0.30,
    "MAX_PUNCT_QUEST": 0.15,
    "MAX_CAPS_SCORE": 0.25,

    "REPEATED_MIN_INDEX": 0.70,
}

# ═══════════════════════════════════════════════════════════════════════
# СЛОВАРИ (значительно расширенные)
# ═══════════════════════════════════════════════════════════════════════

PROFANITY_DICT = [
    r"блять", r"\bнах\w*", r"\bпох\w*", r"ху[йяеи]\w*", r"пизд\w*", r"муд\w+", r"гандон\w*",
    r"охренели", r"офигели", r"очумели", r"сволоч\w*", r"твар[ьи]\w*", r"набить морду",
    r"упыр[иь]\w*", r"придурк\w*", r"дебил\w*", r"тупы[ехм]\w*", r"идиот\w*",
    r"урод\w*", r"козл\w*", r"скотин\w*", r"мраз\w*", r"подонк\w*"
]

INTENSIFIERS = [
    "очень", "крайне", "совершенно", "абсолютно", "просто", "полностью", "категорически",
    "критически", "постоянно", "регулярно", "ежедневно", "еженедельно", "систематически",
    "каждый день", "каждый вечер", "каждую неделю", "несколько раз", "снова", "опять",
    "вновь", "уже", "давно", "долго", "никогда", "невозможно", "длительное время",
    "уже давно", "издевательство", "беспредел", "позор", "кошмар", "бардак", "ужас",
    "доколе", "сколько можно", "нет сил", "в бешенстве", "задыхаемся", "нас травят"
]

COMPLAINT_MARKERS = [
    "нет", "не работает", "отсутств", "сломал", "плохо", "ужас", "грязная", "ржавая",
    "вонь", "течет", "безобразие", "отписка", "игнор", "не чинят", "отключили",
    "проблема", "жалоба", "неисправн", "неполадк", "дефект", "поломка", "авария"
]

REPEATED_MARKERS = [
    "уже не первый раз", "не первый раз", "не в первый раз", "повторно",
    "повторное обращение", "повторная жалоба", "снова обращаюсь", "опять обращаюсь",
    "обращаюсь не впервые", "обращаюсь повторно", "пишу повторно", "в который раз",
    "неоднократно", "много раз обращались", "обращались ранее", "обращались неоднократно",
    "уже писали", "писали ранее", "уже жаловались", "ранее жаловались",
    "очередное обращение", "предыдущее обращение", "на прошлое обращение", "ответа так и нет"
]

CRITICALITY_PATTERNS = {
    "MAXIMAL": [
        r"погиб\w*", r"смерть", r"летальн\w*", r"реанимаци\w*", r"сбили реб[её]нка",
        r"дтп с пострадавшими", r"труп\w*", r"ударило током", r"утечка газа",
        r"пожар\w*", r"взрыв\w*", r"искрит", r"электричество\s+искрит", r"проводка\s+искрит",
        r"обрыв лэп", r"отравлени\w*"
    ],
    "HIGH": [
        r"открытый люк", r"провал\w*", r"аварийност\w*", r"перелом\w*", r"госпитализаци\w*",
        r"сильное кровотечение", r"реб[её]нок пострадал", r"провалился в люк",
        r"сбила машина", r"сотрясени\w*", r"нет воды", r"нет хвс", r"нет гвс", r"нет отопления",
        r"затоплени\w*", r"прорыв\w*трубы", r"обрушени\w*"
    ],
    "MEDIUM": [
        r"упал\w*", r"травм\w*", r"ушиб\w*", r"порез\w*", r"ожог\w*", r"опасно\b",
        r"угроза\b", r"слабый напор", r"нет напора", r"ржавая вода", r"грязная вода",
        r"коричневая вода", r"запах канализации", r"воняет", r"черви", r"плесен\w*",
        r"протечк\w*", r"залива\w*", r"течь\b"
    ],
    "LOW": [
        r"может травмироваться", r"опасно ходить", r"риск\b", r"небезопасно"
    ]
}

DANGER_OBJECTS = [r"люк", r"яма", r"провал", r"голол[её]д", r"сосульк\w*", r"провод\w*", r"крыш\w*"]
HARM_CONSEQUENCES = [r"упал", r"травм\w*", r"сломал\w*", r"кровь\w*", r"пострадал\w*", r"разбил\w*"]

# ═══════════════════════════════════════════════════════════════════════
# РАСШИРЕННЫЕ СЛОВАРИ ЭМОЦИЙ (главное изменение!)
# Добавлены: контекстные фразы, n-граммы, синонимы, бытовые выражения
# ═══════════════════════════════════════════════════════════════════════

EMOTION_DICTIONARIES = {
    "ГНЕВ": {
        "keywords": [
            r"беспредел", r"позор", r"бардак", r"набить морду", r"издевательств\w*",
            r"издеваетесь", r"очковтерательство", r"блокировать", r"сговор", r"халатност\w*",
            r"дурить людей", r"достали", r"дурдом", r"маразм", r"анархия",
            r"всем пофиг", r"всем плевать", r"наплевать", r"обнаглели",
            r"совесть\w* нет", r"ворьё", r"воруют", r"разворовали", r"коррупци\w*",
            r"произвол", r"безобразие", r"хамство", r"хамят", r"нахамили",
            r"грубость", r"грубят", r"обманывают", r"враньё", r"врут",
            r"разгильдяйство", r"бездельник\w*", r"лентяи", r"бюрократ\w*",
            r"чиновник\w*", r"паразит\w*", r"кормушк\w*"
        ],
        "phrases": [
            r"куда смотр\w*", r"где ваша совесть", r"имейте совесть",
            r"как вам не стыдно", r"до каких пор", r"сколько терпеть",
            r"доколе", r"когда это кончится", r"хватит нас дурить",
            r"за что мы платим", r"за наши деньги", r"на наши налоги",
            r"требу\w+ немедленно", r"требу\w+ незамедлительно",
            r"буду жаловаться", r"обращусь в прокуратуру", r"подам в суд",
            r"напишу президенту", r"обращусь к депутат\w*",
            r"вы обязаны", r"немедленно устран\w*", r"привлечь к ответственности"
        ],
        "weight": 1.0
    },
    "ОТЧАЯНИЕ": {
        "keywords": [
            r"помогите", r"безысходност\w*", r"беспомощн\w*", r"выживаем", r"выживают",
            r"как жить", r"последняя надежда", r"задыхаемся", r"страдани\w*",
            r"сил нет", r"терпеть", r"невыносим\w*", r"мучаемся", r"мучени\w*",
            r"невмоготу", r"на пределе", r"больше не мог\w*", r"спасите",
            r"умоляю", r"заклинаю", r"ради бога", r"христом богом",
            r"нет выхода", r"тупик", r"отчаяни\w*", r"безнадёжн\w*", r"безнадежн\w*"
        ],
        "phrases": [
            r"не знаем? что делать", r"не знаю что делать", r"куда обращаться",
            r"никто не помогает", r"никому нет дела", r"остались без\b",
            r"не к кому обратиться", r"последняя инстанция", r"крик о помощи",
            r"крик души", r"вся надежда на вас", r"больше не выдерж\w*",
            r"дети мёрзнут", r"дети мерзнут", r"ребёнок болеет", r"ребенок болеет",
            r"пожилая мать", r"старики страдают", r"инвалид\w* страда\w*",
            r"жить невозможно", r"существовать невозможно", r"выживать приходится"
        ],
        "weight": 1.2
    },
    "СТРАХ": {
        "keywords": [
            r"опасно", r"страшно", r"боимся", r"боюсь", r"угроза", r"риск\w*",
            r"отравиться", r"заболеть", r"паник\w*", r"взрыв\w*", r"убь[её]т",
            r"опасно для жизни", r"смертельно", r"тревог\w*", r"жутко",
            r"ужас\w*", r"кошмар\w*", r"в любой момент", r"катастроф\w*"
        ],
        "phrases": [
            r"боимся за детей", r"боимся за ребёнка", r"боимся за ребенка",
            r"страшно выходить", r"страшно ходить", r"опасно выходить",
            r"опасно ходить", r"может рухнуть", r"может обрушиться",
            r"может упасть", r"может убить", r"может покалечить",
            r"дети могут пострадать", r"не дай бог", r"только вопрос времени",
            r"рано или поздно", r"пока не случилось", r"пока кто-то не погиб\w*",
            r"ждёте пока убьёт", r"ждете пока убьет", r"до первого трупа",
            r"а если ребёнок", r"а если ребенок", r"а вдруг"
        ],
        "weight": 1.1
    },
    "ОТВРАЩЕНИЕ": {
        "keywords": [
            r"грязь", r"мерзост\w*", r"вонь", r"вонище", r"черви", r"вонюч\w*",
            r"туалет", r"фекали\w*", r"слизь", r"помойк\w*", r"гнил\w*",
            r"рыжая", r"ржав\w*", r"сероводород", r"тухл\w*", r"зловони\w*",
            r"крыс\w*", r"тараканы", r"клоп\w*", r"блох\w*", r"мыши",
            r"плесен\w*", r"грибок", r"гниль", r"отхож\w*", r"нечистот\w*"
        ],
        "phrases": [
            r"невозможно дышать", r"дышать нечем", r"открыть окно невозможно",
            r"окна не открыть", r"запах стоит", r"воняет постоянно",
            r"как на помойке", r"как на свалке", r"хуже чем в хлеву",
            r"пить невозможно", r"пить эту воду", r"мыться невозможно",
            r"коричневая жижа", r"из крана течёт", r"из крана течет",
            r"вместо воды", r"цвет\w* воды", r"квас из крана", r"кока.?кола из крана",
            r"канализаци\w* в подвале", r"подвал залит", r"фекалии в подвале"
        ],
        "weight": 1.0
    },
    "СКОРБЬ": {
        "keywords": [
            r"грустно", r"плакать", r"печальн\w*", r"умер\w*", r"погиб\w*",
            r"жаль", r"обидно", r"стыдн\w*", r"горечь", r"горько",
            r"слёзы", r"слезы", r"плач\w*", r"рыдать", r"скорб\w*",
            r"утрат\w*", r"потер\w*", r"лишил\w*"
        ],
        "phrases": [
            r"стыдно жить", r"стыдно за город", r"стыдно за страну",
            r"до слёз", r"до слез", r"хочется плакать",
            r"больно смотреть", r"сердце кровью", r"душа болит",
            r"невыносимо больно", r"очень обидно", r"очень жаль",
            r"обидно до глубины", r"какой позор", r"до чего довели"
        ],
        "weight": 0.9
    },
    "РАЗДРАЖЕНИЕ": {
        "keywords": [
            r"возмущен\w*", r"безобразие", r"отписк\w*", r"надоело", r"сколько можно",
            r"игнорируют", r"бездействи\w*", r"закрывают заявки", r"кормят завтраками",
            r"не реагируют", r"шаблонн\w*", r"формальн\w*", r"отмахиваются",
            r"некомпетентн\w*", r"непрофессионал\w*", r"разгильдяйств\w*"
        ],
        "phrases": [
            r"когда уже", r"сколько ещё ждать", r"сколько еще ждать",
            r"долго ещё", r"долго еще", r"как долго можно",
            r"вечно одно и то же", r"ничего не меняется", r"воз и ныне там",
            r"толку ноль", r"результата нет", r"никакого толку",
            r"одни обещания", r"только обещают", r"пустые обещания",
            r"заявку закрыли", r"а ничего не сделали", r"для галочки"
        ],
        "weight": 0.8
    },
    "АПАТИЯ": {
        "keywords": [
            r"всё равно", r"все равно", r"пофиг", r"бесполезн\w*",
            r"очередной год", r"как всегда", r"привыкли", r"смирились",
            r"махнули рукой", r"плюнули", r"забили", r"наплевали"
        ],
        "phrases": [
            r"устали жаловаться", r"устали писать", r"устали обращаться",
            r"не верим", r"не верю", r"вряд ли поможет", r"ничего не изменится",
            r"всё без толку", r"все без толку", r"ни на что не надеемся",
            r"ни на что не надеюсь", r"для проформы", r"пишу для протокола",
            r"ради отчётности", r"ради отчетности", r"всё бесполезно", r"все бесполезно",
            r"очередная отписка", r"опять отпишетесь", r"знаю что бесполезно"
        ],
        "weight": 1.0
    }
}

# ═══════════════════════════════════════════════════════════════════════
# КОНТЕКСТНЫЕ МОДИФИКАТОРЫ ЭМОЦИЙ
# Правила повышения/понижения скора на основе структуры текста
# ═══════════════════════════════════════════════════════════════════════

EMOTION_CONTEXTUAL_BOOSTERS = {
    # Если найдены маркеры адресатов угроз (прокуратура, суд) → ГНЕВ
    "ГНЕВ": [r"прокуратур\w*", r"суд\b", r"надзорн\w*", r"ответственност\w*"],
    # Если упоминаются дети/старики в связке с проблемой → ОТЧАЯНИЕ
    "ОТЧАЯНИЕ": [r"дет[иям]\w*", r"реб[её]нок", r"малыш\w*", r"пожил\w*", r"стари\w*", r"бабушк\w*"],
    # Если есть слова угрозы + конкретный объект → СТРАХ
    "СТРАХ": [r"обрушени\w*", r"рухн\w*", r"трещин\w*", r"аварийн\w*"],
    # Санитарные термины → ОТВРАЩЕНИЕ
    "ОТВРАЩЕНИЕ": [r"санитарн\w*", r"антисанитари\w*", r"дезинфекци\w*", r"дератизаци\w*"],
}

# Тональные маркеры — вычисляются из тона всего текста для уточнения
TONE_MARKERS = {
    "imperative": [  # Императив → ГНЕВ
        r"требую", r"требуем", r"обязаны", r"извольте", r"немедленно",
        r"незамедлительно", r"в кратчайшие", r"срочно устран\w*"
    ],
    "pleading": [  # Просьба/мольба → ОТЧАЯНИЕ
        r"прошу\b", r"просим\b", r"умоляю", r"заклинаю", r"пожалуйста помогите",
        r"очень прошу", r"убедительно прошу", r"слёзно прошу", r"слезно прошу"
    ],
    "resigned": [  # Смирение → АПАТИЯ
        r"уже не жд\w*", r"не надеемся", r"не надеюсь", r"без надежды",
        r"пишу для галочки", r"просто фиксирую", r"для истории"
    ],
    "sarcastic": [  # Сарказм → РАЗДРАЖЕНИЕ / ГНЕВ (зависит от контекста)
        r"спасибо конечно", r"замечательно просто", r"прекрасная работа",
        r"отличный сервис", r"так держать", r"молодцы что сказать",
        r"браво", r"ну и ладно", r"зато", r"как обычно"
    ]
}


# ═══════════════════════════════════════════════════════════════════════
# УТИЛИТЫ
# ═══════════════════════════════════════════════════════════════════════

def clamp01(x):
    return max(0.0, min(1.0, float(x)))


def normalize_text(text):
    if not text:
        return ""
    text = str(text).replace("\r", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ═══════════════════════════════════════════════════════════════════════
# EMOTION SCORE (интенсивность)
# ═══════════════════════════════════════════════════════════════════════

def calculate_emotion_score(text):
    """
    Эвристический расчет интенсивности эмоций (emotion_score).
    Диапазон: [0..1]
    """
    if not text:
        return 0.0

    tl = text.lower()

    # 1. Базовый негатив
    baseline = 0.0
    complaint_hits = sum(1 for marker in COMPLAINT_MARKERS if marker in tl)
    if complaint_hits > 0:
        baseline = min(0.15, 0.05 * complaint_hits)

    # 2. Пунктуационная интенсивность
    exclam = text.count("!")
    quest = text.count("?")
    ellipsis = text.count("...")
    punct_score = (
        min(CONFIG["MAX_PUNCT_EXCLAM"], exclam / 15.0) +
        min(CONFIG["MAX_PUNCT_QUEST"], quest / 8.0) +
        min(0.05, ellipsis * 0.015)
    )

    # 3. Вклад капса
    letters = [c for c in text if c.isalpha()]
    if letters:
        all_letters_count = len(letters)
        upper_letters_count = sum(1 for c in letters if c.isupper())
        caps_ratio = upper_letters_count / all_letters_count
        # Исключаем тексты полностью в верхнем регистре < 10 символов (может быть заголовок)
        if all_letters_count > 10 and caps_ratio > 0.3:
            caps_score = min(CONFIG["MAX_CAPS_SCORE"], 0.50 * caps_ratio)
        else:
            caps_score = min(CONFIG["MAX_CAPS_SCORE"], 0.25 * caps_ratio)
    else:
        caps_score = 0.0

    # 4. Обсценная/агрессивная лексика
    prof_hits = sum(1 for pattern in PROFANITY_DICT if re.search(pattern, tl, flags=re.IGNORECASE))
    prof_score = min(0.30, 0.12 * prof_hits)

    # 5. Усилители и сильные фразы
    inten_hits = sum(1 for word in INTENSIFIERS if word in tl)
    inten_score = min(0.30, 0.035 * inten_hits)

    # 6. Длина и детальность текста (длинные эмоциональные тексты → больше вовлеченности)
    words = tl.split()
    length_bonus = 0.0
    if len(words) > 50:
        length_bonus = min(0.10, (len(words) - 50) * 0.001)

    # 7. Повторения слов/фраз (эмоциональные люди повторяются)
    repetition_bonus = 0.0
    word_freq = {}
    for w in words:
        if len(w) > 4:
            word_freq[w] = word_freq.get(w, 0) + 1
    repeated_words = sum(1 for count in word_freq.values() if count >= 3)
    if repeated_words > 0:
        repetition_bonus = min(0.08, repeated_words * 0.025)

    total_score = (baseline + punct_score + caps_score + prof_score +
                   inten_score + length_bonus + repetition_bonus)
    return clamp01(total_score)


# ═══════════════════════════════════════════════════════════════════════
# EMOTION CLASS (классификация) — ПОЛНОСТЬЮ ПЕРЕРАБОТАННАЯ
# ═══════════════════════════════════════════════════════════════════════

def classify_emotion_class(text):
    """
    Многоуровневая классификация эмоции обращения.
    
    Алгоритм:
    1. Подсчёт совпадений по расширенным словарям (keywords + phrases)
    2. Контекстные бустеры
    3. Тональные маркеры 
    4. Взвешенная агрегация с учетом приоритета более "сильных" эмоций
    5. Fallback на основе общей тональности (не слепой "РАЗДРАЖЕНИЕ")
    """
    if not text:
        return "РАЗДРАЖЕНИЕ"

    tl = text.lower().replace("ё", "е")
    
    # ─── Шаг 1: Подсчет совпадений по словарям ───
    raw_scores = {}
    matched_details = {}  # Для дебага

    for emotion, data in EMOTION_DICTIONARIES.items():
        keywords = data["keywords"]
        phrases = data["phrases"]
        weight = data["weight"]

        keyword_hits = 0
        phrase_hits = 0
        matched_items = []

        for pattern in keywords:
            if re.search(pattern, tl, flags=re.IGNORECASE):
                keyword_hits += 1
                matched_items.append(f"kw:{pattern}")

        for pattern in phrases:
            if re.search(pattern, tl, flags=re.IGNORECASE):
                phrase_hits += 1
                matched_items.append(f"ph:{pattern}")

        # Фразы весят больше (2x), т.к. они контекстнее
        score = (keyword_hits * 1.0 + phrase_hits * 2.0) * weight
        raw_scores[emotion] = score
        matched_details[emotion] = matched_items

    # ─── Шаг 2: Контекстные бустеры ───
    for emotion, boosters in EMOTION_CONTEXTUAL_BOOSTERS.items():
        booster_hits = sum(1 for pat in boosters if re.search(pat, tl))
        if booster_hits > 0 and raw_scores.get(emotion, 0) > 0:
            # Усиливаем только если уже есть базовые совпадения
            raw_scores[emotion] = raw_scores.get(emotion, 0) * (1.0 + 0.25 * booster_hits)

    # ─── Шаг 3: Тональные маркеры ───
    tone_imperative = sum(1 for pat in TONE_MARKERS["imperative"] if re.search(pat, tl))
    tone_pleading = sum(1 for pat in TONE_MARKERS["pleading"] if re.search(pat,tl)) 
    tone_pleading = sum(1 for pat in TONE_MARKERS["pleading"] if re.search(pat, tl))
    tone_resigned = sum(1 for pat in TONE_MARKERS["resigned"] if re.search(pat, tl))
    tone_sarcastic = sum(1 for pat in TONE_MARKERS["sarcastic"] if re.search(pat, tl))

    # Тональные бонусы
    if tone_imperative > 0:
        raw_scores["ГНЕВ"] = raw_scores.get("ГНЕВ", 0) + tone_imperative * 1.5
    if tone_pleading > 0:
        raw_scores["ОТЧАЯНИЕ"] = raw_scores.get("ОТЧАЯНИЕ", 0) + tone_pleading * 1.5
    if tone_resigned > 0:
        raw_scores["АПАТИЯ"] = raw_scores.get("АПАТИЯ", 0) + tone_resigned * 1.5
    if tone_sarcastic > 0:
        # Сарказм может быть и гневом, и раздражением
        raw_scores["РАЗДРАЖЕНИЕ"] = raw_scores.get("РАЗДРАЖЕНИЕ", 0) + tone_sarcastic * 0.8
        raw_scores["ГНЕВ"] = raw_scores.get("ГНЕВ", 0) + tone_sarcastic * 0.5

    # ─── Шаг 4: Обсценная лексика → бустер ГНЕВ ───
    prof_hits = sum(1 for pattern in PROFANITY_DICT if re.search(pattern, tl, flags=re.IGNORECASE))
    if prof_hits > 0:
        raw_scores["ГНЕВ"] = raw_scores.get("ГНЕВ", 0) + prof_hits * 2.0

    # ─── Шаг 5: Нормализация и выбор победителя ───
    total_raw = sum(raw_scores.values())

    if total_raw == 0:
        # ─── УМНЫЙ FALLBACK (не слепой "РАЗДРАЖЕНИЕ") ───
        return _smart_fallback(text, tl)

    # Находим максимум
    sorted_emotions = sorted(raw_scores.items(), key=lambda x: x[1], reverse=True)
    top_emotion = sorted_emotions[0][0]
    top_score = sorted_emotions[0][1]

    # Проверяем, есть ли явный лидер (отрыв от второго места > 30%)
    if len(sorted_emotions) > 1:
        second_score = sorted_emotions[1][1]
        if second_score > 0 and top_score > 0:
            dominance = (top_score - second_score) / top_score
            # Если нет явного лидера и обе эмоции "сильные" — применяем приоритеты
            if dominance < 0.2:
                top_emotion = _resolve_tie(sorted_emotions[0], sorted_emotions[1])

    return top_emotion


def _smart_fallback(text, tl):
    """
    Умный fallback: когда словари не дали совпадений,
    определяем эмоцию по общей тональности текста.
    """
    # Проверяем общую негативность
    negative_markers = [
        "не ", "нет ", "без ", "ни ", "нельзя", "невозможно", "отсутств",
        "проблем", "жалоб", "претензи", "недовол", "неудовлетвор"
    ]
    negative_count = sum(1 for m in negative_markers if m in tl)

    # Проверяем вопросительный тон
    question_count = text.count("?")

    # Проверяем восклицательный тон
    exclaim_count = text.count("!")

    # Проверяем длину — короткие формальные сообщения
    words = tl.split()
    is_short = len(words) < 15

    # Решение
    if exclaim_count >= 3 and negative_count >= 2:
        return "ГНЕВ"
    elif question_count >= 2 and negative_count >= 1:
        return "РАЗДРАЖЕНИЕ"
    elif negative_count >= 3:
        return "РАЗДРАЖЕНИЕ"
    elif negative_count >= 1 and not is_short:
        return "РАЗДРАЖЕНИЕ"
    elif is_short and negative_count == 0:
        return "NEUTRAL"
    else:
        # Если есть хоть какой-то негатив — раздражение, иначе нейтральный
        if negative_count > 0:
            return "РАЗДРАЖЕНИЕ"
        return "NEUTRAL"


def _resolve_tie(first, second):
    """
    Разрешение ничьей между двумя эмоциями.
    Приоритет: ОТЧАЯНИЕ > СТРАХ > ГНЕВ > ОТВРАЩЕНИЕ > СКОРБЬ > РАЗДРАЖЕНИЕ > АПАТИЯ
    """
    priority = {
        "ОТЧАЯНИЕ": 7,
        "СТРАХ": 6,
        "ГНЕВ": 5,
        "ОТВРАЩЕНИЕ": 4,
        "СКОРБЬ": 3,
        "РАЗДРАЖЕНИЕ": 2,
        "АПАТИЯ": 1,
    }
    e1, s1 = first
    e2, s2 = second

    p1 = priority.get(e1, 0)
    p2 = priority.get(e2, 0)

    # При ничьей по скору — выбираем более "тяжёлую" эмоцию
    if p1 >= p2:
        return e1
    return e2


def get_emotion_confidence(text, emotion_class, raw_scores=None):
    """
    Дополнительная метрика: насколько уверенно мы классифицировали эмоцию.
    0.0 — полная неуверенность (fallback), 1.0 — очень уверены.
    """
    if emotion_class == "NEUTRAL":
        return 0.3

    if not text:
        return 0.1

    tl = text.lower().replace("ё", "е")

    # Подсчитаем заново если не переданы
    if raw_scores is None:
        raw_scores = {}
        for emotion, data in EMOTION_DICTIONARIES.items():
            hits = sum(1 for p in data["keywords"] if re.search(p, tl))
            hits += sum(2 for p in data["phrases"] if re.search(p, tl))
            raw_scores[emotion] = hits * data["weight"]

    total = sum(raw_scores.values())
    if total == 0:
        return 0.2

    top_score = raw_scores.get(emotion_class, 0)
    dominance = top_score / total

    # Нормализуем в [0.3, 1.0] — минимум 0.3 если хоть что-то нашли
    return clamp01(0.3 + 0.7 * dominance)


# ═══════════════════════════════════════════════════════════════════════
# SOCIAL CLASS
# ═══════════════════════════════════════════════════════════════════════

def extract_social_class(text):
    if not text:
        return "Житель"

    tl = text.lower()
    classes = []

    if any(word in tl for word in ["маломобильн", "коляск", "пандус", "инвалид", "ограниченн"]):
        classes.append("Маломобильные")
    if any(word in tl for word in ["дети", "ребенок", "ребёнок", "новорожден", "садик", "сын", "дочь", "малыш", "младенец"]):
        classes.append("Семьи с детьми")
    if any(word in tl for word in ["пенсионер", "старик", "бабушк", "дедушк", "пожил", "престарел"]):
        classes.append("Пенсионеры")
    if any(word in tl for word in ["многодетн"]):
        classes.append("Многодетные")
    if any(word in tl for word in ["ветеран", "участник войны", "блокадник"]):
        classes.append("Ветераны")
    if any(word in tl for word in ["беременн"]):
        classes.append("Беременные")

    if not classes:
        return "Житель"
    return ", ".join(classes)


# ═══════════════════════════════════════════════════════════════════════
# DURATION (длительность проблемы)
# ═══════════════════════════════════════════════════════════════════════

def extract_duration_days(text):
    if not text:
        return None

    tl = text.lower().replace("ё", "е")

    word_numbers = {
        "один": 1, "одна": 1, "два": 2, "две": 2, "три": 3, "четыре": 4,
        "пять": 5, "шесть": 6, "семь": 7, "восемь": 8, "девять": 9, "десять": 10,
        "полтора": 1.5, "пару": 2
    }

    # Дни
    for m in re.finditer(r"(\d+)\s*(?:день|дня|дней|сутки|суток)", tl):
        return int(m.group(1))
    for m in re.finditer(rf"\b({'|'.join(word_numbers.keys())})\s*(?:день|дня|дней|сутки|суток)", tl):
        return int(word_numbers[m.group(1)])

    # Недели
    for m in re.finditer(r"(\d+)\s*(?:недел\w*)", tl):
        return int(m.group(1)) * 7
    for m in re.finditer(rf"\b({'|'.join(word_numbers.keys())})\s*(?:недел\w*)", tl):
        return int(word_numbers[m.group(1)] * 7)

    # Месяцы
    for m in re.finditer(r"(\d+)\s*(?:месяц\w*)", tl):
        return int(m.group(1)) * 30
    for m in re.finditer(rf"\b({'|'.join(word_numbers.keys())})\s*(?:месяц\w*)", tl):
        return int(word_numbers[m.group(1)] * 30)

    # Годы
    for m in re.finditer(r"(\d+)\s*(?:год\w*|лет)", tl):
        return int(m.group(1)) * 365
    for m in re.finditer(rf"\b({'|'.join(word_numbers.keys())})\s*(?:год\w*|лет)", tl):
        return int(word_numbers[m.group(1)] * 365)

    # Полгода, полтора года
    if "полгода" in tl:
        return 180
    if "полтора года" in tl:
        return 540

    # Эвристические бакеты
    fuzzy_mapping = {
        "каждый день": 7, "ежедневно": 7, "постоянно": 14,
        "уже не первый раз": 14, "не первый день": 7, "не первый месяц": 60,
        "давно": 30, "длительное время": 30, "уже месяц": 30,
        "с начала года": 180, "с прошлого года": 365,
        "уже год": 365, "годами": 730, "много лет": 1095,
        "всю жизнь": 3650, "сколько себя помню": 3650
    }

    for phrase, days in fuzzy_mapping.items():
        if phrase in tl:
            return days

    return None


def calculate_duration_score(days):
    if days is None:
        return 0.0
    return 1.0 - math.exp(-days / CONFIG["DURATION_COEF"])


# ═══════════════════════════════════════════════════════════════════════
# REPEATED COMPLAINT (повторные обращения)
# ═══════════════════════════════════════════════════════════════════════

def detect_repeated_complaint(text):
    if not text:
        return False, None
    tl = text.lower().replace("ё", "е")
    for marker in REPEATED_MARKERS:
        if marker.replace("ё", "е") in tl:
            return True, marker
    return False, None


# ═══════════════════════════════════════════════════════════════════════
# CRITICALITY (критичность)
# ═══════════════════════════════════════════════════════════════════════

def detect_criticality_score_and_level(text):
    if not text:
        return 0.0, "НИЗКИЙ"

    tl = text.lower().replace("ё", "е")
    base_score = 0.0
    level = "НИЗКИЙ"

    has_max = any(re.search(pat, tl) for pat in CRITICALITY_PATTERNS["MAXIMAL"])
    has_high = any(re.search(pat, tl) for pat in CRITICALITY_PATTERNS["HIGH"])
    has_med = any(re.search(pat, tl) for pat in CRITICALITY_PATTERNS["MEDIUM"])
    has_low = any(re.search(pat, tl) for pat in CRITICALITY_PATTERNS["LOW"])

    if has_max:
        base_score = CONFIG["W_EXT"]
        level = "МАКСИМАЛЬНЫЙ"
    elif has_high:
        base_score = CONFIG["W_HIGH"]
        level = "ВЫСОКИЙ"
    elif has_med:
        base_score = CONFIG["W_MED"]
        level = "СРЕДНИЙ"
    elif has_low:
        base_score = CONFIG["W_LOW"]
        level = "НИЗКИЙ"

    bonus = 0.0
    has_object = any(re.search(obj, tl) for obj in DANGER_OBJECTS)
    has_harm = any(re.search(harm, tl) for harm in HARM_CONSEQUENCES)

    if has_object and has_harm:
        bonus = 0.10
    elif has_object or has_harm:
        bonus = 0.05

    final_score = clamp01(base_score + bonus)

    if final_score >= 0.90:
        level = "МАКСИМАЛЬНЫЙ"
    elif final_score >= 0.65:
        level = "ВЫСОКИЙ"
    elif final_score >= 0.35:
        level = "СРЕДНИЙ"
    else:
        level = "НИЗКИЙ"

    return final_score, level


# ═══════════════════════════════════════════════════════════════════════
# SYSTEMATICITY (системность)
# ═══════════════════════════════════════════════════════════════════════

def calculate_systematicity_days_and_score(text, address=None, date_history=None):
    systematicity_days = 0
    tl = text.lower() if text else ""

    text_detected = False
    systematic_markers = [
        "постоянно", "каждый год", "уже не первый раз", "из года в год",
        "регулярно", "систематически", "каждый сезон", "каждую весну",
        "каждую осень", "каждую зиму", "каждое лето", "ежегодно",
        "хроническ", "перманентн"
    ]
    if any(word in tl for word in systematic_markers):
        text_detected = True
        systematicity_days = 14

    if address and date_history and len(date_history) > 1:
        sorted_dates = sorted(date_history)
        delta = (sorted_dates[-1] - sorted_dates[0]).days
        systematicity_days = max(systematicity_days, delta)

    score = 1.0 - math.exp(-systematicity_days / CONFIG["SYSTEMATICITY_COEF"])

    if text_detected:
        score = clamp01(score + 0.08)

    return systematicity_days, score


# ═══════════════════════════════════════════════════════════════════════
# COLLECTIVE (коллективные жалобы)
# ═══════════════════════════════════════════════════════════════════════

def detect_collective_complaint(text, is_collective_status=False, n_signers=0):
    if is_collective_status:
        score = 1.0 - math.exp(-n_signers / CONFIG["COLLECTIVE_COEF"])
        return True, n_signers, score

    if not text:
        return False, 0, 0.0

    tl = text.lower()

    collective_patterns = [
        r"\bмы\b", r"\bнам\b", r"\bу нас\b", r"\bнаш\w*\b", r"\bсосед[ия]\b",
        r"\bжител[ияь]\w*\b", r"\bжильц[ыа]\w*\b", r"всем домом", r"весь дом",
        r"наш подъезд", r"наша улица", r"коллективное обращение", r"коллективная жалоба",
        r"от лица жителей", r"просим от имени жителей", r"инициативная группа",
        r"собрание жильцов", r"совет дома", r"все соседи", r"весь подъезд",
        r"все жильцы", r"общедомов\w*"
    ]

    anti_patterns = [
        r"мы с мужем", r"мы с ребенком", r"мы с мамой", r"мы с женой",
        r"мы с девушкой", r"мы семья", r"мы вдвоем", r"мы вдвоём",
        r"мы с супруг\w*", r"мы с братом", r"мы с сестрой"
    ]

    has_collective = any(re.search(pat, tl) for pat in collective_patterns)
    has_anti = any(re.search(pat, tl) for pat in anti_patterns)

    is_collective = False
    score = 0.0
    detected_signers = 0

    if has_collective:
        if has_anti:
            strong_markers = [
                r"коллективн\w*", r"инициативная группа", r"собрание жильцов",
                r"жители дома", r"подписи", r"весь дом", r"весь подъезд"
            ]
            if any(re.search(sm, tl) for sm in strong_markers):
                is_collective = True
        else:
            is_collective = True

    if is_collective:
        signer_match = re.search(r"(\d+)\s*(?:подпис(?:ей|и|антов)|жителей|человек|квартир)", tl)
        if signer_match:
            detected_signers = int(signer_match.group(1))
            score = 1.0 - math.exp(-detected_signers / CONFIG["COLLECTIVE_COEF"])
        else:
            score = CONFIG["COLLECTIVE_MIN"]
            detected_signers = 3

    return is_collective, detected_signers, score


# ═══════════════════════════════════════════════════════════════════════
# ADDRESS EXTRACTION
# ═══════════════════════════════════════════════════════════════════════

ADDRESS_PATTERNS = [
    r"""
    (?:дом[а]?|д\.)\s*№?\s*(?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)\s+(?:по|на)\s+
    (?P<street_type>ул\.?|улице|улица|проспекте|проспект|пр-т\.?|пр\.?|шоссе|проезде|проезд|переулке|переулок|пер\.?|бульваре|бульвар|бул\.?|набережной|набережная|наб\.?|площади|площадь|пл\.?)\s+
    (?P<street>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})(?=,|\.|;|:|\n|$)
    """,
    r"""
    (?:по|на)\s+
    (?P<street_type>ул\.?|улице|улица|проспекте|проспект|пр-т\.?|пр\.?|шоссе|проезде|проезд|переулке|переулок|пер\.?|бульваре|бульвар|бул\.?|набережной|набережная|наб\.?|площади|площадь|пл\.?)\s+
    (?P<street>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})\s*,?\s*(?:дом|д\.)\s*№?\s*(?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """,
    r"""
    (?:(?:г\.?|город|пос\.?|поселок|посёлок|д\.?|деревня|село|рп|пгт)\s+[А-Яа-яЁёA-Za-z0-9\-\s]+,\s*)?
    (?P<street_type>ул\.?|улица|проспект|пр-т\.?|пр\.?|шоссе|проезд|переулок|пер\.?|бульвар|бул\.?|набережная|наб\.?|площади|площадь|пл\.?)\s+
    (?P<street>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})\s*,?\s*(?:дом|д\.)\s*№?\s*(?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """,
    r"""
    (?P<street>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-]*)\s+
    (?P<street_type>проезд|проспект|бульвар|переулок|шоссе|набережная|площадь|аллея|тупик|улица|ул\.?)\s*,?\s+
    (?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """,
    r"""
    (?P<micro_type>мкр\.?|микрорайон)\s+(?P<micro>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})\s*,?\s*(?:дом|д\.)\s*№?\s*(?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """
]

STREET_TYPE_NORMALIZATION = {
    "ул": "ул.", "ул.": "ул.", "улица": "ул.", "улице": "ул.",
    "проспект": "проспект", "проспекте": "проспект", "пр-т": "проспект", "пр.": "проспект",
    "шоссе": "шоссе", "проезд": "проезд", "проезде": "проезд",
    "переулок": "переулок", "переулке": "переулок", "пер.": "переулок",
    "бульвар": "бульвар", "бульваре": "бульвар", "бул.": "бульвар",
    "набережная": "набережная", "набережной": "набережная", "наб.": "набережная",
    "площадь": "площадь", "площади": "площадь", "пл.": "площадь"
}


def normalize_street_type(value):
    value = normalize_text(value).lower().replace("ё", "е")
    return STREET_TYPE_NORMALIZATION.get(value, value)


def clean_address_piece(value):
    value = normalize_text(value).strip(" ,.;:")
    value = re.sub(
        r"\s+(обращаюсь|обращаемся|с\s+жалобой|жалобой|прошу|просим|находится|расположен[аоы]?|возле|около|рядом|напротив)\b.*$",
        "", value, flags=re.IGNORECASE
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
    if not candidates:
        simple = re.search(
            r"дом[а]?\s*№?\s*(\d+[А-Яа-яA-Za-z0-9\-\/]*)\s+по\s+улице\s+([А-Яа-яЁёA-Za-z0-9\-\s]+?)(?=,|\.|;|:|\n|$)",
            text, flags=re.IGNORECASE
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
        key = item.lower().replace("ё", "е")
        if key not in seen:
            seen.add(key)
            unique.append(item)
    unique.sort(key=len)
    return unique[0] if unique else ""


# ═══════════════════════════════════════════════════════════════════════
# MAIN ANALYSIS FUNCTION
# ═══════════════════════════════════════════════════════════════════════

def analyze_appeal(subject, text, is_collective_status=False, n_signers=0, address_history_dates=None):
    """
    Комплексный анализ обращения.
    """
    subject = subject or ""
    text = text or ""
    full_text = f"{subject}\n{text}"

    clean_text = normalize_text(full_text)

    # 1. Эмоциональность и класс эмоции
    emotion_score = calculate_emotion_score(clean_text)
    emotion_class = classify_emotion_class(clean_text)

    # 2. Доверие к эмоции на основе длины текста
    text_len = len(clean_text)
    conf_text = max(
        CONFIG["CONF_TEXT_MIN"],
        1.0 - math.exp(-text_len / CONFIG["CONF_TEXT_THRESHOLD"])
    )

    # Скорректированная эмоция
    emotion_adj = emotion_score * conf_text

    # 3. Уверенность в классификации эмоции
    emotion_class_confidence = get_emotion_confidence(clean_text, emotion_class)

    # 4. Социальный класс
    social_class = extract_social_class(clean_text)

    # 5. Длительность проблемы
    duration_days = extract_duration_days(clean_text)
    duration_score = calculate_duration_score(duration_days)

    # 6. Адрес
    address = extract_address(clean_text)

    # 7. Системность проблемы
    systematicity_days, systematicity_score = calculate_systematicity_days_and_score(
        clean_text, address=address, date_history=address_history_dates
    )

    # 8. Коллективность
    is_collective, final_signers, collective_score = detect_collective_complaint(
        clean_text, is_collective_status=is_collective_status, n_signers=n_signers
    )

    # 9. Критичность и факт-скор
    fact_score = 0.0
    if address:
        fact_score += 1.0
    if duration_days is not None:
        fact_score += 1.0
    if is_collective:
        fact_score += 1.0

    fact_max = 3.0
    fact01 = clamp01(fact_score / fact_max)

    criticality_score, criticality_level = detect_criticality_score_and_level(clean_text)

    # Повторное обращение
    is_repeated, repeated_marker = detect_repeated_complaint(clean_text)

    # Формула criticality_adj
    criticality_adj = clamp01(
        CONFIG["CRITICAL_COEF"] * fact01 + (1.0 - CONFIG["CRITICAL_COEF"]) * criticality_score
    )

    # 10. Итоговый индекс
    index_raw = (
        CONFIG["WEIGHT_EMOTIONAL"] * emotion_adj +
        CONFIG["WEIGHT_CRITICALITY"] * criticality_adj +
        CONFIG["WEIGHT_DURATION"] * duration_score +
        CONFIG["WEIGHT_SYSTEMATICITY"] * systematicity_score +
        CONFIG["WEIGHT_COLLECTIVE"] * collective_score
    )

    # 11. Защита "коротко, но опасно"
    index_final = index_raw
    if criticality_level == "МАКСИМАЛЬНЫЙ":
        index_final = max(index_raw, 0.70)
    elif criticality_level == "ВЫСОКИЙ":
        index_final = max(index_raw, 0.50)

    # Повторная жалоба
    if is_repeated:
        index_final = max(index_final, CONFIG["REPEATED_MIN_INDEX"])

    index_final = clamp01(index_final)

    # Приоритет реагирования
    priority_level, priority_label = get_priority(index_final)

    return {
        "created_by": "rule_based_analyzer_v3",
        "analyzed_at": datetime.now().isoformat(timespec="seconds"),

        "emotion_class": emotion_class,
        "emotion_label": emotion_class.capitalize(),
        "emotion_score": round(emotion_score, 3),
        "emotion_confidence": round(conf_text, 3),
        "emotion_class_confidence": round(emotion_class_confidence, 3),
        "emotion_adj": round(emotion_adj, 3),

        "social_class": social_class,

        "criticality_level": criticality_level,
        "criticality_score": round(criticality_score, 3),
        "criticality_adj": round(criticality_adj, 3),
        "fact_score": fact_score,

        "duration_days": duration_days,
        "duration_score": round(duration_score, 3),

        "systematicity_days": systematicity_days,
        "systematicity_score": round(systematicity_score, 3),

        "isCollective": is_collective,
        "n_signers": final_signers,
        "collective_score": round(collective_score, 3),

        "address": address,
        "is_repeated": is_repeated,
        "repeated_marker": repeated_marker,
        "index_final": round(index_final, 3),
        "priority_level": priority_level,
        "priority_label": priority_label
    }

def get_priority(index_value):
    if index_value >= 0.85:
        return "urgent", "Срочно"
    elif index_value >= 0.70:
        return "high", "Высокий приоритет"
    elif index_value >= 0.40:
        return "medium", "Средний приоритет"
    return "planned", "Плановый"