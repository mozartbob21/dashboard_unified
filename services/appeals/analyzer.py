import re
import math
from datetime import datetime

# Настраиваемые параметры системы (Конфиг согласно п. 12 ТЗ)
CONFIG = {
    # Веса для итогового индекса
    "WEIGHT_EMOTIONAL": 0.30,
    "WEIGHT_CRITICALITY": 0.30,
    "WEIGHT_DURATION": 0.20,
    "WEIGHT_SYSTEMATICITY": 0.10,
    "WEIGHT_COLLECTIVE": 0.10,
    
    # Параметры доверия к тексту
    "CONF_TEXT_THRESHOLD": 300, # символов
    "CONF_TEXT_MIN": 0.20,
    
    # Параметры коллективных жалоб
    "COLLECTIVE_MIN": 0.2,
    "COLLECTIVE_COEF": 15.0, # скорость насыщения n_signers
    
    # Длительность и системность
    "DURATION_COEF": 20.0,      # дней
    "SYSTEMATICITY_COEF": 90.0, # дней
    
    # Критичность
    "CRITICAL_COEF": 0.60,      # Соотношение баллов фактов и ключевых слов (60% факты, 40% слова)
    "W_LOW": 0.25,
    "W_MED": 0.50,
    "W_HIGH": 0.75,
    "W_EXT": 1.00,

    # === ДОБАВЛЕННЫЕ ПАРАМЕТРЫ ДЛЯ ИСПРАВЛЕНИЯ ОШИБКИ ===
    "MAX_PUNCT_EXCLAM": 0.30,   # Максимальный балл за восклицательные знаки (!)
    "MAX_PUNCT_QUEST": 0.15,    # Максимальный балл за вопросительные знаки (?)
    "MAX_CAPS_SCORE": 0.25,     # Максимальный балл за КАПСЛОК
}

# Обсценная/агрессивная лексика
PROFANITY_DICT = [
    r"блять", r"нах\w*", r"пох\w*", r"ху[йяеи]\w*", r"пизд\w*", r"муд\w+", r"гандон\w*",
    r"охренели", r"офигели", r"очумели", r"сволочи", r"твари", r"набить морду", r"упыри",
    r"придурки", r"дебилы", r"тупые"
]

# Усилители эмоций и сильные фразы из ТЗ и Excel
INTENSIFIERS = [
    "очень", "крайне", "совершенно", "абсолютно", "просто", "полностью", "категорически",
    "критически", "постоянно", "регулярно", "ежедневно", "еженедельно", "систематически",
    "каждый день", "каждый вечер", "каждую неделю", "несколько раз", "снова", "опять",
    "вновь", "уже", "давно", "долго", "никогда", "невозможно", "длительное время",
    "уже давно", "издевательство", "беспредел", "позор", "кошмар", "бардак", "ужас",
    "доколе", "сколько можно", "нет сил", "в бешенстве", "задыхаемся", "нас травят"
]

# Жалобные маркеры для базового негатива (п. 4.3.6)
COMPLAINT_MARKERS = [
    "нет", "не работает", "отсутств", "сломал", "плохо", "ужас", "грязная", "ржавая",
    "вонь", "течет", "безобразие", "отписка", "игнор", "не чинят", "отключили"
]

# Словари критичности по уровням (п. 8.1)
CRITICALITY_PATTERNS = {
    "MAXIMAL": [
        r"погиб\w*", r"смерть", r"летальный", r"реанимация", r"сбили реб[её]нка",
        r"дтп с пострадавшими", r"труп\w*", r"ударило током", r"утечка газа",
        r"пожар", r"взрыв", r"искрит", r"электричество\s+искрит", r"проводка\s+искрит",
        r"обрыв лэп", r"отравление"
    ],
    "HIGH": [
        r"открытый люк", r"провал", r"аварийность", r"перелом\w*", r"госпитализация",
        r"сильное кровотечение", r"реб[её]нок пострадал", r"провалился в люк",
        r"сбила машина", r"сотрясение", r"нет воды", r"нет хвс", r"нет гвс", r"нет отопления"
    ],
    "MEDIUM": [
        r"упал\w*", r"травма\w*", r"ушиб\w*", r"порез\w*", r"ожог\w*", r"опасно\b",
        r"угроза\b", r"слабый напор", r"нет напора", r"ржавая вода", r"грязная вода",
        r"коричневая вода", r"запах канализации", r"воняет", r"черви"
    ],
    "LOW": [
        r"может травмироваться", r"опасно ходить", r"риск\b", r"небезопасно"
    ]
}

# Объекты опасности и вред для контекстного бонуса (п. 8.1)
DANGER_OBJECTS = [r"люк", r"яма", r"провал", r"голол[её]д", r"сосульки", r"провод", r"крыша"]
HARM_CONSEQUENCES = [r"упал", r"травма\w*", r"сломал\w*", r"кровь\w*", r"пострадал\w*", r"разбил\w*"]

# Психолингвистические словари для классификации эмоций по 7 классам (п. 4.1)
EMOTION_DICTIONARIES = {
    "ГНЕВ": [
        r"беспредел", r"позор", r"бардак", r"набить морду", r"издевательство", r"издеваетесь",
        r"очковтерательство", r"блокировать", r"сговор", r"халатность", r"дурить людей",
        r"достали", r"дурдом", r"маразм", r"анархия", r"капитализм", r"всем пофиг", r"всем плевать"
    ],
    "ОТЧАЯНИЕ": [
        r"помогите", r"безысходность", r"беспомощно", r"выживают", r"как жить", r"последняя надежда",
        r"не знаем что делать", r"задыхаемся", r"страдания", r"сил нет", r"терпеть"
    ],
    "СТРАХ": [
        r"опасно", r"страшно", r"боимся", r"угроза", r"риск", r"отравиться", r"заболеть",
        r"паника", r"взрыв", r"убьет", r"опасно для жизни", r"смертельно"
    ],
    "ОТВРАЩЕНИЕ": [
        r"грязь", r"мерзость", r"вонь", r"черви", r"вонючая", r"туалет", r"фекалии", r"слизь",
        r"помойка", r"гнилая", r"рыжая", r"ржавая", r"сероводород", r"квас", r"кока-кола", r"тухл"
    ],
    "СКОРБЬ": [
        r"грустно", r"плакать", r"печально", r"умер", r"погиб", r"жаль", r"обидно", r"стыд"
    ],
    "РАЗДРАЖЕНИЕ": [
        r"возмущен", r"безобразие", r"отписка", r"надоело", r"сколько можно", r"игнорируют",
        r"бездействие", r"закрывают заявки", r"кормят завтраками", r"не реагируют", r"шаблонно"
    ],
    "АПАТИЯ": [
        r"все равно", r"пофиг", r"устали жаловаться", r"бесполезно", r"очередной год", r"как всегда"
    ]
}

def clamp01(x):
    return max(0.0, min(1.0, float(x)))

def normalize_text(text):
    if not text:
        return ""
    text = str(text).replace("\r", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def calculate_emotion_score(text):
    """
    Эвристический расчет интенсивности эмоций (emotion_score) согласно п. 4.3 ТЗ.
    Диапазон: [0..1]
    """
    if not text:
        return 0.0
        
    tl = text.lower()
    
    # 1. Базовый негатив (п. 4.3.6)
    baseline = 0.0
    if any(marker in tl for marker in COMPLAINT_MARKERS):
        baseline = 0.05
        
    # 2. Пунктуационная интенсивность (п. 4.3.1)
    exclam = text.count("!")
    quest = text.count("?")
    punct_score = min(CONFIG["MAX_PUNCT_EXCLAM"], exclam / 20.0) + min(CONFIG["MAX_PUNCT_QUEST"], quest / 10.0)
    
    # 3. Вклад капса (п. 4.3.2)
    letters = [c for c in text if c.isalpha()]
    if letters:
        all_letters_count = len(letters)
        upper_letters_count = sum(1 for c in letters if c.isupper())
        caps_ratio = upper_letters_count / all_letters_count
        caps_score = min(CONFIG["MAX_CAPS_SCORE"], 0.35 * caps_ratio)
    else:
        caps_score = 0.0
        
    # 4. Обсценная/агрессивная лексика (п. 4.3.3)
    prof_score = 0.0
    for pattern in PROFANITY_DICT:
        if re.search(pattern, tl, flags=re.IGNORECASE):
            prof_score = 0.20
            break
            
    # 5. Усилители и сильные фразы (п. 4.3.4)
    inten_hits = sum(1 for word in INTENSIFIERS if word in tl)
    inten_score = min(0.25, 0.04 * inten_hits)
    
    total_score = baseline + punct_score + caps_score + prof_score + inten_score
    return clamp01(total_score)

def classify_emotion_class(text):
    """
    Классифицирует эмоцию обращения на один из 7 классов на основе словарей (п. 4.1 ТЗ)
    """
    if not text:
        return "РАЗДРАЖЕНИЕ" # Default
        
    tl = text.lower()
    scores = {}
    
    for emotion, patterns in EMOTION_DICTIONARIES.items():
        count = 0
        for pattern in patterns:
            if re.search(pattern, tl, flags=re.IGNORECASE):
                count += 1
        scores[emotion] = count
        
    max_emotion = max(scores, key=scores.get)
    if scores[max_emotion] == 0:
        return "РАЗДРАЖЕНИЕ" # Значение по умолчанию для жалоб
    return max_emotion

def extract_social_class(text):
    """
    Формирует социальный класс на основе ТЗ (п. 5). Не участвует в формуле индекса.
    """
    if not text:
        return "Житель"
        
    tl = text.lower()
    classes = []
    
    if any(word in tl for word in ["маломобильн", "коляск", "пандус", "инвалид"]):
        classes.append("Маломобильные")
    if any(word in tl for word in ["дети", "ребенок", "ребёнок", "новорожден", "садик", "сын", "дочь", "малыш"]):
        classes.append("Семьи с детьми")
    if any(word in tl for word in ["пенсионер", "старик", "бабушк", "дедушк", "пожил"]):
        classes.append("Пенсионеры")
    if any(word in tl for word in ["многодетн"]):
        classes.append("Многодетные")
    if any(word in tl for word in ["ветеран"]):
        classes.append("Ветераны")
        
    if not classes:
        return "Житель"
    return ", ".join(classes)

def extract_duration_days(text):
    """
    Вычленяет явные и неявные указания времени из текста (п. 6.1)
    """
    if not text:
        return None
        
    tl = text.lower().replace("ё", "е")
    
    word_numbers = {
        "один": 1, "одна": 1, "два": 2, "две": 2, "три": 3, "четыре": 4, 
        "пять": 5, "шесть": 6, "семь": 7, "восемь": 8, "девять": 9, "десять": 10
    }
    
    # Регулярные выражения для явных указаний времени
    # Шаблоны дней
    for m in re.finditer(r"(\d+)\s*(?:день|дня|дней|сутки|суток)", tl):
        return int(m.group(1))
    for m in re.finditer(rf"\b({'|'.join(word_numbers.keys())})\s*(?:день|дня|дней|сутки|суток)", tl):
        return word_numbers[m.group(1)]
        
    # Шаблоны недель
    for m in re.finditer(r"(\d+)\s*(?:недел)", tl):
        return int(m.group(1)) * 7
    for m in re.finditer(rf"\b({'|'.join(word_numbers.keys())})\s*(?:недел)", tl):
        return word_numbers[m.group(1)] * 7
        
    # Шаблоны месяцев
    for m in re.finditer(r"(\d+)\s*(?:месяц)", tl):
        return int(m.group(1)) * 30
    for m in re.finditer(rf"\b({'|'.join(word_numbers.keys())})\s*(?:месяц)", tl):
        return word_numbers[m.group(1)] * 30
        
    # Шаблоны лет
    for m in re.finditer(r"(\d+)\s*(?:год|лет)", tl):
        return int(m.group(1)) * 365
        
    # Эвристические бакеты (п. 6.1.4)
    fuzzy_mapping = {
        "каждый день": 7, "ежедневно": 7, "постоянно": 14, "уже не первый раз": 14,
        "давно": 30, "длительное время": 30, "уже месяц": 30, "уже год": 365, "годами": 730
    }
    
    for phrase, days in fuzzy_mapping.items():
        if phrase in tl:
            return days
            
    return None

def calculate_duration_score(days):
    if days is None:
        return 0.0
    return 1.0 - math.exp(-days / CONFIG["DURATION_COEF"])

def detect_criticality_score_and_level(text):
    """
    Расчет критичности ситуации по ключевым словам и логическим правилам (п. 8)
    """
    if not text:
        return 0.0, "НИЗКИЙ"
        
    tl = text.lower().replace("ё", "е")
    base_score = 0.0
    level = "НИЗКИЙ"
    
    # 1. Поиск совпадений по словарям
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
        
    # 2. Добавление бонуса за контекст (Объект опасности + Вред)
    bonus = 0.0
    has_object = any(re.search(obj, tl) for obj in DANGER_OBJECTS)
    has_harm = any(re.search(harm, tl) for harm in HARM_CONSEQUENCES)
    
    if has_object and has_harm:
        bonus = 0.10
        
    final_score = clamp01(base_score + bonus)
    
    # Переопределение уровня после применения бонуса, если применимо
    if final_score >= 0.90:
        level = "МАКСИМАЛЬНЫЙ"
    elif final_score >= 0.65:
        level = "ВЫСОКИЙ"
    elif final_score >= 0.35:
        level = "СРЕДНИЙ"
    else:
        level = "НИЗКИЙ"
        
    return final_score, level

def calculate_systematicity_days_and_score(text, address=None, date_history=None):
    """
    Оценка системности проблемы (п. 7)
    """
    systematicity_days = 0
    tl = text.lower() if text else ""
    
    # Эвристика по тексту (п. 7.1.1)
    text_detected = False
    if any(word in tl for word in ["постоянно", "каждый год", "уже не первый раз", "из года в год", "регулярно"]):
        text_detected = True
        systematicity_days = 14 # Условный период при словесном триггере
        
    # Расчет по истории дат по конкретному адресу (п. 7.1.2)
    if address and date_history and len(date_history) > 1:
        # date_history - список объектов datetime для данного адреса
        sorted_dates = sorted(date_history)
        delta = (sorted_dates[-1] - sorted_dates[0]).days
        systematicity_days = max(systematicity_days, delta)
        
    score = 1.0 - math.exp(-systematicity_days / CONFIG["SYSTEMATICITY_COEF"])
    
    # Добавляем надбавочный фиксированный коэффициент при текстовом обнаружении (п. 7.1.1)
    if text_detected:
        score = clamp01(score + 0.08)
        
    return systematicity_days, score

def detect_collective_complaint(text, is_collective_status=False, n_signers=0):
    """
    Автоматическое выявление коллективных жалоб по ТЗ (п. 10)
    """
    if is_collective_status:
        # Если статус явный, то поиск по ключевым словам не производится (п. 10.1.2)
        score = 1.0 - math.exp(-n_signers / CONFIG["COLLECTIVE_COEF"])
        return True, n_signers, score
        
    if not text:
        return False, 0, 0.0
        
    tl = text.lower()
    
    # Маркеры коллективности (regex с поддержкой падежей)
    collective_patterns = [
        r"\bмы\b", r"\bнам\b", r"\bу нас\b", r"\bнаш\w*\b", r"\bсосед[ия]\b", 
        r"\bжител[ияь]\w*\b", r"\bжильц[ыа]\w*\b", r"всем домом", r"весь дом",
        r"наш подъезд", r"наша улица", r"коллективное обращение", r"коллективная жалоба",
        r"от лица жителей", r"просим от имени жителей", r"инициативная группа", r"собрание жильцов"
    ]
    
    # Anti-правила (п. 10.1.3)
    anti_patterns = [
        r"мы с мужем", r"мы с ребенком", r"мы с мамой", r"мы с женой", r"мы с девушкой",
        r"мы семья", r"мы вдвоем", r"мы вдвоём"
    ]
    
    has_collective = any(re.search(pat, tl) for pat in collective_patterns)
    has_anti = any(re.search(pat, tl) for pat in anti_patterns)
    
    is_collective = False
    score = 0.0
    detected_signers = 0
    
    if has_collective:
        if has_anti:
            # Требуется дополнительное жесткое совпадение, не только слабое "мы"
            strong_markers = [r"коллективн\w*", r"инициативная группа", r"собрание жильцов", r"жители дома", r"подписи"]
            if any(re.search(sm, tl) for sm in strong_markers):
                is_collective = True
        else:
            is_collective = True
            
    if is_collective:
        # Пытаемся распарсить количество подписантов
        signer_match = re.search(r"(\d+)\s*(?:подпис(?:ей|и|антов)|жителей)", tl)
        if signer_match:
            detected_signers = int(signer_match.group(1))
            score = 1.0 - math.exp(-detected_signers / CONFIG["COLLECTIVE_COEF"])
        else:
            # При наличии только ключевых слов, score = COLLECTIVE_MIN = 0.2 (п. 10.1.4)
            score = CONFIG["COLLECTIVE_MIN"]
            detected_signers = 3 # Эквивалент по ТЗ
            
    return is_collective, detected_signers, score

def analyze_appeal(subject, text, is_collective_status=False, n_signers=0, address_history_dates=None):
    """
    Суперпрофессиональный комплексный анализ обращения.
    Интегрирует все требования ТЗ без потери исходного функционала.
    """
    subject = subject or ""
    text = text or ""
    full_text = f"{subject}\n{text}"
    
    # Предварительная подготовка (п. 13.2)
    clean_text = normalize_text(full_text)
    
    # 1. Эмоциональность и Класс эмоции
    emotion_score = calculate_emotion_score(clean_text)
    emotion_class = classify_emotion_class(clean_text)
    
    # 2. Доверие к эмоции на основе длины текста (п. 9)
    text_len = len(clean_text)
    conf_text = max(
        CONFIG["CONF_TEXT_MIN"],
        1.0 - math.exp(-text_len / CONFIG["CONF_TEXT_THRESHOLD"])
    )
    
    # Скорректированная эмоция
    emotion_adj = emotion_score * conf_text
    
    # 3. Социальный класс
    social_class = extract_social_class(clean_text)
    
    # 4. Длительность проблемы
    duration_days = extract_duration_days(clean_text)
    duration_score = calculate_duration_score(duration_days)
    
    # 5. Адрес
    address = extract_address(clean_text)
    
    # 6. Системность проблемы
    systematicity_days, systematicity_score = calculate_systematicity_days_and_score(
        clean_text, address=address, date_history=address_history_dates
    )
    
    # 7. Коллективность
    is_collective, final_signers, collective_score = detect_collective_complaint(
        clean_text, is_collective_status=is_collective_status, n_signers=n_signers
    )
    
    # 8. Критичность и Балл фактов (fact_score ∈ [0..3], п. 8.4)
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
    
    # Формула criticality_adj (п. 8.4)
    criticality_adj = clamp01(
        CONFIG["CRITICAL_COEF"] * fact01 + (1.0 - CONFIG["CRITICAL_COEF"]) * criticality_score
    )
    
    # 9. Линейная комбинация весов (Итоговый индекс заявки index_raw, п. 11.1)
    index_raw = (
        CONFIG["WEIGHT_EMOTIONAL"] * emotion_adj +
        CONFIG["WEIGHT_CRITICALITY"] * criticality_adj +
        CONFIG["WEIGHT_DURATION"] * duration_score +
        CONFIG["WEIGHT_SYSTEMATICITY"] * systematicity_score +
        CONFIG["WEIGHT_COLLECTIVE"] * collective_score
    )
    
    # 10. Защита "коротко, но опасно" (п. 11.2)
    index_final = index_raw
    if criticality_level == "МАКСИМАЛЬНЫЙ":
        index_final = max(index_raw, 0.70)
    elif criticality_level == "ВЫСОКИЙ":
        index_final = max(index_raw, 0.50)
        
    index_final = clamp01(index_final)
    
    # Приоритет реагирования
    priority_level, priority_label = get_priority(index_final)
    
    return {
        "created_by": "rule_based_analyzer_v2_res",
        "analyzed_at": datetime.now().isoformat(timespec="seconds"),
        
        "emotion_class": emotion_class,
        "emotion_label": emotion_class.capitalize(),
        "emotion_score": round(emotion_score, 3),
        "emotion_confidence": round(conf_text, 3),
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
        "index_final": round(index_final, 3),
        "priority_level": priority_level,
        "priority_label": priority_label
    }

# Вспомогательные функции
def get_priority(index_value):
    if index_value >= 0.85:
        return "urgent", "Срочно"
    elif index_value >= 0.70:
        return "high", "Высокий приоритет"
    elif index_value >= 0.40:
        return "medium", "Средний приоритет"
    return "planned", "Плановый"

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

# Шаблоны адресов и вспомогательный код парсера
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
    (?P<micro_type>мкр\.?|микрорайон)\s+(?P<micro>[А-Яа-яЁёA-Za-z0-9][А-Яа-яЁёA-Za-z0-9\-\s]{1,80})\s*,?\s*(?:дом|д\.)\s*№?\s*(?P<house>\d+[А-Яа-яA-Za-z0-9\-\/]*)
    """
]

STREET_TYPE_NORMALIZATION = {
    "ул": "ул.", "ул.": "ул.", "улица": "ул.", "улице": "ул.",
    "проспект": "проспект", "проспекте": "проспект", "пр-т": "проспект", "пр.": "проспект",
    "шоссе": "шоссе", "проезд": "проезд", "проезде": "проезд",
    "переулок": "переулок", "пер.": "переулок", "бульвар": "бульвар", "площадь": "площадь"
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