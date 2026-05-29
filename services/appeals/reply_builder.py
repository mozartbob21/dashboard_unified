from __future__ import annotations

import re


EMPATHY_BLOCKS = {
    "angry": """Понимаем Ваше обеспокоенность и недовольство сложившейся ситуацией. Описанные Вами обстоятельства будут рассмотрены по существу.

Информация, указанная в обращении, будет проверена в пределах компетенции ответственных специалистов. При подтверждении изложенных фактов будут рассмотрены возможные меры реагирования в установленном порядке.""",

    "anxious": """Понимаем Вашу обеспокоенность. Вопросы, которые могут затрагивать безопасность, комфорт проживания или нормальную работу объектов и служб, требуют внимательного рассмотрения.

Указанная Вами информация будет проанализирована и, при наличии оснований, направлена ответственным специалистам для проверки и принятия мер в пределах их полномочий.""",

    "sad": """Понимаем, что описанная ситуация могла доставить Вам неудобства. Ваше обращение принято к рассмотрению, изложенные обстоятельства будут проверены по существу.

Благодарим Вас за направленную информацию. Она поможет более точно оценить ситуацию и подготовить ответ по существу поставленного вопроса.""",

    "urgent": """Понимаем, что обозначенный вопрос требует внимательного и своевременного рассмотрения. Информация, изложенная в обращении, будет рассмотрена с учётом характера описанной ситуации.

При наличии оснований сведения будут направлены ответственным специалистам для проверки и принятия возможных мер реагирования в установленном порядке.""",

    "thankful": """Благодарим Вас за обращение и активную гражданскую позицию. Направленная Вами информация будет рассмотрена по существу.

Такая обратная связь помогает своевременно выявлять вопросы, требующие внимания со стороны ответственных служб.""",

    "neutral": """Ваше обращение принято и будет рассмотрено в установленном порядке.

Изложенные сведения будут проанализированы по существу, при наличии оснований информация будет направлена ответственным специалистам для проверки и принятия возможных мер.""",
}


DEFAULT_REPLY_TEMPLATE = """Уважаемый(ая) заявитель!

{{ empathy_block }}

По существу обращения сообщаем:
{{ facts }}

Дополнительно информируем, что обращение рассмотрено в установленном порядке. При наличии дополнительных сведений Вы можете направить их для приобщения к материалам рассмотрения.

С уважением,
Подмосковная Нейрона Роботовна
"""


def _normalize_text(value):
    return str(value or "").strip()


def _normalize_key(value):
    return _normalize_text(value).lower().replace("ё", "е")


def _cleanup_reply_text(text):
    text = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def get_empathy_block(emotion_class, item=None):
    """
    Возвращает только публичный эмпатичный блок для заявителя.

    ВАЖНО:
    сюда не добавляем внутреннюю аналитику:
    - баллы;
    - приоритеты;
    - эмоции;
    - критичность;
    - matched_words;
    - фразы вида «система учла параметры».
    Всё это можно показывать специалисту в интерфейсе, но нельзя вставлять в официальный ответ.
    """
    emotion_class = _normalize_key(emotion_class or "neutral")
    return _cleanup_reply_text(
        EMPATHY_BLOCKS.get(emotion_class, EMPATHY_BLOCKS["neutral"])
    )


def generate_reply_from_template(template_text, facts_text, item):
    template_text = template_text or DEFAULT_REPLY_TEMPLATE
    facts_text = facts_text or "Информация по обращению принята к рассмотрению."

    item = item or {}
    analysis = item.get("analysis_data") or {}
    emotion_class = analysis.get("emotion_class", "neutral")

    empathy_block = get_empathy_block(emotion_class, item=item)

    replacements = {
        "{{ request_id }}": item.get("request_id", ""),
        "{{ subject }}": item.get("subject", ""),
        "{{ sender_email }}": item.get("sender_email", ""),
        "{{ facts }}": facts_text,
        "{{ empathy_block }}": empathy_block,

        # Эти поля оставлены для совместимости шаблонов,
        # но в стандартный публичный текст они не подставляются.
        "{{ emotion_class }}": analysis.get("emotion_class", ""),
        "{{ emotion_label }}": analysis.get("emotion_label", ""),
        "{{ criticality_level }}": analysis.get("criticality_level", ""),
        "{{ criticality_label }}": analysis.get("criticality_label", ""),
        "{{ address }}": analysis.get("address", ""),
        "{{ priority_label }}": analysis.get("priority_label", ""),
        "{{ final_index }}": str(analysis.get("index_final", "")),
    }

    result = template_text

    for key, value in replacements.items():
        result = result.replace(key, str(value or ""))

    return _cleanup_reply_text(result)
