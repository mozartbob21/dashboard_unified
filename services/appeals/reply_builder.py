EMPATHY_BLOCKS = {
    "angry": "Понимаем Ваше возмущение и обеспокоенность сложившейся ситуацией. Обращение принято в работу, изложенные обстоятельства будут проверены.",
    "anxious": "Понимаем Вашу обеспокоенность вопросом безопасности. Информация будет рассмотрена с учётом возможных рисков и необходимости принятия мер.",
    "sad": "Понимаем, что сложившаяся ситуация доставляет Вам неудобства. Ваше обращение принято и будет рассмотрено по существу.",
    "urgent": "Понимаем срочность обозначенного вопроса. Информация будет направлена для оперативного рассмотрения ответственными специалистами.",
    "thankful": "Благодарим Вас за обращение и активную гражданскую позицию.",
    "neutral": "Ваше обращение принято и будет рассмотрено в установленном порядке.",
}


DEFAULT_REPLY_TEMPLATE = """Уважаемый(ая) заявитель!

{{ empathy_block }}

По существу обращения сообщаем:
{{ facts }}

Дополнительно информируем, что обращение рассмотрено в установленном порядке. При наличии дополнительных сведений Вы можете направить их для приобщения к материалам рассмотрения.

С уважением,
Подмосковная Нейрона Роботовна
"""


def get_empathy_block(emotion_class):
    return EMPATHY_BLOCKS.get(emotion_class or "neutral", EMPATHY_BLOCKS["neutral"])


def generate_reply_from_template(template_text, facts_text, item):
    template_text = template_text or DEFAULT_REPLY_TEMPLATE
    facts_text = facts_text or "Информация по обращению принята к рассмотрению."

    analysis = item.get("analysis_data") or {}
    emotion_class = analysis.get("emotion_class", "neutral")
    empathy_block = get_empathy_block(emotion_class)

    replacements = {
        "{{ request_id }}": item.get("request_id", ""),
        "{{ subject }}": item.get("subject", ""),
        "{{ sender_email }}": item.get("sender_email", ""),
        "{{ facts }}": facts_text,
        "{{ empathy_block }}": empathy_block,
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

    return result.strip()