"""Presentation text rules, using synthetic data only and no app startup."""
import unittest
from services.collective import presentation_text as text


class CollectiveTextTests(unittest.TestCase):
    def test_explicit_address_in_fact_is_kept_alongside_annotation(self):
        from services.collective.presentation_text import content_blocks
        blocks = content_blocks({'Аннотация/краткое содержание': 'Просим провести ремонт.',
                                 'Факт': 'Адрес: г. Тестовый, ул. Примерная, д. 7'})
        self.assertEqual(blocks[0], ('Адрес: г. Тестовый, ул. Примерная, д. 7', True))
        self.assertIn(('Просим провести ремонт.', False), blocks)

    def test_markup_preserves_paragraphs_and_literal_comparisons(self):
        value = '<p>Давление &lt; 3, температура &gt;80.</p><p>Уровень <= 5<br>Повторный замер > 2.</p>'
        self.assertEqual(text.plain_text(value),
                         'Давление < 3, температура >80.\nУровень <= 5\nПовторный замер > 2.')
        self.assertEqual(text.plain_text('давление < 3, температура >80'),
                         'давление < 3, температура >80')
        self.assertEqual(text.plain_text('x<3 и y>80'), 'x<3 и y>80')

    def test_encoded_html_and_unknown_notation(self):
        self.assertEqual(text.plain_text('&lt;b&gt;В работе&lt;/b&gt;&amp;nbsp;'), 'В работе')
        self.assertEqual(text.plain_text('Условие <порог> сохранено'), 'Условие <порог> сохранено')
        self.assertEqual(text.plain_text('<p>Факт</p><script>alert(1)</script><!--noise--><p>Ответ</p>'), 'Факт\nОтвет')

    def test_leading_address_and_service_fields_are_separate(self):
        row = {'Аннотация/краткое содержание':
               'Россия, Московская область, Тестовый округ, ул. Примерная, д. 10\n'
               'Номер Добродела: 123456; ППМО: 123-456; Авторизация прошла через ЕСИА; '
               'Давление < 3. Просим завершить ремонт до 15 октября.',
               'Источник обращения ЕЦУР': 'Добродел',
               'Ведомство исполнителя': 'Министерство жилищно-коммунального хозяйства Московской области'}
        self.assertEqual(text.content_blocks(row), [
            ('Адрес: Тестовый округ, ул. Примерная, д. 10', True),
            ('Добродел / Министерство ЖКХ', False),
            ('Давление < 3. Просим завершить ремонт до 15 октября.', False)])

    def test_explicit_address_and_leading_metadata_on_separate_lines(self):
        value = 'Номер Добродела: 123456\nАдрес: ул. Примерная, д. 10; Просим ремонт\nВторой абзац'
        self.assertEqual(text.split_annotation(value),
                         ('ул. Примерная, д. 10', 'Просим ремонт\nВторой абзац'))

    def test_service_words_inside_substantive_prose_are_not_removed(self):
        examples = [
            'Номер Добродела: 123456 не соответствует заявлению.',
            'ППМО: 123-456 требует исправления.',
            'Авторизация прошла через ЕСИА, но ответ недоступен.',
            '10000 жителей остаются без воды.',
            'Просим проверить.\nНомер Добродела: 123456',
            'Просим проверить; Авторизация прошла через ЕСИА; это важно.',
        ]
        for value in examples:
            with self.subTest(value=value):
                self.assertEqual(text.split_annotation(value), ('', value))

    def test_unknown_organisations_and_sources_are_not_length_cut(self):
        organisation = 'АО «Тестовый региональный водоканал», филиал «Северный округ» ' + 'подразделение ' * 10
        row = {'Источник обращения ЕЦУР': '-', 'Источник': 'МСЭД',
               'Ведомство исполнителя': organisation}
        self.assertEqual(text.source_line(row), 'МСЭД / ' + organisation.strip())
        self.assertEqual(text.short_org('Администрация Тестового округа'), 'ОМСУ')
        self.assertEqual(text.short_org('Письмо администрации округа'), 'Письмо администрации округа')

    def test_long_content_and_answer_keep_final_facts_and_deadlines(self):
        content = 'Подробности обращения. ' * 100 + 'Обязательный последний факт: давление 2 бар.'
        answer = 'Результаты проверки. ' * 100 + 'Обещанный срок: 15 октября 2026 года.'
        self.assertEqual(text.content_blocks({'Аннотация/краткое содержание': content}), [(content, False)])
        self.assertEqual(text.status_blocks({'Статус': 'В работе', 'Ответ': answer}),
                         [('В работе', True), (answer, False)])

    def test_status_and_answer_are_distinct_without_invented_status(self):
        self.assertEqual(text.status_blocks({'Статус': '', 'Ответ': '<p>Ремонт завершён.</p><p>Проверка 15 октября.</p>'}),
                         [('Статус не указан', True), ('Ремонт завершён.\nПроверка 15 октября.', False)])
        self.assertEqual(text.status_blocks({}), [('Статус не указан', True)])

    def test_only_complete_identical_duplicate_is_collapsed(self):
        shared = 'Это общий длинный префикс больше сорока символов для проверки '
        value = shared + 'срок 15 октября; ' + shared + 'срок 18 октября'
        self.assertEqual(text.split_annotation(value), ('', value))
        self.assertEqual(text.split_annotation('Просим ремонт; Просим ремонт'), ('', 'Просим ремонт'))

    def test_empty_annotation_uses_fact_and_does_not_duplicate_address(self):
        self.assertEqual(text.content_blocks({'Аннотация/краткое содержание': '<p> </p>', 'Факт': 'Факт обращения'}),
                         [('Факт обращения', False)])
        self.assertEqual(text.content_blocks({'Факт': 'Адрес: ул. Примерная, д. 10'}),
                         [('Адрес: ул. Примерная, д. 10', True)])
        self.assertEqual(text.content_blocks({}), [('Содержание не указано', False)])


if __name__ == '__main__':
    unittest.main()
