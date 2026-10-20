"""Contracts for the three additional DataLens sources; synthetic inputs only."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from services.water_dashboard import builder
from services.water_dashboard.config import SOURCES
from services.water_dashboard.details import (
    KONTUR_COMPLETION_HEADER, KONTUR_TASK_HEADER,
    is_kontur_ranking_table, source_details,
)
from services.water_dashboard.metrics import primary_available, source_metrics


def widget(label, value, caption=''):
    return {'label': label, 'value': value, 'caption': caption}


def values(sid, widgets):
    return {item['id']: item['value'] for item in source_metrics(sid, {'widgets': widgets}, [])}


def table(headers, rows):
    return {'headers': headers, 'rows': rows}


def kontur_table(rows):
    return table(['ОМСУ', 'РСО', KONTUR_TASK_HEADER, KONTUR_COMPLETION_HEADER], rows)


class NewSourceTests(unittest.TestCase):
    def test_registry_has_ten_unique_sources(self):
        self.assertEqual(len(SOURCES), 10)
        self.assertEqual(len({item['id'] for item in SOURCES}), 10)
        self.assertEqual(len({item['url'] for item in SOURCES}), 10)
        self.assertEqual({item['id'] for item in SOURCES[-3:]},
                         {'controller_tasks', 'energy_audit', 'gkh_kontur'})

    def test_zero_is_valid_and_missing_main_is_not_replaced(self):
        specs = [
            ('controller_tasks', 'Всего задач', 'Просрочено'),
            ('energy_audit', 'Доля (%) успешных переходов в оптимальную ценовую категорию', 'Всего'),
            ('gkh_kontur', '% выполнения задач ДК', 'Всего задач'),
        ]
        for sid, main, secondary in specs:
            with self.subTest(sid=sid):
                self.assertTrue(primary_available(source_metrics(sid, {'widgets': [widget(main, '0')]}, [])))
                self.assertFalse(primary_available(source_metrics(sid, {'widgets': [widget(secondary, '15')]}, [])))
                self.assertFalse(primary_available(source_metrics(sid, {}, [])))

    def test_controller_totals_are_taken_from_widgets_not_table_workload(self):
        data = {
            'widgets': [widget('Всего задач', '20'), widget('Просрочено', '3')],
            'tables': [table(['Контролер', '#', 'Контроль сроков', 'Категория'],
                             [['Контролёр А', '15', 'Есть просрочки', 'Категория А']])],
        }
        result = {item['id']: item['value'] for item in source_metrics('controller_tasks', data, [])}
        self.assertEqual(result, {'tasks_total': 20, 'overdue_tasks': 3})
        detail = source_details('controller_tasks', data)
        self.assertEqual(detail['entity_type'], 'controller')
        self.assertEqual(detail['groups'][0]['title'], 'Топ-5 по числу задач на проверке')
        self.assertEqual(detail['groups'][0]['items'][0]['value'], 15)
        self.assertNotIn('municipality', detail['entities'][0])
        self.assertIn('не число просроченных', detail['basis'])

    def test_controller_duplicates_conflicts_and_missing_people(self):
        headers = ['Контролер', '#', 'Контроль сроков', 'Категория']
        rows = [[f'Контролёр {i}', str(i), 'В графике', 'Категория'] for i in range(7)]
        rows += [['null', '80', '', ''], ['', '90', '', ''], ['Неизвестный', '—', '', '']]
        data = {'tables': [table(headers, rows), table(headers, rows),
                           table(headers, [['Контролёр 2', '9', 'В графике', 'Категория']])]}
        detail = source_details('controller_tasks', data)
        self.assertEqual([row['value'] for row in detail['groups'][0]['items']], [6, 5, 4, 3, 1])
        self.assertEqual(detail['coverage']['excluded'], 2)
        self.assertEqual(detail['coverage']['entities'], 6)

    def test_energy_repeated_titles_are_matched_with_exact_captions(self):
        widgets = [
            widget('Доля (%) успешных переходов в оптимальную ценовую категорию', '18'),
            widget('Всего', '120', 'приборов учета (шт.)'),
            widget('Всего', '999', 'другой показатель'),
            widget('Приборы учета', '30', 'с потреблением свыше 10 000 кВт·ч в месяц (шт.)'),
            widget('Затраты', '12,5', 'на оплату ЭЭ (все ПУ) (млн руб.)'),
            widget('Затраты', '8,2', 'на оплату ЭЭ ПУ (свыше 10 000 кВт·ч в месяц) (млн руб.)'),
            widget('Потребление', '30,4', '(все ПУ) (млн кВт*ч)'),
            widget('Потребление', '20,1', '(все ПУ) свыше 10 000 кВт·ч в месяц (млн кВт*ч)'),
        ]
        result = values('energy_audit', widgets)
        self.assertEqual(result, {
            'optimal_tariff_pct': 18, 'meters_total': 120, 'meters_high_consumption': 30,
            'energy_cost_total_mln': 12.5, 'energy_cost_high_mln': 8.2,
            'consumption_total_mln_kwh': 30.4, 'consumption_high_mln_kwh': 20.1,
        })
        self.assertEqual(result, values('energy_audit', list(reversed(widgets))))
        details = source_details('energy_audit', {'metrics': source_metrics('energy_audit', {'widgets': widgets}, [])})
        self.assertEqual(details['kind'], 'metrics')
        self.assertEqual(len(details['metrics']), 7)
        self.assertEqual(details['metrics'][3]['unit'], 'млн ₽')

    def test_energy_ambiguous_missing_and_invalid_captions_do_not_produce_values(self):
        base = widget('Затраты', '10', 'на оплату ЭЭ (все ПУ) (млн руб.)')
        self.assertEqual(values('energy_audit', [base, base])['energy_cost_total_mln'], 10)
        conflicting = {**base, 'value': '11'}
        self.assertIsNone(values('energy_audit', [base, conflicting])['energy_cost_total_mln'])
        self.assertIsNone(values('energy_audit', [widget('Затраты', '99')])['energy_cost_total_mln'])
        self.assertIsNone(values('energy_audit', [{**base, 'value': '10%'}])['energy_cost_total_mln'])
        self.assertIsNone(values('energy_audit', [widget('Всего', '12.5', 'приборов учета (шт.)')])['meters_total'])

    def test_percentages_are_bounded_and_kontur_keeps_official_formula(self):
        widgets = [widget('% выполнения задач ДК', '12,28%'),
                   widget('Всего задач', '900'), widget('Закрыто задач', '180')]
        result = values('gkh_kontur', widgets)
        self.assertEqual(result, {'completion_pct': 12.28, 'tasks_total': 900, 'tasks_closed': 180})
        self.assertNotEqual(result['completion_pct'], result['tasks_closed'] / result['tasks_total'] * 100)
        for sid, label in [('gkh_kontur', '% выполнения задач ДК'),
                           ('energy_audit', 'Доля (%) успешных переходов в оптимальную ценовую категорию')]:
            for invalid in ('101%', '-1', '—', True):
                with self.subTest(sid=sid, value=invalid):
                    self.assertFalse(primary_available(source_metrics(sid, {'widgets': [widget(label, invalid)]}, [])))

    def test_kontur_ranks_full_table_and_preserves_zero(self):
        rows = [[f'Округ {i}', f'РСО {i}', str(100 + i), str(i * 10)] for i in range(8)]
        data = {'ranking_tables': [kontur_table(rows[:5]), kontur_table(rows[5:]), kontur_table(rows[:5])]}
        details = source_details('gkh_kontur', data)
        self.assertEqual(details['entity_type'], 'rso')
        self.assertEqual(details['coverage']['entities'], 8)
        self.assertEqual([row['value'] for row in details['groups'][0]['items']], [0, 10, 20, 30, 40])
        self.assertEqual([row['value'] for row in details['groups'][1]['items']], [70, 60, 50, 40, 30])
        self.assertEqual(details['entities'][0]['municipality'], 'Округ 0')
        self.assertEqual(details['entities'][0]['task_id'], '100')

    def test_kontur_branches_are_distinct_conflicts_excluded(self):
        rows = [
            ['Округ', 'РСО', '101', '0'], ['Округ', 'РСО', '102', '100'],
            ['Округ', 'Спорная РСО', '103', '30'], ['Округ', 'Спорная РСО', '103', '40'],
            ['Округ', 'Неизвестно', '104', '—'], ['Округ', 'Ошибочно', '105', '150'],
        ]
        details = source_details('gkh_kontur', {'ranking_tables': [kontur_table(rows)]})
        self.assertEqual(details['coverage']['entities'], 2)
        self.assertEqual(details['coverage']['excluded'], 3)
        self.assertEqual({row['task_id'] for row in details['entities']}, {'101', '102'})

    def test_kontur_preserves_legal_abbreviations_in_organisation_names(self):
        name = 'МУП "ТЕСТОВЫЙ ВОДОКАНАЛ"'
        details = source_details('gkh_kontur', {
            'ranking_tables': [kontur_table([['ОКРУГ', name, '101', '90']])],
        })
        self.assertEqual(details['entities'][0]['name'], name)
        self.assertEqual(details['entities'][0]['municipality'], 'Округ')

    def test_kontur_partial_pagination_never_masquerades_as_full_ranking(self):
        incomplete = kontur_table([['Округ', 'РСО', '101', '99']])
        details = source_details('gkh_kontur', {
            'ranking_tables': [incomplete], 'ranking_error': 'Вторая страница не получена',
        })
        self.assertEqual(details['entities'], [])
        self.assertEqual(details['coverage']['note'], 'Вторая страница не получена')
        self.assertTrue(is_kontur_ranking_table(incomplete))
        self.assertFalse(is_kontur_ranking_table(table(['ОМСУ', 'Номер задачи', 'Статус'], [])))
        # A first-page overview cannot silently be used as a complete ranking.
        self.assertEqual(source_details('gkh_kontur', {'tables': [incomplete]})['entities'], [])

    def test_source_failure_preserves_last_saved_measurements(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'snapshot.json'
            with patch.object(builder, 'SNAPSHOT_FILE', path):
                before = builder.build_snapshot({
                    'gkh_kontur': {'widgets': [widget('% выполнения задач ДК', '25')]},
                }, source_ids=['gkh_kontur'])
                after = builder.build_snapshot({
                    'gkh_kontur': {'error': 'Нет соединения'},
                }, source_ids=['gkh_kontur'])
                source = after['sources']['gkh_kontur']
                self.assertFalse(source['ok'])
                self.assertEqual(source['metrics'], before['sources']['gkh_kontur']['metrics'])
                self.assertEqual(source['updated_at'], before['sources']['gkh_kontur']['updated_at'])
                self.assertEqual(after['last_updated_sources'], [])


if __name__ == '__main__':
    unittest.main()
