"""Reuse the platform's verified Pentaho connection and stored admin credentials."""
import re
from services.auth.integrations import credentials
from services.mingkh.client import Pentaho as BasePentaho, PortalError, periods as base_periods


class Pentaho(BasePentaho):
    QUERIES = frozenset({
        'q_top_subtopic', 'q_omsu', 'q_departments', 'q_last_updated',
        'q_download_detalization', 'q_download_detalization_address', 'q_map',
    })
    FILTERS = frozenset({'source', 'omsu', 'department', 'direction', 'synt_group', 'subtopic', 'fact'})

    def __init__(self):
        value = credentials('mingkh')
        if not value or not value.get('username') or not value.get('password'):
            raise PortalError('Администратору нужно сохранить доступ к МинЖКХ в разделе «Логины и пароли».')
        super().__init__(value['username'], value['password'])

    @staticmethod
    def periods(curr_start, curr_end, prev_start, prev_end, **overrides):
        if any(type(value) is not int or not 0 <= value <= 15000
               for value in (curr_start, curr_end, prev_start, prev_end)):
            raise ValueError('Некорректный период запроса к порталу.')
        if curr_start < curr_end or prev_start < prev_end:
            raise ValueError('Некорректный порядок дат периода.')
        if set(overrides) - Pentaho.FILTERS:
            raise ValueError('Неизвестный фильтр портала.')
        for value in overrides.values():
            if not re.fullmatch(r'-1|[0-9]{1,10}', str(value)):
                raise ValueError('Некорректное значение фильтра портала.')
        result = base_periods(curr_start, curr_end, prev_start, prev_end)
        result.update({key: str(value) for key, value in overrides.items()})
        return result
