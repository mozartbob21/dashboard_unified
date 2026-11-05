"""Numeric indicators from sources actually refreshed in this dashboard run."""
import math

from services.water_dashboard.config import SOURCES
from services.water_dashboard.metrics import source_metrics

DEFINITIONS = {
    source['id'] + '_' + metric['id']: (source['name'] + ' · ' + metric['label'], metric['unit'])
    for source in SOURCES for metric in source_metrics(source['id'], {}, [])
}


def project(payload, params):
    checked = payload.get('last_checked_sources')
    updated = payload.get('last_updated_sources')
    if not isinstance(checked, list) or not isinstance(updated, list) or not checked:
        return {}
    selected = params.get('source')
    if selected and selected != 'all' and set(checked) != {selected}:
        return {}
    result = {'sources': len(checked), 'failed_sources': len(set(checked) - set(updated))}
    for key in set(checked) & set(updated):
        item = payload.get('sources', {}).get(key, {})
        if item.get('ok') is not True or item.get('metric_schema') != 1:
            continue
        for metric in item.get('metrics', []):
            name = key + '_' + str(metric.get('id'))
            value = metric.get('value')
            if (name in DEFINITIONS and metric.get('unit') == DEFINITIONS[name][1]
                    and type(value) in (float, int) and math.isfinite(value)):
                result[name] = value
    return result
