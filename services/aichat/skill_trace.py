"""Small public activity metadata; never persist tool arguments or model reasoning."""
import re


def step(value):
    if not isinstance(value, dict):
        return None
    ident = value.get('id')
    if not isinstance(ident, str) or not re.fullmatch(r'[a-zA-Z0-9_.:-]{1,80}', ident):
        return None
    if value.get('kind') not in {'routing', 'skill', 'tool', 'model', 'attachment'}:
        return None
    if value.get('status') not in {'running', 'done', 'error'}:
        return None
    label = value.get('label')
    if not isinstance(label, str):
        return None
    result = {'type': 'step', 'id': ident, 'kind': value['kind'],
              'label': label[:180], 'status': value['status']}
    if isinstance(value.get('skill_id'), str):
        result['skill_id'] = value['skill_id'][:64]
    return result


def normalize_run(value):
    if not isinstance(value, dict):
        return None
    result = {'skills': [], 'tools': [], 'steps': [], 'warnings': []}
    for item in (value.get('skills') or [])[:2]:
        if isinstance(item, dict) and isinstance(item.get('id'), str) and isinstance(item.get('title'), str):
            result['skills'].append({'id': item['id'][:64], 'title': item['title'][:120]})
    for item in (value.get('tools') or [])[:4]:
        if isinstance(item, dict) and isinstance(item.get('name'), str):
            result['tools'].append({'name': item['name'][:80], 'title': str(item.get('title') or item['name'])[:120],
                                    'status': 'done' if item.get('status') == 'done' else 'error'})
    for item in (value.get('steps') or [])[:40]:
        safe = step(item)
        if safe:
            result['steps'].append(safe)
    for item in (value.get('warnings') or [])[:8]:
        if isinstance(item, str):
            result['warnings'].append(item[:300])
    return result
