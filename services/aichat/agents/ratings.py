"""Authorized chat entry point for the reviewed complaints-neighbors-rating agent."""
from __future__ import annotations

import asyncio
import hashlib
import math
import re
import threading
from urllib.parse import quote

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import JSONResponse, Response

from core.roles import check_module_access
from services.aichat import jobs, private_storage
from services.aichat.job_api import identity
from services.auth.home_preferences import account_key
from utils.db import get_db_connection

ID = 'complaints-neighbors-rating'
MODULE = 'mingkh'
TITLE = 'Рейтинг домов по жалобам и соседям'
DESCRIPTION = 'Ранжирует дома по числу жалоб, находит ближайшие дома с жалобами и создаёт отчёт Excel.'
FILENAME = 'Рейтинг домов по жалобам и соседям.xlsx'
MAX_INPUT = 8 * 1024 * 1024
MAX_OUTPUT = 16 * 1024 * 1024
_CALCULATION_SLOT = threading.Semaphore(1)
PRIVATE = {'Cache-Control':'private, no-store'}
router = APIRouter(prefix='/aichat/api/agents', tags=['aichat'])


def require_access(user):
    if not user or not user.get('is_active',True):
        raise HTTPException(401,'Требуется авторизация.')
    if not check_module_access(user,MODULE):
        raise HTTPException(403,'Для рейтинга домов требуется право на блок жалоб МинЖКХ.')


@router.get('')
def catalog(user=Depends(identity)):
    agents = [{'id':ID,'title':TITLE,'description':DESCRIPTION}] if check_module_access(user,MODULE) else []
    return JSONResponse({'agents':agents},headers=PRIVATE)


def publish_file(conn, row, content, requirements):
    """Called inside the same transaction as the final job status and chat reply."""
    if MODULE not in requirements or not isinstance(content,bytes) or not 0<len(content)<=MAX_OUTPUT:
        raise ValueError('Не удалось сохранить итоговый файл агента.')
    conn.execute('INSERT INTO aichat_agent_artifacts(id,dialog_id,contents) VALUES (?,?,?)',
                 (row['id'],row['dialog_id'],content))
    return {'id':row['id'],'filename':FILENAME}


@router.get('/artifacts/{ident}')
def download(ident: str, user=Depends(identity)):
    require_access(user)
    if not re.fullmatch(r'[0-9a-f]{32}',ident):
        raise HTTPException(404,'Файл не найден.')
    # Reuse the exact same permission projection as the result in the shared chat.
    ticket=jobs.get(user,ident)
    artifact=(ticket.get('result') or {}).get('artifact') or {}
    if ticket['status']!='succeeded' or artifact.get('id')!=ident:
        raise HTTPException(404,'Готового файла нет.')
    with get_db_connection() as conn:
        row=conn.execute('SELECT contents FROM aichat_agent_artifacts WHERE id=? AND dialog_id=?',
                         (ident,ticket['dialog_id'])).fetchone()
    if not row:
        raise HTTPException(404,'Файл не найден.')
    return Response(row['contents'],media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={**PRIVATE,'Content-Disposition':"attachment; filename*=UTF-8''"+quote(FILENAME),
                             'X-Content-Type-Options':'nosniff'})


def _text(form, key, default='', limit=240):
    values=form.getlist(key)
    if not values:
        return default
    if len(values)!=1 or not isinstance(values[0],str) or len(values[0])>limit:
        raise HTTPException(422,'Проверьте значения полей формы агента.')
    return values[0].strip()


def _radius(form, key, default):
    value=_text(form,key,str(default))
    if not re.fullmatch(r'[0-9]{1,6}',value) or not 1<=int(value)<=50_000:
        raise HTTPException(422,'Радиус должен быть целым числом от 1 до 50 000 метров.')
    return int(value)


def _options(form):
    city=_text(form,'agent_city','Чехов',80)
    if not re.fullmatch(r'[а-яёА-ЯЁa-zA-Z][а-яёА-ЯЁa-zA-Z0-9 .-]{0,79}',city):
        raise HTTPException(422,'Укажите название округа.')
    bounds=_text(form,'agent_bbox')
    bbox=None
    if bounds:
        try:
            bbox=tuple(float(v.strip()) for v in bounds.split(','))
            west,east,south,north=bbox
            if not all(math.isfinite(v) for v in bbox) or not (-180<=west<east<=180 and -90<=south<north<=90):
                raise ValueError()
        except ValueError:
            raise HTTPException(422,'Границы должны задавать прямоугольник: долгота от, долгота до, широта от, широта до.') from None
    elif city.casefold()!='чехов':
        raise HTTPException(422,'Для другого округа укажите его границы. Значения по умолчанию относятся к Чехову.')
    radius=_radius(form,'agent_radius',100)
    wide=_radius(form,'agent_radius_wide',200)
    if wide<=radius:
        raise HTTPException(422,'Второй радиус должен быть больше первого.')
    return {'city':city,'bbox':bbox,'radius':radius,'radius_wide':wide}


async def _book(form,key,label):
    values=form.getlist(key)
    if len(values)!=1 or not hasattr(values[0],'read') or not str(getattr(values[0],'filename','')).casefold().endswith('.xlsx'):
        raise HTTPException(422,'Приложите книгу XLSX в поле «'+label+'».')
    raw=await values[0].read(MAX_INPUT+1)
    if len(raw)>MAX_INPUT:
        raise HTTPException(413,'Каждая книга должна быть не больше 8 МБ.')
    if not raw or not raw.startswith(b'PK\x03\x04'):
        raise HTTPException(422,'Файл «'+label+'» не является книгой XLSX.')
    return raw


def _preview_address(address):
    text = ' '.join(address.split())
    return text if len(text) <= 240 else text[:239] + '…'


def _description(summary):
    stats=summary['stats']
    lines=[f"{summary['city']} · Рейтинг домов по жалобам",
           f"Даты строк ЕДС: {stats['period_eds']}. Даты строк ЕЦУР: {stats['period_ecur']}.",
           f"Строк ЕДС: {stats['rows_eds']}. Строк ЕЦУР: {stats['rows_ecur']}. "
           f"Исключено помеченных дублей: {stats['dup_flagged']}, скрытых: {stats['dup_hidden']}.",
           f"Уникальных жалоб: {stats['unique_complaints']}. К домам привязано: {stats['in_rating']}, "
           f"без привязки: {stats['unplaced']}. В рейтинге домов: {stats['houses']}.",
           f"Расстояние измерено по прямой между центрами домов. Радиус «рядом» принят равным {summary['radius']} м, "
           f"дополнительный радиус — {summary['radius_wide']} м."]
    if summary['top10']:
        lines+=['','Дома с наибольшим числом жалоб:']
        for row in summary['top10'][:10]:
            distance='расстояние не определено' if row.get('distance') is None else f"до ближайшего дома с жалобами {row['distance']} м"
            lines.append(f"{row['rank']}. {_preview_address(row['address'])} — {row['count']} жалоб; {distance}.")
    if stats['bad_point_houses']:
        lines.append(f"Домов без проверенных координат: {stats['bad_point_houses']}. Они сохранены в рейтинге, но исключены из расчёта соседей.")
        for row in summary['bad_points'][:10]:
            lines.append(f"• {_preview_address(row['address'])} — {row['count']} жалоб.")
        if stats['bad_point_houses']>10:
            lines.append('Полный перечень отмечен в книге.')
    lines.append('Часть заявок о тепловых сетях может быть привязана к дому по упоминанию в соцсетях. Для анализа отопления используйте отдельный лист «Только отопление».')
    return '\n'.join(lines)


async def submit(request: Request, form):
    user=request.state.user
    require_access(user)
    if _text(form,'agent_id')!=ID:
        raise HTTPException(404,'Такой агент недоступен.')
    if _text(form,'background')!='true' or form.getlist('file'):
        raise HTTPException(422,'Запустите расчёт из раздела «ИИ-агенты», указав две книги XLSX.')
    options=_options(form)
    geo=await _book(form,'geo_file','Справочник координат')
    complaints=await _book(form,'complaints_file','Книга жалоб')
    if geo==complaints:
        raise HTTPException(422,'Укажите две разные книги: справочник координат и книгу жалоб.')
    dialog_id=_text(form,'dialog_id',limit=64)
    fingerprint={'agent':ID,'dialog_id':dialog_id,'options':options,
                 'geo':hashlib.sha256(geo).hexdigest(),'complaints':hashlib.sha256(complaints).hexdigest()}
    ticket,fresh=jobs.submit(user,_text(form,'request_id',limit=100),fingerprint,
                             dialog_id=dialog_id,title=TITLE,required_modules=[MODULE])
    if fresh:
        async def work(publish):
            from services.aichat.agents.neighbors_rating import build_rating, RatingError
            require_access(jobs.current_user(user))
            private_storage.append_message(ticket['dialog_id'],'user',
                f"Рассчитать рейтинг домов по жалобам: {options['city']}. Радиусы {options['radius']} и {options['radius_wide']} м.",
                owner=account_key(user),event_id=ticket['job_id']+':user',required_modules=[MODULE])
            publish({'id':'rating','kind':'tool','label':'Расчёт принят в очередь','status':'running'})
            def calculate():
                # Bound memory used by open workbooks and pairwise distance calculations.
                with _CALCULATION_SLOT:
                    require_access(jobs.current_user(user))
                    if not jobs.alive(ticket['job_id']):
                        raise RatingError('Срок выполнения задачи истёк. Запустите новый расчёт.')
                    publish({'id':'rating','kind':'tool','label':'Сопоставляю адреса и рассчитываю рейтинг','status':'running'})
                    return build_rating(geo,complaints,**options)
            try:
                workbook,summary=await asyncio.to_thread(calculate)
            except RatingError as error:
                publish({'id':'rating','kind':'tool','label':'Расчёт не завершён','status':'error'})
                return {'ok':False,'response_kind':'error','required_modules':[MODULE],'answer':str(error)}
            require_access(jobs.current_user(user))
            publish({'id':'rating','kind':'tool','label':'Рейтинг рассчитан, книга проверена','status':'done'})
            return {'ok':True,'required_modules':[MODULE],'response_kind':'answer','answer':_description(summary),
                    '_agent_workbook':workbook}
        jobs.start(ticket,work)
    return JSONResponse(ticket,status_code=202,headers=PRIVATE)
