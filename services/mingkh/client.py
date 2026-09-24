"""MINZHKH CDA queries ported from the supplied dashboard, with verified TLS."""
import ssl
import os
from datetime import date, timedelta
import httpx

BASE = 'https://cur.bi.mosreg.ru/pentaho/plugin/cda/api/doQuery'
CDA_PATH = '/public/ЕЦУР_v3/ЕЦУР_api.cda'
DEFAULT_FILTERS = {key:'-1' for key in ('message_type','source','direction','synt_group','subtopic','fact','omsu','department')}
DEFAULT_FILTERS.update(curator='17',population_limit='all',spammers='0')
QUERIES = {'q_download_detalization','q_map','q_last_updated'}


class PortalError(RuntimeError):
    def __init__(self, message, *, retryable=False):
        super().__init__(message)
        self.retryable = retryable


def year_ago(days, today=None):
    today=today or date.today()
    target=today-timedelta(days=days)
    try:
        target=target.replace(year=target.year-1)
    except ValueError:
        target=target.replace(year=target.year-1,day=28)
    return (today-target).days


def periods(cs,ce,ps,pe):
    result=dict(DEFAULT_FILTERS)
    for key,value in {'curr_period_start':cs,'curr_period_end':ce,'prev_period_start':ps,'prev_period_end':pe,
                      'appg_period_start':year_ago(cs),'appg_period_end':year_ago(ce),
                      'start_periods':cs,'end_periods':ce,'appg_start_periods':year_ago(cs),'appg_end_periods':year_ago(ce)}.items():
        result[key]=f'{value} day'
    result.update({key:'-1' for key in ('chart_direction_id','chart_group_id','chart_fact_id','chart_subtopic_id','direction_id','synt_group_id')})
    return result


class Pentaho:
    def __init__(self,username,password):
        self.username=username
        self.password=password

    def query(self,query,params=None,*,timeout=180):
        if query not in QUERIES:
            raise PortalError('Неизвестный запрос к порталу.')
        request={'path':CDA_PATH,'dataAccessId':query,'outputType':'json'}
        request.update({'param'+k:v for k,v in (params or {}).items()})
        try:
            ca=os.getenv('MINGKH_CA_BUNDLE') or None
            context=ssl.create_default_context(cafile=ca)
            with httpx.Client(auth=(self.username,self.password),verify=context,trust_env=False,
                              follow_redirects=False,timeout=httpx.Timeout(timeout,connect=15)) as client:
                with client.stream('GET',BASE,params=request) as response:
                    if response.status_code in (401,403) or response.is_redirect:
                        raise PortalError('Портал не принял логин и пароль. Администратору нужно проверить доступ МИНЖКХ.')
                    if response.status_code!=200:
                        raise PortalError(f'Портал временно недоступен (HTTP {response.status_code}).',
                                          retryable=response.status_code == 429 or response.status_code >= 500)
                    payload=bytearray()
                    for part in response.iter_bytes():
                        payload.extend(part)
                        if len(payload)>40*1024**2:
                            raise PortalError('Выгрузка слишком большая. Выберите меньший период.')
            import json
            result=json.loads(payload)
            columns=[str(c['colName']) for c in result['metadata']]
            rows=result['resultset']
            if not isinstance(rows,list) or len(rows)>300000:
                raise PortalError('Слишком много обращений. Выберите меньший период.')
            return columns,rows
        except PortalError:
            raise
        except (httpx.TimeoutException, httpx.NetworkError) as exc:
            cause = exc
            seen = set()
            while cause is not None and id(cause) not in seen:
                seen.add(id(cause))
                if isinstance(cause, ssl.SSLError):
                    raise PortalError('Не удалось проверить защищённое соединение с cur.bi.mosreg.ru. Проверьте сертификаты на компьютере-сервере.') from None
                cause = cause.__cause__ or cause.__context__
            raise PortalError('Соединение с порталом прервалось. Повторите загрузку.', retryable=True) from None
        except (httpx.TransportError,OSError):
            raise PortalError('Нет защищённого соединения с cur.bi.mosreg.ru. Проверьте доступ с компьютера-сервера и сертификаты.')
        except (ValueError,KeyError,TypeError):
            raise PortalError('Портал вернул неожиданный формат данных. Проверьте доступ и повторите.')

    def last_updated(self):
        _,rows=self.query('q_last_updated')
        if not rows or not rows[0]:
            return 'Нет данных',None
        return str(rows[0][0]),rows[0][1] if len(rows[0])>1 else None
