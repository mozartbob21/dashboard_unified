"""Pure calculations. No IO, model calls or fabricated source values."""
import math
import re
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP, localcontext
from .methodology import ALL_OMSU, ZATO, K7_ITEMS, K7_RSO, REDMINE_ORIGIN, GSHEET_ID

PERIODS = ['1 квартал', '1 полугодие', '9 месяцев', '11 месяцев']
CONFIG = {
 4: dict(name='Увеличение платы за негативное воздействие', short='Негативное воздействие', query=53, metric='Выполнение плана, %', note='Методика 2026 года. Выполнение плана по каждой РСО усредняется по ОМСУ. Выполнившие план делят первое место; ЗАТО ранжируются отдельно.'),
 5: dict(name='Снижение сверхнормативных расходов', short='Сверхнормативные расходы', query=52, metric='Балл', note='Методика 2026 года. Средний балл РСО; больше — лучше. Обнуление всех задач ОМСУ означает 51 место, для ЗАТО — 5.'),
 6: dict(name='Увеличение платы за отсутствие ИПУ', short='Отсутствие ИПУ', query=54, metric='Доля ЛС без ИПУ, %', note='Условные места по исходному архиву: всем ОМСУ присвоено первое место. Данных «ЛС факт» для полного расчёта в RM нет; балл приведён для справки.'),
 7: dict(name='Снижение затрат на электроэнергию', short='Электроэнергия', metric='Выполнение дорожной карты, %', note='Предварительная оценка по свободному тексту таблицы, не официальный рейтинг. Только 9 месяцев: 17 мероприятий. Пустые ответы означают отсутствие подтверждения; неоднозначные записи требуют проверки.'),
}
PLANS = {1:[1,2,4,6],2:[1,3,6,8],3:[1,4,7,10]}
CATEGORIES = {1:'НВВ < 300 млн руб.',2:'НВВ от 300 до 600 млн руб.',3:'НВВ ≥ 600 млн руб.'}
for _ind, _roman, _direction, _scale in [(4,'IV','desc',None),(5,'V','desc',10),(6,'VI','asc',None),(7,'VII','desc',100)]:
    CONFIG[_ind].update(roman=_roman,better=_direction,scale=_scale,
        source_url=f"{REDMINE_ORIGIN}/issues?query_id={CONFIG[_ind]['query']}" if _ind<7 else
                   f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}/edit")
CONFIG[6]['metric']='Балл (доля ЛС без ИПУ), %'
CONFIG[7]['name']='Снижение затрат на оплату электроэнергии'
CONFIG[7]['metric']='Балл К7'
CONFIG[4].update(about='Критерий отражает эффективность работы РСО по взиманию платы за негативное воздействие на централизованные системы водоотведения. Характеризуется долей сбора платы за негативное воздействие в НВВ РСО на 2026 год, утверждённой Комитетом по ценам и тарифам МО.',rule='Выполнил план — 1 место, остальные — по % выполнения. Факт = Снв / НВВ, план — по категории РСО.')
CONFIG[5].update(about='Критерий отражает работу РСО по снижению сверхнормативных расходов воды: снижение потерь при транспортировке и выявление бездоговорного водопотребления. Балл от 0 до 10; если у ОМСУ несколько РСО/задач — берётся средний балл.',rule='Больше балл — выше место; несколько РСО — среднее. Балл обнулён — 51 место (ЗАТО — 5).')
CONFIG[6].update(about='Цель — стимулировать РСО к полному учёту ресурсов и установке ИПУ холодной/горячей воды либо применению повышающего коэффициента 3 при расчётах по нормативу.',rule='Всем ОМСУ — 1 место, как в исходном своде: ЛСфакт из методики в RM нет. Балл — для справки.')
CONFIG[7].update(about='Снижение затрат на оплату электроэнергии РСО за счёт выполнения мероприятий «дорожной карты» по переводу точек учёта на оптимальную ценовую категорию. Данные — из Google-таблицы (лист «Исполнение дорожной карты»).',rule='Балл РСО = выполнено мероприятий дорожной карты / план × 100, ОМСУ — среднее. По ответам РСО в таблице, официально ещё не оценивался.')

def norm(value): return str(value or '').upper().replace('Ё','Е').strip()
def source_string(value):
    """Match String(value) for the JSON values used by Redmine and XLSX."""
    if value is True:return 'true'
    if value is False:return 'false'
    if isinstance(value,float) and math.isfinite(value) and value.is_integer():return str(int(value))
    if isinstance(value,list):return ','.join('' if v is None else source_string(v) for v in value)
    if isinstance(value,dict):return '[object Object]'
    return str(value)
def num(value):
    if value is None or isinstance(value, bool): return None
    try:
        text=source_string(value).strip().strip('\ufeff').strip().replace(',','.',1)
        if not text:return None
        if re.fullmatch(r'0[xX][0-9a-fA-F]+|0[bB][01]+|0[oO][0-7]+',text):n=float(int(text,0))
        elif re.fullmatch(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?',text):n=float(text)
        else:return None
        return n if math.isfinite(n) else None
    except (ValueError, TypeError, OverflowError): return None

def rnd(n):
    """Math.round(n * 100) / 100, including floats immediately below a half."""
    scaled=n*100
    if not math.isfinite(scaled):return scaled
    floor=math.floor(scaled)
    return (floor+(scaled-floor>=0.5))/100
def cf(issue, ident):
    return next((f.get('value') for f in issue.get('custom_fields',[]) if isinstance(f,dict) and f.get('id')==ident),None)
def group(omsu): return 'zato' if norm(omsu) in ZATO else 'main'
def zeroed(issue,p): return p>0 and source_string(cf(issue,[None,259,260,261][p])) == [None,'337','339','341'][p]
def rank(rows, offset=0):
    ordered=sorted(rows,key=lambda r:-r['value'])
    for n,row in enumerate(ordered): row['place']=ordered[n-1]['place'] if n and abs(ordered[n-1]['value']-row['value'])<1e-9 else offset+n+1

def rating(ind,p,issues):
    by=defaultdict(list)
    for i in issues:
        omsu=str((i.get('assigned_to') or {}).get('name') or '').strip()
        if omsu and num(cf(i,17)) is not None: by[omsu].append(i)
    rows=[]
    for omsu, entries in by.items():
        row=dict(omsu=omsu,group=group(omsu),no_docs=False,details=[],extra={})
        if ind==4:
            for i in entries:
                for name,amount,nvv in [(cf(i,70) or i.get('subject','РСО'),cf(i,165+p),cf(i,169)),(cf(i,208) or 'РСО 2',cf(i,209+p),cf(i,213))]:
                    c,n=num(amount),num(nvv)
                    if c is None or not n:continue
                    cat=1 if n<300000 else 2 if n<600000 else 3
                    fact=rnd(c/n*100);plan=PLANS[cat][p]
                    row['details'].append(dict(name=str(name),amount=c,nvv=n,fact=fact,category=cat,category_label=CATEGORIES[cat],plan=plan,value=rnd(fact/plan*100)))
            rs=row['details']
            if not rs:continue
            row['value']=rnd(sum(x['value'] for x in rs)/len(rs))
            row['extra']={'fact':rnd(sum(x['amount'] for x in rs)/sum(x['nvv'] for x in rs)*100),'plan':list(dict.fromkeys(x['plan'] for x in rs))}
        else:
            vals=[]
            for i in entries:
                v=0 if ind==5 and zeroed(i,p) else num(cf(i,158+p if ind==5 else [189,195,201,207][p]))
                row['details'].append(dict(name=str(i.get('subject') or 'Задача РСО'),value=v,zeroed=ind==5 and zeroed(i,p),issue_id=i.get('id')))
                if v is not None:vals.append(v)
            if not vals:continue
            row['value']=sum(vals)/len(vals)  # Do not round before ranking K5.
            if ind==5:row['no_docs']=all(zeroed(i,p) for i in entries)
            else:
                start=[184,190,196,202][p]
                total=sum(num(cf(i,start)) or 0 for i in entries)
                parts=[sum(num(cf(i,k)) or 0 for i in entries) for k in range(start+1,start+5)]
                without=sum(num(cf(i,k)) or 0 for i in entries for k in range(start+1,start+5))
                if total:row['extra']={'total':total,'without':without,'parts':parts}
        row['rso_names']=[d['name'] for d in row['details']] if ind==4 else list(dict.fromkeys(str(i.get('subject')) for i in entries if i.get('subject'))) if ind==5 else []
        rows.append(row)
    for g in ('main','zato'):
        grp=[r for r in rows if r['group']==g]
        ok=[r for r in grp if not r['no_docs']]
        for r in grp:
            if r['no_docs']:r['place']=5 if g=='zato' else 51
        if ind==4:
            achieved=[r for r in ok if r['value']>=100]
            for r in achieved:r['place']=1
            rank([r for r in ok if r['value']<100],len(achieved))
        elif ind==6:
            for r in ok:r['place']=1
        else:rank(ok)
    rows=sorted(rows,key=lambda r:(r['group']=='zato',r['place'],-r['value'] if ind!=6 else r['value']))
    return decorate_period(ind,p,rows,[])

def k7cell(value,item,previous):
    if value is None or not source_string(value).strip():return {'ok':False,'reason':'не заполнено'}
    s=re.sub(r'\s+',' ',source_string(value).lower()).strip(); n=item['n']
    if re.fullmatch(r'[0-9]+(?:[.,][0-9]+)?',s):
        number=float(s.replace(',','.'))
        if item.get('pct'):return {'ok':True} if number>=100 else dict(ok=False,reason=f'выполнено на {fmt(number,0)} %')
        if number>0:return {'ok':True}
        if n>=15:return dict(ok=True,reason='«0» — несоответствующих ПУ нет')
        return dict(ok=previous,reason='«0» — целесообразных точек нет' if previous else '«0», а предыдущий шаг не выполнен')
    if re.search(r'[0-9]+\s*%',s):return dict(ok=False,reason='выполнено частично')
    # JavaScript's \b tests ASCII word characters: the Cyrillic у itself is
    # non-word, so a boundary after it exists only before an ASCII word char.
    if n==15 and re.search(r'[0-9]+\s*(точ|при|пу(?=[A-Za-z0-9_]))',s):return {'ok':True}
    if n==16 and re.search(r'направлен|заявк|исх|№|и-[0-9][0-9]-',s) and not re.search(r'не\s*направл|будет',s):return {'ok':True}
    if re.search(r'не\s*выполн|в работе|буд(ет|ут)|не\s*направл|проводится|вед[её]тся|выполняется|в стадии|на стадии|не\s*получ|на этапе|ожида|осталось|осталась|в ноябре|в сентябре|некоррект',s):return dict(ok=False,reason='не завершено')
    if re.search(r'выполн|получен|направ|не\s*треб|не\s*целесообраз|не\s*выявл|в наличии|данные полные|данные есть|сформирован|проведен|выявлен|заменен|устан|запрос|исх|вх\.|заявк|письм|№|и-[0-9][0-9]-|ответ|экономи|истек мпи|^да$|точ(ек|ки|ка) уч',s):return {'ok':True}
    if n==17:return dict(ok=False,reason='нет подтверждения, что замена ПУ выполнена')
    return dict(ok=False,reason='неясная запись',unclear=True)

def rating7(sheet):
    plan=K7_ITEMS[:17];candidates=defaultdict(list);unmatched=0
    for row in sheet:
        name=str(row[2] or '').strip() if len(row)>2 else ''
        if not name:continue
        slot=next((i for i,x in enumerate(K7_RSO) if re.search(x['pattern'],name,re.I if x['ignorecase'] else 0) and (not x['where'] or (len(row)>3 and norm(row[3])==x['where']))),None)
        if slot is None:unmatched+=1;continue
        items=[]
        for item in plan:
            raw=row[item['col']] if len(row)>item['col'] else None
            evaluation=k7cell(raw,item,items[-1]['ok'] if items else True)
            items.append(dict(item,raw=('' if raw is None else source_string(raw)),**evaluation))
        candidates[slot].append(dict(items=items,sheet_name=name,sheet_omsu=str(row[3] or '').strip() if len(row)>3 else ''))
    by=defaultdict(list)
    for index,slot in enumerate(K7_RSO):
        options=candidates[index]
        best=max(options,key=lambda candidate:sum(bool(x['ok']) for x in candidate['items'])) if options else None
        items=best['items'] if best else [dict(x,raw='',ok=False,reason='РСО нет в таблице') for x in plan]
        done=sum(bool(x['ok']) for x in items)
        by[slot['omsu']].append(dict(name=slot['rso'],value=rnd(done/17*100),done=done,plan=17,found=bool(options),items=items,duplicates=len(options),sheet_name=best['sheet_name'] if best else None,sheet_omsu=best['sheet_omsu'] if best else None))
    rows=[dict(omsu=o,group=group(o),value=rnd(sum(x['value'] for x in rs)/len(rs)),details=rs,rso_names=[x['name'] for x in rs],no_docs=False,extra={'done':sum(x['done'] for x in rs),'plan':17*len(rs)}) for o,rs in by.items()]
    for g in ('main','zato'):rank([r for r in rows if r['group']==g])
    rows=sorted(rows,key=lambda r:(r['group']=='zato',r['place'],r['omsu']))
    return decorate_period(7,2,rows,[]),unmatched

def fmt(value, digits=2):
    """Russian display values; never round the value used for ranking."""
    if value is None:return '—'
    decimal=Decimal(str(value))
    if not decimal.is_finite():return '—'
    with localcontext() as context:
        context.prec=max(28,decimal.adjusted()+digits+3)
        rounded=decimal.quantize(Decimal(1).scaleb(-digits),rounding=ROUND_HALF_UP)
    return f'{rounded:,.{digits}f}'.replace(',','\u00a0').replace('.',',')


def place_tip(ind,row,rows):
    peers=[r for r in rows if r['group']==row['group']]
    ties=sum(r['place']==row['place'] for r in peers)
    text=f"{row['place']} место"+(' среди ЗАТО (рейтинг отдельно, 1–5)' if row['group']=='zato' else '')
    if row['no_docs']:
        text+=f"\nБалл обнулён (документы не представлены) — последнее место: {5 if row['group']=='zato' else 51}."
    elif ind==4:
        text+= ('\nПлан выполнен — по методике 1 место у всех выполнивших.' if row['value']>=100 else
                f"\nПлан не выполнен: места идут после {sum(r['value']>=100 and not r['no_docs'] for r in peers)} выполнивших, по убыванию % выполнения.")
    elif ind==6:text+='\nПо методике посчитать нельзя (нет ЛСфакт в RM) — как и в исходном рейтинге, у всех 1 место.'
    if ties>1 and not(ind==6 or ind==4 and row['value']>=100):text+=f'\nДелит место ещё с {ties-1} ОМСУ — одинаковое значение.'
    return text


def value_tip(ind,row,p):
    entries=row['details'];per=PERIODS[p]
    if ind==4:
        parts=[]
        for d in entries:
            parts.append((d['name']+':\n' if len(entries)>1 else '')+
                f"Факт = С{p+1}нв / НВВ × 100 = {fmt(d['amount'])} / {fmt(d['nvv'])} × 100 = {fmt(d['fact'])} %\n"+
                f"Категория {d['category']} ({d['category_label']}) → план за {per}: {d['plan']} %\n"+
                f"Выполнение = {fmt(d['fact'])} / {d['plan']} × 100 = {fmt(d['value'])} %")
        if len(entries)>1:parts.append(f"По ОМСУ — среднее: ({' + '.join(fmt(d['value']) for d in entries)}) / {len(entries)} = {fmt(row['value'])} %")
        return '\n\n'.join(parts)+('\n\nПлан выполнен → 1 место.' if row['value']>=100 else '\n\nПлан не выполнен.')
    if ind==5:
        zero='\n\nВ задаче RM отмечено «Обнулить балл» — балл засчитан как 0.' if any(d['zeroed'] for d in entries) else ''
        if len(entries)==1:return f"Балл за {per} из задачи RM: {fmt(row['value'])} (из 10 возможных)."+zero
        used=[d for d in entries if d['value'] is not None]
        return f"Средний балл по {len(used)} РСО/задачам:\n"+'\n'.join(f"  {d['name']}: {fmt(d['value'])}" for d in used)+f"\n\n({' + '.join(fmt(d['value']) for d in used)}) / {len(used)} = {fmt(row['value'])}"+zero
    if ind==6:
        ex=row['extra']
        if not ex:return f"Балл за {per} из задачи RM: {fmt(row['value'])} %."
        calc=ex['without']/ex['total']*100
        return f"Балл = ЛС без ИПУ / всего ЛС × 100\n= {fmt(ex['without'],0)} / {fmt(ex['total'],0)} × 100 = {fmt(calc)} %"+(f"\n\nВ задаче RM указан балл {fmt(row['value'])} %." if abs(calc-row['value'])>0.01 else '')+'\n\nНа место не влияет: в исходной методике у всех 1 место.'
    text=f'Балл РСО = выполнено мероприятий / плановых за {per} × 100\n'+'\n'.join((d['name']+': ' if len(entries)>1 else '')+f"{d['done']} / {d['plan']} × 100 = {fmt(d['value'])}"+(' (РСО нет в таблице)' if not d['found'] else '') for d in entries)
    if len(entries)>1:text+=f"\n\nК7 = среднее по {len(entries)} РСО: ({' + '.join(fmt(d['value']) for d in entries)}) / {len(entries)} = {fmt(row['value'])}"
    return text


def decorate_period(ind,p,rows,previous_rows):
    """Presentation rules from the archive, calculated before any UI filtering.

    Color is the third of the full group by place, not score. Ties retain the
    same color; the original deliberately does not force tiny groups green.
    """
    prev={r['omsu']:r['place'] for r in previous_rows}
    maximum=max([r['value'] for r in rows]+[CONFIG[ind]['scale'] or 0]) or 1
    for row in rows:
        peers=[r for r in rows if r['group']==row['group']]
        count=len(peers);ties=sum(r['place']==row['place'] for r in peers)
        row['group_size']=count
        row['place_label']=str(row['place']) if row['no_docs'] or ties==1 else f"{row['place']}–{row['place']+ties-1}"
        row['zone']='z1' if row['place']<=count/3 else 'z2' if row['place']<=2*count/3 else 'z3'
        bar=row['value']/maximum if CONFIG[ind]['better']=='desc' else 1-row['value']/maximum
        row['bar_pct']=float(Decimal.from_float(float(max(2,min(100,bar*100)))).quantize(Decimal('.1'),rounding=ROUND_HALF_UP))
        row['tag']='zeroed' if row['no_docs'] else ('achieved' if row['value']>=100 else 'not_achieved') if ind==4 else None
        row['previous_place']=prev.get(row['omsu'])
        row['delta']=None if row['previous_place'] is None else row['previous_place']-row['place']
        row['place_tooltip']=place_tip(ind,row,rows)
        row['value_tooltip']=value_tip(ind,row,p)
        if ind==4:
            entries=row['details']
            row['extra']['fact_tooltip']='\n'.join((d['name']+': ' if len(entries)>1 else '')+f"С{p+1}нв / НВВ × 100 = {fmt(d['amount'])} / {fmt(d['nvv'])} × 100 = {fmt(d['fact'])} %" for d in entries)+(f"\nВместе: ΣС / ΣНВВ × 100 = {fmt(row['extra']['fact'])} %" if len(entries)>1 else '')+f'\n\nС{p+1}нв — сбор платы за негативное воздействие нарастающим итогом с начала года, тыс. руб. (так он заполнен в RM).'
            row['extra']['plan_tooltip']='\n'.join((d['name']+': ' if len(entries)>1 else '')+f"НВВ {fmt(d['nvv'])} тыс. руб. → категория {d['category']} → план {d['plan']} %" for d in entries)+f'\n\nПлан за {PERIODS[p]} по методике:\n'+ '\n'.join(f'{cat} кат. ({CATEGORIES[cat]}): {PLANS[cat][p]} %' for cat in PLANS)
        elif ind==6 and row['extra']:
            ex=row['extra'];ex['without_tooltip']='ЛС без ИПУ = ЛС1 + ЛС2 + ЛС3 + ЛС4\n= '+' + '.join(fmt(v,0) for v in ex['parts'])+f" = {fmt(ex['without'],0)}"
    return rows


def period_stats(ind,rows,previous_rows):
    have={norm(r['omsu']) for r in rows}
    return dict(count=len(rows),zato_count=sum(r['group']=='zato' for r in rows),
                average=sum(r['value'] for r in rows)/len(rows) if rows else None,
                achieved_count=sum(r['value']>=100 for r in rows) if ind==4 else None,
                has_previous=bool(previous_rows),
                nonparticipants=[dict(omsu=o,group=group(o),na=True) for o in ALL_OMSU if o not in have])


def build(sources):
    indicators={}
    for ind in (4,5,6,7):
        src=sources.get(str(ind),{})
        periods=[rating(ind,p,src['data']) if ind<7 else rating7(src['data'])[0] if p==2 else [] for p in range(4)] if 'data' in src else [[],[],[],[]]
        for p,rows in enumerate(periods):decorate_period(ind,p,rows,periods[p-1] if p else [])
        counts=[len(rs) for rs in periods];maximum=max(counts)
        default=next((p for p in range(3,-1,-1) if counts[p] and counts[p]>=maximum/2),None)
        indicators[str(ind)]={**CONFIG[ind],'periods':periods,'period_stats':[period_stats(ind,rows,periods[p-1] if p else []) for p,rows in enumerate(periods)],'default_period':default,'updated_at':src.get('updated_at'),'error':src.get('error'),'provisional':ind in (6,7)}
    return dict(methodology_year=2026,period_labels=PERIODS,municipalities=ALL_OMSU,indicators=indicators)
