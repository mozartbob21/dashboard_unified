"""Pure calculations. No IO, model calls or fabricated source values."""
import math
import re
from collections import defaultdict
from .methodology import ALL_OMSU, ZATO, K7_ITEMS, K7_RSO

PERIODS = ['1 квартал', '1 полугодие', '9 месяцев', '11 месяцев']
CONFIG = {
 4: dict(name='Увеличение платы за негативное воздействие', short='Негативное воздействие', query=53, metric='Выполнение плана, %', note='Методика 2026 года. Выполнение плана по каждой РСО усредняется по ОМСУ. Выполнившие план делят первое место; ЗАТО ранжируются отдельно.'),
 5: dict(name='Снижение сверхнормативных расходов', short='Сверхнормативные расходы', query=52, metric='Балл', note='Методика 2026 года. Средний балл РСО; больше — лучше. Обнуление всех задач ОМСУ означает 51 место, для ЗАТО — 5.'),
 6: dict(name='Увеличение платы за отсутствие ИПУ', short='Отсутствие ИПУ', query=54, metric='Доля ЛС без ИПУ, %', note='Условные места по исходному архиву: всем ОМСУ присвоено первое место. Данных «ЛС факт» для полного расчёта в RM нет; балл приведён для справки.'),
 7: dict(name='Снижение затрат на электроэнергию', short='Электроэнергия', metric='Выполнение дорожной карты, %', note='Предварительная оценка по свободному тексту таблицы, не официальный рейтинг. Только 9 месяцев: 17 мероприятий. Пустые ответы означают отсутствие подтверждения; неоднозначные записи требуют проверки.'),
}
PLANS = {1:[1,2,4,6],2:[1,3,6,8],3:[1,4,7,10]}

def norm(value): return str(value or '').upper().replace('Ё','Е').strip()
def num(value):
    if value is None or isinstance(value, bool): return None
    try:
        text=str(value).strip().replace(',','.')
        if not text:return None
        n=float(text)
        return n if math.isfinite(n) else None
    except (ValueError, TypeError): return None

def rnd(n): return math.floor(n*100+0.5)/100  # Match Math.round in supplied methodology.
def cf(issue, ident):
    return next((f.get('value') for f in issue.get('custom_fields',[]) if isinstance(f,dict) and f.get('id')==ident),None)
def group(omsu): return 'zato' if norm(omsu) in ZATO else 'main'
def zeroed(issue,p): return p>0 and str(cf(issue,[None,259,260,261][p])) == [None,'337','339','341'][p]
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
                    row['details'].append(dict(name=str(name),amount=c,nvv=n,fact=fact,plan=plan,value=rnd(fact/plan*100)))
            rs=row['details']
            if not rs:continue
            row['value']=rnd(sum(x['value'] for x in rs)/len(rs))
            row['extra']={'fact':rnd(sum(x['amount'] for x in rs)/sum(x['nvv'] for x in rs)*100),'plan':sorted(set(x['plan'] for x in rs))}
        else:
            vals=[]
            for i in entries:
                v=0 if ind==5 and zeroed(i,p) else num(cf(i,158+p if ind==5 else [189,195,201,207][p]))
                row['details'].append(dict(name=str(i.get('subject') or 'Задача РСО'),value=v,zeroed=ind==5 and zeroed(i,p)))
                if v is not None:vals.append(v)
            if not vals:continue
            row['value']=sum(vals)/len(vals)  # Do not round before ranking K5.
            if ind==5:row['no_docs']=all(zeroed(i,p) for i in entries)
            else:
                start=[184,190,196,202][p]
                total=sum(num(cf(i,start)) or 0 for i in entries)
                without=sum(num(cf(i,k)) or 0 for i in entries for k in range(start+1,start+5))
                if total:row['extra']={'total':total,'without':without}
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
    return sorted(rows,key=lambda r:(r['group']=='zato',r['place'],-r['value'] if ind!=6 else r['value'],r['omsu']))

def k7cell(value,item,previous):
    if value is None or not str(value).strip():return {'ok':False,'reason':'Не заполнено'}
    s=re.sub(r'\s+',' ',str(value).lower()).strip(); n=item['n']
    if re.fullmatch(r'\d+(?:[.,]\d+)?',s):
        number=float(s.replace(',','.'))
        if item['pct']:return dict(ok=number>=100,reason='' if number>=100 else 'Выполнено частично')
        if number>0:return {'ok':True}
        return dict(ok=n>=15 or previous,reason='Нулевое количество; оценка зависит от предыдущего этапа')
    if re.search(r'\d+\s*%',s):return dict(ok=False,reason='Процент указан текстом; требуется проверка исходной записи')
    if n==15 and re.search(r'\d+\s*(точ|при|пу\b)',s):return {'ok':True}
    if n==16 and re.search(r'направлен|заявк|исх|№|и-\d\d-',s) and not re.search(r'не\s*направл|будет',s):return {'ok':True}
    if re.search(r'не\s*выполн|в работе|буд(ет|ут)|не\s*направл|проводится|вед[её]тся|выполняется|в стадии|на стадии|не\s*получ|на этапе|ожида|осталось|осталась|в ноябре|в сентябре|некоррект',s):return dict(ok=False,reason='Не завершено')
    if re.search(r'выполн|получен|направ|не\s*треб|не\s*целесообраз|не\s*выявл|в наличии|данные полные|данные есть|сформирован|проведен|выявлен|заменен|устан|запрос|исх|вх\.|заявк|письм|№|и-\d\d-|ответ|экономи|истек мпи|^да$|точ(ек|ки|ка) уч',s):return {'ok':True}
    return dict(ok=False,reason='Нет однозначного подтверждения выполнения',unclear=True)

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
            items.append(dict(name=item['name'],raw=('' if raw is None else str(raw))[:1000],**evaluation))
        candidates[slot].append(items)
    by=defaultdict(list)
    for index,slot in enumerate(K7_RSO):
        options=candidates[index]
        items=max(options,key=lambda xs:sum(bool(x['ok']) for x in xs)) if options else [dict(name=x['name'],raw='',ok=False,reason='РСО нет в таблице') for x in plan]
        done=sum(bool(x['ok']) for x in items)
        by[slot['omsu']].append(dict(name=slot['rso'],value=rnd(done/17*100),done=done,plan=17,found=bool(options),items=items,duplicates=len(options)))
    rows=[dict(omsu=o,group=group(o),value=rnd(sum(x['value'] for x in rs)/len(rs)),details=rs,no_docs=False,extra={'done':sum(x['done'] for x in rs),'plan':17*len(rs)}) for o,rs in by.items()]
    for g in ('main','zato'):rank([r for r in rows if r['group']==g])
    return sorted(rows,key=lambda r:(r['group']=='zato',r['place'],r['omsu'])),unmatched

def build(sources):
    indicators={}
    for ind in (4,5,6,7):
        src=sources.get(str(ind),{})
        periods=[rating(ind,p,src['data']) if ind<7 else rating7(src['data'])[0] if p==2 else [] for p in range(4)] if 'data' in src else [[],[],[],[]]
        counts=[len(rs) for rs in periods];maximum=max(counts)
        default=next((p for p in range(3,-1,-1) if counts[p] and counts[p]>=maximum/2),None)
        indicators[str(ind)]={**CONFIG[ind],'periods':periods,'default_period':default,'updated_at':src.get('updated_at'),'error':src.get('error'),'provisional':ind in (6,7)}
    return dict(methodology_year=2026,period_labels=PERIODS,municipalities=ALL_OMSU,indicators=indicators)
