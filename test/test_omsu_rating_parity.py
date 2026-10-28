"""Synthetic differential checks against the supplied developer's JavaScript.

The reference fixture contains only calculation functions, no executable archive,
keys, source requests, page bootstrap or real operational records.
"""
import json
import os
from pathlib import Path
import random
import shutil
import subprocess

import pytest

from services.omsu_rating import calculate as calc


def issue(omsu,fields,subject='РСО'):
    return {'id':1,'assigned_to':{'name':omsu},'subject':subject,
            'custom_fields':[{'id':k,'value':v} for k,v in {17:1,**fields}.items()]}


def reference(payload,method='evaluate'):
    node=os.getenv('NODE_BINARY') or shutil.which('node')
    if not node:pytest.skip('Node.js is needed only for original JavaScript differential tests')
    fixture=Path(__file__).parent/'fixtures/omsu_rating_original_rules.js'
    script="const fs=require('node:fs');const ref=require(process.argv[1]);const input=JSON.parse(fs.readFileSync(0,'utf8'));process.stdout.write(JSON.stringify(ref[process.argv[2]](input)));"
    result=subprocess.run([node,'-e',script,str(fixture),method],input=json.dumps(payload,ensure_ascii=False),capture_output=True,text=True,timeout=15,check=True)
    return json.loads(result.stdout)


@pytest.mark.parametrize('ind',[4,5,6])
def test_original_js_calculation_parity_all_periods(ind):
    rng=random.Random(947+ind)
    data=[]
    names=['ЩЕЛКОВО','ДУБНА','ИСТРА','ВЛАСИХА','ВОСХОД','ХИМКИ','КЛИН']
    for index,omsu in enumerate(names):
        for duplicate in range(1+index%3):
            fields={17:0 if index==0 else 1}
            for p in range(4):
                fields[165+p]=rng.choice([None,0,1000,200000,'6,67',12000.005])
                fields[209+p]=rng.choice([None,0,4000,15000])
                fields[158+p]=rng.choice([None,0,4,6.6667,6.67,10])
                fields[[189,195,201,207][p]]=rng.choice([None,0,2,50,99])
                if p:fields[[0,259,260,261][p]]=rng.choice([None,'0',['','337','339','341'][p]])
            fields.update({169:rng.choice([100000,299999,300000,599999,600000]),213:rng.choice([None,0,100000,600000]),208:'Вторая РСО'})
            data.append(issue(omsu,fields,f'РСО {index}-{duplicate}'))
    data.extend([issue('',{158:8}),issue('НЕ УЧАСТВУЕТ',{17:None,158:10})])
    expected=reference({'ind':ind,'data':{str(ind):data}})
    result=calc.build({str(ind):{'data':data}})['indicators'][str(ind)]
    assert result['default_period']==expected['default_period']
    for period,rows in enumerate(result['periods']):
        comparable=[{key:row[key] for key in ('omsu','group','value','place','no_docs','place_label')} for row in rows]
        assert comparable==expected['periods'][period]


def test_original_js_k7_slots_duplicate_selection_and_missing_participants():
    sheets=[]
    pairs=[('МУП «БКС»','БАЛАШИХА'),('АО «ПТО ГХ»','ДУБНА'),('МУП «Истринский водоканал»','ИСТРА'),
           ('АО «Истринская теплосеть»','ИСТРА'),('МУП «Теплосеть»','ДОМОДЕДОВО'),
           ('АО «Водоканал» г.о. Красногорск','КРАСНОГОРСК'),('МУП «Водоканал»','ЛЮБЕРЦЫ')]
    for index,(name,omsu) in enumerate(pairs):
        row=['']*22;row[2]=name;row[3]=omsu
        for item in calc.K7_ITEMS[:17]:row[item['col']]=100 if item['pct'] else ('выполнено' if (index+item['n'])%3 else 'в работе')
        sheets.append(row)
    sheets.append(['','','АО «ПТО ГХ»','ДУБНА',None])  # Less complete duplicate must lose.
    expected=reference({'ind':7,'data':{'7':sheets}})
    result=calc.build({'7':{'data':sheets}})['indicators']['7']
    comparable=[{key:row[key] for key in ('omsu','group','value','place','no_docs','place_label')} for row in result['periods'][2]]
    assert comparable==expected['periods'][2]
    dubna=next(r for r in result['periods'][2] if r['omsu']=='ДУБНА')
    assert dubna['details'][0]['found'] is True and dubna['details'][0]['duplicates']==2
    assert dubna['details'][0]['sheet_name']=='АО «ПТО ГХ»'


@pytest.mark.parametrize('value,item,previous',[
    (None,{'n':1},True),('',{'n':1},True),(100,{'n':5,'pct':True},True),
    (99,{'n':5,'pct':True},True),('100%',{'n':5,'pct':True},True),
    (0,{'n':13},False),(0,{'n':13},True),(0,{'n':15},False),
    ('осталось 12 точек',{'n':15},False),('заявки направлены, ведётся работа',{'n':16},False),
    ('будет направлено письмо',{'n':16},False),('не требуется',{'n':17},False),
    ('через личный кабинет Россети',{'n':17},True),('неизвестная запись',{'n':1},True),
    ('не выполнено',{'n':17},True),('выявлено',{'n':12},True),
    ('2 пу',{'n':15},True),('осталось 2 ПУ',{'n':15},True),('2 пуabc',{'n':15},True),
    ('٢ %',{'n':1},True),('١ пу',{'n':15},True),('и-٢٣-xxx',{'n':16},True),
])
def test_original_js_k7_interpretation_and_explanation(value,item,previous):
    expected=reference(dict(value=value,item=item,previous=previous),'cell')
    assert calc.k7cell(value,item,previous)==expected


def test_color_zones_tied_ranges_and_bars_use_unfiltered_group_not_fixed_51():
    rows=calc.rating(5,0,[issue(name,{158:score}) for name,score in [('Д',10),('Е',10),('А',8),('Б',6),('В',4),('Г',0),('ВЛАСИХА',3)]])
    by={r['omsu']:r for r in rows}
    assert by['Д']['place_label']==by['Е']['place_label']=='1–2'
    assert [by[n]['zone'] for n in ['Д','А','Б','В','Г']]==['z1','z2','z2','z3','z3']
    assert by['Г']['bar_pct']==2 and by['Д']['bar_pct']==100
    assert by['ВЛАСИХА']['zone']=='z3' and by['ВЛАСИХА']['group_size']==1
    assert by['ВЛАСИХА']['bar_pct']==30  # Same full-rating scale as original.
    assert [r['omsu'] for r in rows[:2]]==['Д','Е']  # Original stable source tie order.


def test_no_docs_ranges_and_ascending_k6_progress():
    rows=calc.rating(5,1,[issue('А',{159:4,259:'337'}),issue('Б',{159:5,259:'337'})])
    assert [r['place_label'] for r in rows]==['51','51']
    rows=calc.rating(6,0,[issue('А',{189:10}),issue('Б',{189:50})])
    assert [r['place_label'] for r in rows]==['1–2','1–2']
    assert [r['bar_pct'] for r in rows]==[80,2]


def test_period_metadata_and_previous_period_exact_name_semantics():
    source=[issue('ДУБНА',{158:5,159:10}),issue('ИСТРА',{158:10,159:5}),issue('ХИМКИ',{159:1})]
    result=calc.build({'5':{'data':source}})['indicators']['5']
    rows={r['omsu']:r for r in result['periods'][1]};stats=result['period_stats'][1]
    assert rows['ДУБНА']['previous_place']==2 and rows['ДУБНА']['delta']==1
    assert rows['ИСТРА']['delta']==-1 and rows['ХИМКИ']['delta'] is None
    assert stats['average']==16/3 and stats['has_previous'] and stats['count']==3
    assert len(stats['nonparticipants'])==53
    assert not result['period_stats'][0]['has_previous']


def test_k4_category_boundaries_and_methodology_details_are_retained():
    data=[issue('А',{166:1000,169:299999,210:2000,213:300000}),issue('Б',{166:3000,169:600000})]
    rows=calc.rating(4,1,data)
    details=[d for r in rows for d in r['details']]
    assert {(d['nvv'],d['category'],d['plan']) for d in details}=={(299999,1,2),(300000,2,3),(600000,3,4)}
    assert all('нарастающим итогом' in r['extra']['fact_tooltip'] for r in rows)
    assert all('Категория' in r['value_tooltip'] for r in rows)
    assert all(r['tag']=='not_achieved' for r in rows)


def test_k4_plan_display_and_sort_use_original_rso_order():
    row=calc.rating(4,1,[issue('А',{166:1000,169:600000,210:2000,213:100000})])[0]
    assert row['extra']['plan']==[4,2]


def test_original_js_numeric_field_conversion_and_array_zero_flags():
    values=['0x10','0b11','0o10','+0x10','1_000','\ufeff 3,5 ',['5'],[],['1','2'],True,False,'6,6667','NaN','Infinity','.5','2e1']
    assert [calc.num(v) for v in values]==reference(values,'numbers')
    row=calc.rating(5,1,[issue('А',{159:5,259:['337']})])[0]
    assert row['no_docs'] is True and row['place']==51
    assert calc.zeroed(issue('А',{259:337.0}),1) is True


def test_all_original_methodology_constants_and_rso_patterns_match():
    original=reference({},'constants')
    assert calc.K7_ITEMS==original['items']
    assert calc.K7_RSO==original['rso']
    assert calc.ALL_OMSU==original['municipalities']
    assert calc.ZATO==original['zato']
    assert {str(k):v for k,v in calc.PLANS.items()}==original['plans']


@pytest.mark.parametrize('ind',[4,5,6,7])
def test_original_visual_rules_are_computed_before_filtering(ind):
    data=[]
    for i,omsu in enumerate(['ИСТРА','ДУБНА','ХИМКИ','КЛИН','ЩЕЛКОВО','ФРЯЗИНО','ВЛАСИХА','ВОСХОД']):
        data.append(issue(omsu,{165:1000*(i+1),166:150*(8-i),169:600000,158:i/3,159:(8-i)/3,189:i+1,195:8-i}))
    payload={'7':[]} if ind==7 else {str(ind):data}
    p=2 if ind==7 else 1
    expected=reference({'ind':ind,'data':payload,'period':p},'presentation')
    actual=calc.build({str(ind):{'data':payload[str(ind)]}})['indicators'][str(ind)]['periods'][p]
    assert [{k:r[k] for k in ('omsu','zone','bar_pct','delta')} for r in actual]==expected


@pytest.mark.parametrize('value',[1e30,1e100,1e308])
def test_large_finite_source_values_keep_data_endpoint_renderable(value):
    result=calc.build({'5':{'data':[issue('ДУБНА',{158:value})]}})['indicators']['5']
    row=result['periods'][0][0]
    assert row['value']==value and row['place']==1
    assert calc.fmt(value).endswith(',00') and '—' not in row['value_tooltip']


def test_original_math_round_at_float_immediately_below_half():
    data=[issue('ДУБНА',{165:0.00004999999999999999,169:1})]
    expected=reference({'ind':4,'data':{'4':data}})['periods'][0][0]
    actual=calc.rating(4,0,data)[0]
    assert actual['value']==expected['value']==0
