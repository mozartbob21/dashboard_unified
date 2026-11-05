from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook, load_workbook

from services.aichat.agents import neighbors_rating as calc


def xlsx(sheets):
    wb=Workbook()
    wb.remove(wb.active)
    for title,rows in sheets.items():
        sh=wb.create_sheet(title)
        for row in rows:
            sh.append(row)
    b=BytesIO()
    wb.save(b)
    wb.close()
    return b.getvalue()


def sample(only_one=False, empty=False):
    geo = [calc.GEO_COLUMNS,
           ['01','Чехов г, Дружбы ул, 8/2','37.49 55.14','Московская область, Чехов',''],
           ['02','Чехов г, Дружбы ул, 4','37.4907 55.14','Чехов',''],
           ['03','Чехов г, Мира ул, 3','37.4918 55.14','Чехов',''],
           ['04','Чехов г, Троицкое с, 6','38.5 55.75','Мытищи','']]
    eds=[calc.EDS_COLUMNS]
    ecur=[calc.ECUR_COLUMNS]
    if not empty:
        eds.append(['-123','Инцидент','Прикладной текст не нужен','01','2026-10-01'])
    if not only_one and not empty:
        eds.extend([['456','ЕДС','Адрес','02','2026-10-02'],['457','ЕДС','Адрес','04','2026-10-03']])
        ecur.extend([
            ['0',None,'ГВС','ЕЦУР',None,'111','2026-10-03'],
            [None,None,'Отопление','ЕЦУР','Чехов','123','2026-10-04'],
            [None,'Заполнен','Другое','ЕЦУР','Россия, Московская область, Чехов, улица Дружбы, 8к2, кв. 123, этаж 5','333','2026-10-05'],
            [None,None,'Тепловые сети','ЕЦУР','Чехов, улица Мира, 3','335','2026-10-06'],
            [None,None,'Сети','+7 999 000 11 22','Россия, Московская область, Чехов, улица Южная, 3, кв. 123, owner@example.test','336','2026-10-07'],
            [None,None,'Сети','имя@example.org',None,'337','2026-10-07']
        ])
    return xlsx({'Лист1':geo}),xlsx({calc.SHEET_EDS:eds,calc.SHEET_ECUR:ecur})


def test_rating_cache_distances_themes_bad_points_unplaced_and_privacy():
    data, summary=calc.build_rating(*sample())
    stats=summary['stats']
    assert [stats[k] for k in ('rows_eds','rows_ecur','dup_flagged','dup_hidden','unique_complaints','unplaced','in_rating','houses','houses_heat_only','bad_point_houses')]==[3,6,1,1,7,2,5,4,3,1]
    assert summary['top10'][0]['count']==2
    assert summary['top10'][0]['address']=='Чехов г, Дружбы ул, 8/2'
    assert summary['top10'][0]['distance'] in (44,45) and summary['top10'][0]['neighbors']==1
    with BytesIO(data) as src:
        wb=load_workbook(src,data_only=True)
        assert wb.sheetnames==['Рейтинг','Только отопление','Не привязаны','Как считали']
        assert wb['Рейтинг'].freeze_panes=='C2' and wb['Только отопление'].freeze_panes=='C2'
        assert wb['Рейтинг']['A3'].value==wb['Рейтинг']['A4'].value==2
        assert wb['Рейтинг']['C2'].value==2 and wb['Рейтинг']['D2'].value==wb['Рейтинг']['E2'].value==1
        main_near = {row[1]:row[9] for row in list(wb['Рейтинг'].values)[1:]}
        heat_near = {row[1]:row[6] for row in list(wb['Только отопление'].values)[1:]}
        assert main_near['Чехов г, Дружбы ул, 4']==2 and heat_near['Чехов г, Дружбы ул, 4']==1
        assert wb['Как считали']['B11'].value==1
        assert wb['Как считали']['B8'].value==wb['Как считали']['B9'].value==5
        missing=list(wb['Не привязаны'].values)
        assert 'кв. 123' not in str(missing) and '@' not in str(missing) and '999 000' not in str(missing)
        assert wb['Рейтинг']['M2'].value and wb['Рейтинг']['G5'].value is None
        wb.close()
    wb=load_workbook(BytesIO(data))
    assert wb['Рейтинг']['A2'].value=='=RANK(C2,$C$2:$C$5)'
    assert wb['Рейтинг']['C2'].value=='=SUM(D2:F2)'
    wb.close()


@pytest.mark.parametrize('empty',[True,False])
def test_no_complaints_or_single_house_no_self_neighbor(empty):
    data,result=calc.build_rating(*sample(only_one=True,empty=empty))
    assert result['stats']['median_nearest_m'] is None
    if empty:
        assert result['top10']==[]
    else:
        r=result['top10'][0]
        assert r['nearest'] is None and r['distance'] is None and r['neighbors']==0
    wb=load_workbook(BytesIO(data),data_only=True)
    assert wb['Как считали']['B10'].value==int(not empty)
    wb.close()


def test_geo_match_requires_place_and_prefers_eds_over_yandex():
    geo=[calc.GEO_COLUMNS,
         ['01','Чехов г, Мещерское с, 14','37.51 55.14','Чехов',''],
         ['02','Чехов г, Мещерское с, 14А','bad coordinates','Мытищи','kind":"locality","name":"село Мещерское";kind":"house","name":"14"'],
         ['03','Чехов-8 г, Южная ул, 4','37.49 55.14','Чехов-8','']]
    ecur=[calc.ECUR_COLUMNS,
          [None,None,'Отопление','ЕДДС','Чехов, село Мещерское, 14','','2026-10-10'],
          [None,None,'Сети','ЕДДС','Чехов, рабочий посёлок Столбовая, улица Южная, 4','','2026-10-10']]
    _,r=calc.build_rating(xlsx({'Справочник':geo}),xlsx({calc.SHEET_EDS:[calc.EDS_COLUMNS],calc.SHEET_ECUR:ecur}))
    assert r['stats']['unplaced']==1 and r['stats']['in_rating']==1
    assert r['top10'][0]['address']=='Чехов г, Мещерское с, 14'


def test_literal_excel_payload_is_not_a_formula_and_repeated_id_fails_closed():
    geo, _=sample()
    ecur=[calc.ECUR_COLUMNS,[None,None,'Сети','+SUM(2,2)','Крюково','123','2026-10-05']]
    data, _=calc.build_rating(geo,xlsx({calc.SHEET_EDS:[calc.EDS_COLUMNS],calc.SHEET_ECUR:ecur}))
    wb=load_workbook(BytesIO(data))
    assert wb['Не привязаны']['D2'].value=='+SUM(2,2)' and wb['Не привязаны']['D2'].data_type=='s'
    wb.close()
    ecur.append(ecur[-1])
    with pytest.raises(calc.RatingError,match='повторяется внутренний'):
        calc.build_rating(geo,xlsx({calc.SHEET_EDS:[calc.EDS_COLUMNS],calc.SHEET_ECUR:ecur}))
    ecur[1][3]='=SUM(2,2)'
    with pytest.raises(calc.RatingError,match='без формул'):
        calc.build_rating(geo,xlsx({calc.SHEET_EDS:[calc.EDS_COLUMNS],calc.SHEET_ECUR:ecur[:2]}))


def test_missing_code_and_wrong_district_not_published():
    geo,_=sample()
    c=xlsx({calc.SHEET_EDS:[calc.EDS_COLUMNS,['123','ЕДС','','234','2026-10-10']],calc.SHEET_ECUR:[calc.ECUR_COLUMNS]})
    with pytest.raises(calc.RatingError,match='Код дома из ЕДС'):
        calc.build_rating(geo,c)
    with pytest.raises(calc.RatingError,match='границы'):
        calc.build_rating(geo,c,city='Дубна')
    with pytest.raises(calc.RatingError,match='нет домов'):
        calc.build_rating(geo,c,city='Дубна',bbox=[37,38,55,56])


def test_xml_entities_and_non_xlsx_rejected():
    geo,comp=sample()
    b=BytesIO()
    with ZipFile(BytesIO(geo)) as clean, ZipFile(b,'w',ZIP_DEFLATED) as bad:
        for name in clean.namelist():
            raw=clean.read(name)
            if name=='xl/workbook.xml':
                raw=b'<!DOCTYPE test [<!ENTITY hack "boom">]>'+raw
            bad.writestr(name,raw)
    with pytest.raises(calc.RatingError,match='DTD'):
        calc.build_rating(b.getvalue(),comp)
    with pytest.raises(calc.RatingError):
        calc.build_rating(b'not an XLSX',comp)
