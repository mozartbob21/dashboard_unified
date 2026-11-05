"""Exercise a Qwen-selected read through the same dispatcher as an ordinary chat turn."""
import json
from unittest.mock import Mock
import pytest
from fastapi import HTTPException
from services import check_history
from services.aichat import read_agent, report_context

QUESTION='На какой стадии подключение РСО?'
USER={'id':12,'modules':['water-dashboard']}
def test_planned_report_reads_saved_dashboard_then_synthesizes_with_bound_scope(monkeypatch):
    from services import water_ai_context as water
    snapshot={'schema_version':2,'sources':{'gkh_kontur':{'metric_schema':1,'ok':True,'updated_at':'2026-11-03',
              'metrics':[{'id':'roadmap_pct','label':'Доля выполнения','value':47,'unit':'%'}]}},'table':[]}
    monkeypatch.setattr(water,'load_snapshot',lambda:snapshot)
    model=Mock(return_value=json.dumps({'tool':'report','modules':['water-dashboard']}))
    monkeypatch.setattr('services.summarizer.engine._qwen_chat',model)
    evidence=[]
    def narrate(context,modules):
        assert modules==['water-dashboard']
        evidence.append(json.loads(context))
        return 'Показатель сохранён в сводном дашборде.'
    reply=read_agent.resolve(QUESTION,USER,{}, {},narrate=narrate)
    model.assert_called_once()
    assert evidence and {s['module'] for s in evidence[0]['sources']}=={'water-dashboard'}
    assert any(m.get('value')==47 for s in evidence[0]['sources'] for m in s.get('metrics',[]))
    assert '47' in reply['answer'] and '2026-11-03' in str(reply['sources'])
    assert reply['scope']['modules']==['water-dashboard']

def test_authorization_rechecked_between_planning_and_source_read(monkeypatch):
    monkeypatch.setattr('services.summarizer.engine._qwen_chat',lambda *a,**k:'{"tool":"report","modules":["water-dashboard"]}')
    read=Mock(side_effect=AssertionError('a revoked source must not open'))
    monkeypatch.setattr(report_context,'prepare',read)
    with pytest.raises(HTTPException) as exc:
        read_agent.resolve(QUESTION,USER,{}, {},refresh=lambda:{**USER,'modules':[]})
    assert exc.value.status_code==403
    read.assert_not_called()

def test_read_agent_does_not_start_checks_or_replay_history_for_causes(monkeypatch):
    monkeypatch.setattr(check_history,'compare',Mock(side_effect=AssertionError('wrong tool')))
    from services.aichat import agent_intent
    for phrase in ['Чехов за неделю по обращениям', 'что там по Чехову-7 по причинам жалоб']:
        choice=agent_intent.decide(phrase,{'modules':['mingkh']})
        assert choice and (choice['municipality']=='Чехов' if 'неделю' in phrase else choice['action']=='clarify')
