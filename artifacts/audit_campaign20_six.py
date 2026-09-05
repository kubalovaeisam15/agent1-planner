import copy
import json
import math
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'tools'))
from build_grp import validate_project_spec
from schedule_ir import ScheduleProject, validate_schedule_ir

BASE=ROOT/'artifacts/campaign20_six_corpuses_20260905'
RUN=BASE/'runs/20260905-112254-171550'
manifest=json.loads((BASE/'manifest.json').read_text(encoding='utf-8'))
checks=[]
for case in manifest['cases']:
    ident=case['id']
    p=json.loads((BASE/case['input']).read_text(encoding='utf-8'))
    ir=json.loads((RUN/ident/(ident+'.ir.json')).read_text(encoding='utf-8'))
    tasks=ir['tasks']
    failures=[]
    for n,c in enumerate(p['корпуса'],1):
        pref=f'К{n}. '
        raft=[t for t in tasks if t['name']==pref+'По договору Фундаменты Корпус']
        below=[t for t in tasks if t['name'].startswith(pref+'По договору Монолитные конструкции ниже')]
        first=[t for t in tasks if t['name']==pref+'1 этаж Монолит']
        expected_raft=45 if c['этажей_надземных']<=45 else 60
        if len(raft)!=1 or raft[0]['duration_days']!=expected_raft: failures.append(pref+'raft')
        if len(below)!=1 or below[0]['duration_days']!=45*c['этажей_подземных']: failures.append(pref+'below')
        if len(first)!=1: failures.append(pref+'first')
        elif raft and below:
            pred=below[0] if c['этажей_подземных'] else raft[0]
            if first[0]['predecessors']!=[{'predecessor_id':pred['task_id'],'type':'FS','lag_days':0}]:
                failures.append(pref+'first-link')
    for pile in p['нулевой_цикл']['сваи']:
        label='БНС' if pile['тип']=='БНС' else 'Забивные'
        rows=[t for t in tasks if 'По договору Свайное основание '+label in t['name']]
        expected=math.ceil(pile['количество']/((1 if label=='БНС' else 10)*pile['установок']))
        if expected and (len(rows)!=1 or rows[0]['duration_days']!=expected): failures.append('pile '+label)
    checks.append({'id':ident,'corpus_count':len(p['корпуса']),'issues':failures})

base=json.loads((BASE/'inputs/valid-001.json').read_text(encoding='utf-8'))
probes=[]
for label, mutate in [
    ('facade-list',lambda p:p['фасад'].update(тип=[])),
    ('facade-object',lambda p:p['фасад'].update(тип={})),
    ('negative-depth',lambda p:p['корпуса'][0].update(этажей_подземных=-3)),
    ('zero-rigs',lambda p:p['нулевой_цикл']['сваи'][0].update(установок=0)),
]:
    p=copy.deepcopy(base);mutate(p)
    with (BASE/(label+'.repro.json')).open('x',encoding='utf-8') as f: json.dump(p,f,ensure_ascii=False,indent=2)
    try: result={'errors':validate_project_spec(p)}
    except Exception as e: result={'crash':type(e).__name__+': '+str(e)}
    probes.append({'probe':label,**result})

original=json.loads((RUN/'valid-018/valid-018.ir.json').read_text(encoding='utf-8'))
for label in ['wrong-parent','wrong-summary-dates']:
    ir=copy.deepcopy(original)
    if label=='wrong-parent':
        leaf=next(t for t in ir['tasks'] if t['name']=='К6. 1 этаж Монолит')
        leaf['parent_id']=next(t['task_id'] for t in ir['tasks'] if t['name']=='К1. 1 этаж Монолит')
    else:
        head=next(t for t in ir['tasks'] if t['task_type']=='summary' and not t['predecessors'])
        head['start']='1900-01-01'
    errors=validate_schedule_ir(ScheduleProject.from_dict(ir))
    with (BASE/(label+'.repro.ir.json')).open('x',encoding='utf-8') as f:json.dump(ir,f,ensure_ascii=False,indent=2)
    probes.append({'probe':label,'errors':[e.code for e in errors]})

report={'independent_checks':checks,'mutations':probes}
with (BASE/'independent-audit.json').open('x',encoding='utf-8') as f:json.dump(report,f,ensure_ascii=False,indent=2)
print(json.dumps(report,ensure_ascii=False,indent=2))
