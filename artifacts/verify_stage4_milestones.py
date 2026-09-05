"""Независимый отчёт о связях БТИ и результатах предметной проверки 20 проектов."""
import json
import sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from schedule_ir import ScheduleProject,validate_schedule_ir
from project_ir_validation import validate_project_against_ir
BASE=ROOT/'artifacts/campaign20_six_corpuses_20260905'
RUN=BASE/'runs/20260905-230041-846320'
OLD=BASE/'runs/20260905-114958-951447'
def read(path):
    return json.loads(path.read_text(encoding='utf-8'))
results=[]
for case in read(BASE/'manifest.json')['cases']:
    ident=case['id']
    spec=read(BASE/case['input'])
    ir=ScheduleProject.from_dict(read(RUN/ident/(ident+'.ir.json')))
    old=read(OLD/ident/(ident+'.ir.json'))
    before=next(t for t in old['tasks'] if t['name']=='Готовность к обмерам БТИ')
    after=next(t for t in ir.tasks if t.name=='Готовность к обмерам БТИ')
    by_id={t.task_id:t for t in ir.tasks}
    issues=validate_schedule_ir(ir)+validate_project_against_ir(spec,ir)
    results.append({'id':ident,'corpuses':len(spec['корпуса']),
                    'errors':[i.code for i in issues],'before_bti':before['finish'],
                    'after_bti':after.finish.isoformat(),'before_links':len(before['predecessors']),
                    'after_links':len(after.predecessors),
                    'predecessors':[{'name':by_id[p.predecessor_id].name,
                                     'type':by_id[p.predecessor_id].task_type,'lag':p.lag_days}
                                    for p in after.predecessors]})
with (RUN/'stage4-validation.json').open('x',encoding='utf-8') as f:
    json.dump(results,f,ensure_ascii=False,indent=2)
print({'projects':len(results),'corpuses':sum(r['corpuses'] for r in results),
       'errors':sum(len(r['errors']) for r in results),
       'bti_dates_changed':sum(r['before_bti']!=r['after_bti'] for r in results)})
assert all(not r['errors'] and r['after_links']==r['corpuses'] and
           all(p['type']=='summary' and p['lag']==-160 for p in r['predecessors']) for r in results)
