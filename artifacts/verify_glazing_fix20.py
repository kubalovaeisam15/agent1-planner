"""Повторная независимая сверка сохранённой матрицы после исправления."""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / 'artifacts/campaign20_six_corpuses_20260905'
RUN = BASE / 'runs/20260905-114958-951447'
sys.path.insert(0, str(ROOT/'tools'))
from project_ir_validation import validate_project_against_ir
from schedule_ir import ScheduleProject, validate_schedule_ir

def read(path):
    return json.loads(path.read_text(encoding='utf-8'))

results = []
for case in read(BASE/'manifest.json')['cases']:
    spec = read(BASE/case['input'])
    raw = read(RUN/case['id']/(case['id']+'.ir.json'))
    ir = ScheduleProject.from_dict(raw)
    issues = [i.code for i in validate_schedule_ir(ir) + validate_project_against_ir(spec, ir)]
    for n, corpus in enumerate(spec['корпуса'], 1):
        prefix = f'К{n}. '
        rows = [t for t in ir.tasks if t.name.startswith(prefix)]
        raft = next(t for t in rows if t.name == prefix+'По договору Фундаменты Корпус')
        below = next(t for t in rows if t.name.startswith(prefix+'По договору Монолитные конструкции ниже'))
        first = next(t for t in rows if t.name == prefix+'1 этаж Монолит')
        expected = below if corpus['этажей_подземных'] else raft
        if raft.duration_days != (45 if corpus['этажей_надземных'] <= 45 else 60):
            issues.append('FOUNDATION-DURATION')
        if below.duration_days != 45*corpus['этажей_подземных']:
            issues.append('BELOW-DURATION')
        if [(p.predecessor_id,p.type,p.lag_days) for p in first.predecessors] != [(expected.task_id,'FS',0)]:
            issues.append('FIRST-FLOOR-PREDECESSOR')
    results.append({'id':case['id'],'corpuses':len(spec['корпуса']),'issues':issues})
example = read(RUN/'valid-004/valid-004.ir.json')
contour = next(t for t in example['tasks'] if t['name']=='К3. Закрыт тепловой контур по корпусу')
result = {'cases':results, 'example_valid004_K3':contour}
with (RUN/'independent-fix-validation.json').open('x',encoding='utf-8') as f:
    json.dump(result,f,ensure_ascii=False,indent=2)
print(json.dumps({'cases':len(results),'corpuses':sum(r['corpuses'] for r in results),
                  'errors':sum(len(r['issues']) for r in results),'contour':contour},ensure_ascii=False,indent=2))
assert not any(r['issues'] for r in results)
