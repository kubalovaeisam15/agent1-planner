"""Read-only проверка сохранённых IR и двух исходных воспроизведений V03."""
import hashlib
import json
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'tools'))
from schedule_ir import ScheduleProject, validate_schedule_ir
BASE = ROOT/'artifacts/campaign20_six_corpuses_20260905'
RUN = BASE/'runs/20260905-114958-951447'
results = []
for case in json.loads((BASE/'manifest.json').read_text(encoding='utf-8'))['cases']:
    path = RUN/case['id']/(case['id']+'.ir.json')
    raw = path.read_bytes()
    issues = validate_schedule_ir(ScheduleProject.from_json(raw.decode('utf-8')))
    results.append({'case':case['id'],'sha256':hashlib.sha256(raw).hexdigest(),
                    'errors':[{'code':i.code,'task_id':i.task_id} for i in issues]})
mutations = []
for name, expected in [('wrong-parent','IR-PARENT'), ('wrong-summary-dates','IR-SUMMARY-DATES')]:
    path = BASE/(name+'.repro.ir.json')
    issues = validate_schedule_ir(ScheduleProject.from_json(path.read_text(encoding='utf-8')))
    mutations.append({'case':name,'expected':expected,'codes':sorted({i.code for i in issues})})
result = {'cases':results,'mutations':mutations}
with (BASE/'STAGE2_VALIDATION_20260905.json').open('x',encoding='utf-8') as stream:
    json.dump(result,stream,ensure_ascii=False,indent=2)
print(json.dumps(result,ensure_ascii=False,indent=2))
assert all(not r['errors'] for r in results)
assert all(r['expected'] in r['codes'] for r in mutations)
