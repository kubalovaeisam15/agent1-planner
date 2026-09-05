"""Повторное чтение 20 Excel/IR, без изменения графиков."""
import contextlib
import io
import json
import hashlib
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from schedule_ir import ScheduleProject, validate_schedule_ir
import validate_grp
BASE = ROOT/'artifacts/campaign20_six_corpuses_20260905'
RUN = BASE/'runs/20260905-114958-951447'
results = []
for case in json.loads((BASE/'manifest.json').read_text(encoding='utf-8'))['cases']:
    ident = case['id']
    ir_path = RUN/ident/(ident+'.ir.json')
    xlsx = RUN/ident/(ident+'.xlsx')
    raw = ir_path.read_bytes()
    issues = validate_schedule_ir(ScheduleProject.from_json(raw.decode('utf-8')))
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        status = validate_grp.validate(validate_grp.load(xlsx),validate_grp.load_reserve_base(xlsx),
                                       validate_grp.load_milestone_maps(xlsx))
    results.append({'id':ident,'ir_sha256':hashlib.sha256(raw).hexdigest(),
                    'xlsx_sha256':hashlib.sha256(xlsx.read_bytes()).hexdigest(),
                    'ir_errors':[i.code for i in issues],'excel_exit':status,
                    'excel_log':output.getvalue()})
with (BASE/'STAGE3_VALIDATION_20260905.json').open('x',encoding='utf-8') as f:
    json.dump(results,f,ensure_ascii=False,indent=2)
print({'cases':len(results),'ir_failures':sum(bool(r['ir_errors']) for r in results),
       'excel_failures':sum(r['excel_exit']!=0 for r in results)})
assert all(not r['ir_errors'] and r['excel_exit']==0 for r in results)
