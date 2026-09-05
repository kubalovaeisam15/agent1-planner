import copy
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT / 'artifacts/campaign20_six_corpuses_20260905'
sys.path.insert(0, str(ROOT / 'tools'))
from schedule_ir import ScheduleProject, validate_schedule_ir

out = BASE / 'minimal-repros'
out.mkdir(exist_ok=False)
base = json.loads((BASE / 'inputs/valid-001.json').read_text(encoding='utf-8'))
results = []
for label, glass in [('mixed', {'пвх': True, 'витражи': True, 'витражи_на_всю_высоту': False}),
                     ('full', {'пвх': False, 'витражи': True, 'витражи_на_всю_высоту': True})]:
    p = copy.deepcopy(base)
    p['корпуса'][0]['остекление'] = glass
    inp, xlsx, ir = [out / (label + ext) for ext in ('.json', '.xlsx', '.ir.json')]
    with inp.open('x', encoding='utf-8') as f:
        json.dump(p, f, ensure_ascii=False, indent=2)
    proc = subprocess.run([sys.executable, str(ROOT/'tools/build_grp.py'), str(inp), str(xlsx), '--ir', str(ir)], cwd=ROOT, capture_output=True, text=True, encoding='utf-8', errors='replace')
    with (out/(label+'.log')).open('x', encoding='utf-8') as f:
        f.write(proc.stdout + proc.stderr)
    result = {'case': label, 'build_exit': proc.returncode}
    if ir.exists():
        doc = json.loads(ir.read_text(encoding='utf-8'))
        rows = doc['tasks']
        result['ir_errors'] = [e.code for e in validate_schedule_ir(ScheduleProject.from_dict(doc))]
        result['stained_count'] = sum('К1. По договору Монтаж светопрозрачных конструкций Витраж' in t['name'] for t in rows)
        result['masonry_count'] = sum(t['name'] == 'К1. По договору Кладка наружных стен' for t in rows)
    results.append(result)
doc = json.loads((BASE/'runs/20260905-112254-171550/valid-004/valid-004.ir.json').read_text(encoding='utf-8'))
next(t for t in doc['tasks'] if t['task_id'] == '1554')['predecessors'] = []
with (out/'missing-contour-predecessors.ir.json').open('x', encoding='utf-8') as f:
    json.dump(doc, f, ensure_ascii=False, indent=2)
results.append({'case':'missing-contour-predecessors', 'ir_errors':[e.code for e in validate_schedule_ir(ScheduleProject.from_dict(doc))]})
with (out/'results.json').open('x', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
print(json.dumps(results, ensure_ascii=False, indent=2))
