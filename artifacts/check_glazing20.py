import json
from pathlib import Path
BASE=Path(__file__).resolve().parent/'campaign20_six_corpuses_20260905'
RUN=BASE/'runs/20260905-112254-171550'
manifest=json.loads((BASE/'manifest.json').read_text(encoding='utf-8'))
findings=[]
for case in manifest['cases']:
    p=json.loads((BASE/case['input']).read_text(encoding='utf-8'))
    tasks=json.loads((RUN/case['id']/(case['id']+'.ir.json')).read_text(encoding='utf-8'))['tasks']
    for n,corpus in enumerate(p['корпуса'],1):
        pref=f'К{n}. '
        rows=[t for t in tasks if t['name'].startswith(pref)]
        glass=corpus['остекление']
        stained=[t for t in rows if 'По договору Монтаж светопрозрачных конструкций Витраж' in t['name']]
        masonry=[t for t in rows if t['name']==pref+'По договору Кладка наружных стен']
        if glass['витражи'] and not stained:
            findings.append({'case':case['id'],'corpus':n,'code':'MISSING-STAINED',
                             'evidence':[t for t in rows if 'По договору Монтаж светопрозрачных' in t['name'] or 'Кровля/Парапет Монолит' in t['name'] or 'Закрыт тепловой контур по корпусу' in t['name']]})
        if glass['витражи_на_всю_высоту'] and masonry:
            findings.append({'case':case['id'],'corpus':n,'code':'UNEXPECTED-MASONRY','evidence':masonry})
with (BASE/'glazing-findings.json').open('x',encoding='utf-8') as f: json.dump(findings,f,ensure_ascii=False,indent=2)
from collections import Counter
print(Counter(f['code'] for f in findings))
print('affected projects',len({f['case'] for f in findings}))
print(json.dumps([f for f in findings if (f['case'],f['corpus']) in [('valid-004',3),('valid-003',1)]],ensure_ascii=False,indent=2))
