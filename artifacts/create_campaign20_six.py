import importlib.util
import random
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
source = ROOT / '.agents/skills/grp-agent-tester/scripts/campaign.py'
spec = importlib.util.spec_from_file_location('campaign', source)
campaign = importlib.util.module_from_spec(spec)
spec.loader.exec_module(campaign)
base_make = campaign.make_spec


def make_case(index, rng):
    project, dims = base_make(index, rng)
    count = index % 6 + 1
    depth = (index // 6 + index % 6) % 4
    facade = campaign.FACADES[(index // 6 + index) % 3]
    profile = (index + index // 6) % 4
    glass = list(campaign.GLAZING) + [
        {'пвх': False, 'витражи': True, 'витражи_на_всю_высоту': False}]
    labels = ['pvc', 'mixed', 'full_stained', 'partial_stained']
    floors = [9, 44, 45, 46, 60, 61, 75]
    corpuses = []
    for n in range(count):
        corpuses.append({
            'код': f'К{n+1}', 'этажей_надземных': floors[(index+n) % len(floors)],
            'этажей_подземных': (depth+n) % 4 if index >= 12 else depth,
            'секций': 1+n%3, 'сложный_конструктив': bool((index+n)%2),
            'остекление': dict(glass[(index+n)%4]),
        })
    project['название'] = f'Аудит 20: проект {index+1:02d}, {count} корпусов'
    project['корпуса'] = corpuses
    project['фасад']['тип'] = facade
    project['нулевой_цикл']['сваи'] = [
        {'тип': k, 'количество': q, 'установок': r} for k,q,r in campaign.PILE_SETS[profile]]
    project['отделка']['доля_квартир_с_чистовой'] = (0,0.5,1)[(index+index//6)%3]
    project['паркинг']['этажей_подземных'] = max(1, depth)
    dims.update(corpus_count=count, floors=[c['этажей_надземных'] for c in corpuses],
                underground_floors=depth, underground_by_corpus=[c['этажей_подземных'] for c in corpuses],
                facade=facade, pile_profile=profile,
                glazing=[labels[(index+n)%4] for n in range(count)],
                finish_share=project['отделка']['доля_квартир_с_чистовой'],
                support='exploratory' if count > 2 else 'confirmed')
    return project, dims


campaign.make_spec = make_case
if __name__ == '__main__':
    campaign.create_campaign(ROOT / 'artifacts/campaign20_six_corpuses_20260905',
                             20, 20260905, False)
