# 岗位评分（示例画像：应届、0年经验、Unity/前端/实施方向——请按自己的情况改规则）：权重 能进30 / 对口25 / 稳定口碑20 / 薪资15 / 成长10
import os, sys, re, json, glob
sys.stdout.reconfigure(encoding='utf-8')
import warnings; warnings.filterwarnings('ignore')
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

WORK = os.environ.get('JOBHUNT_HOME', './workspace')
BASE = WORK + '/data/'
df = pd.read_excel(BASE + '杭州岗位汇总_20261023.xlsx', sheet_name='岗位明细')

# 猎聘卡片里的"学生可投"标记（只在 raw 里有）
student_only = set()
for f in glob.glob(BASE + 'raw/liepin_hz_*.json'):
    try:
        for it in json.load(open(f, encoding='utf-8')).get('list', []):
            if '学生可投' in it.get('lines', []):
                student_only.add(it['href'].split('?')[0])
    except Exception:
        pass

S = lambda v: '' if pd.isna(v) else str(v)

def track(r):
    return '实习' if r['是否实习'] == '是' else '正式'

def exclude(r):
    return bool(re.search(r'售前|销售|运营|客服|讲师|招生|地推|培训|审核|达人|主播', S(r['岗位'])))

def s_enter(r):
    exp, edu = S(r['经验']), S(r['学历'])
    if re.search(r'不限|应届|在校|实习|1年以[下内]|无经验|经验不限', exp):
        sc = 35
    else:
        m = re.match(r'(\d+)', exp)
        lo = int(m.group(1)) if m else 3
        sc = 8 if lo <= 1 else 3 if lo == 2 else 0
    if re.search(r'硕士|博士', edu): sc = 0
    up = size_up(r['规模']); t = S(r['岗位'])
    if re.search(r'高级|资深|专家|主管|负责人|架构|经理|leader|Lead|组长', t, re.I): sc = min(sc, 3)
    if track(r) == '正式':
        if up >= 10000: sc = min(sc, 14)      # 大厂：投递量大，无经验者很难过筛选
        elif up >= 1000: sc = min(sc, 22)
    elif up >= 10000: sc = min(sc, 20)         # 大厂实习同样竞争激烈
    if '统招' in edu: sc -= 3
    if '不可投' in S(cohort(r)): sc = min(sc, 3)
    elif '已结束' in S(cohort(r)): sc = min(sc, 18)
    if track(r) == '实习' and grad_flag(r) == '存疑(仅限在校)': sc = min(sc, 6)
    return max(sc, 0)

def grad_flag(r):
    if track(r) != '实习': return ''
    t = S(r['岗位'])
    if S(r['链接']).split('?')[0] in student_only or re.search(r'在校|在读|2[78]届|暑期|寒假', t + S(r['经验'])):
        return '存疑(仅限在校)'
    return '可能可投'

def s_fit(r):
    t = S(r['岗位']); tags = S(r['技能标签']); txt = t + ' ' + tags
    if re.search(r'美术|特效|动效|原画|建模|动画师|TA|设计师|策划', t): return 3
    if re.search(r'Unity|游戏客户端|游戏开发|U3D|虚幻|Unreal', t, re.I): sc = 25
    elif re.search(r'前端|Vue|Web|H5|小程序', t, re.I): sc = 22
    elif re.search(r'实施|交付|部署', t): sc = 20
    elif re.search(r'C#|\.NET|上位机', t, re.I): sc = 18
    elif re.search(r'大模型|AI|Agent|智能体|LLM|算法应用', t, re.I): sc = 14
    elif re.search(r'开发|工程师|程序', t): sc = 8
    else: sc = 0
    if sc: sc += min(3, len(re.findall(r'Vue|JavaScript|C#|Unity|HTML5|Java\b|MySQL', txt, re.I)))
    return min(sc, 25)

def size_up(s):
    s = S(s)
    if '以下' in s: return 10
    n = [int(x) for x in re.findall(r'\d+', s)]
    return max(n) if n else 0

def s_stable(r):
    up = size_up(r['规模'])
    sc = 14 if up >= 10000 else 12 if up >= 1000 else 10 if up >= 500 else 7 if up >= 100 else 4 if up >= 20 else 1
    if S(r['多平台']) not in ('', '1'): sc += 3
    if re.search(r'上市|不需要融资|[DEF]轮', S(r['融资阶段'])): sc += 3
    if S(r['外包嫌疑']): sc -= 12
    return round(max(0, min(sc, 20)) * 0.75)

def s_pay(r):
    m = re.search(r'(\d+)(?:-(\d+))?元/天', S(r['薪资']))
    if m:
        d = (int(m.group(1)) + int(m.group(2) or m.group(1))) / 2
        return 2 if d < 100 else 6 if d < 200 else 10 if d < 300 else 13
    lo, hi = r['薪资下限K'], r['薪资上限K']
    if pd.isna(lo) or pd.isna(hi): return 7 if '面议' in S(r['薪资']) else 5
    mid = (lo + hi) / 2
    return 3 if mid < 6 else 6 if mid < 8 else 9 if mid < 10 else 12 if mid < 14 else 15

def s_grow(r):
    sc = 5 if re.search(r'互联网|软件|计算机|游戏|人工智能|通信|IT', S(r['行业'])) else 2
    if size_up(r['规模']) >= 1000 or '上市' in S(r['融资阶段']): sc += 3
    if re.search(r'Vue|React|Unity|Python|大模型|Agent', S(r['岗位']) + S(r['技能标签']), re.I): sc += 2
    return min(sc, 10)

def outsourced(r):
    t = S(r['岗位']) + ' ' + S(r['公司'])
    hit = re.search(r'外包|派遣|外派|驻场|OD|德科|人力|劳务|猎头|项目外包', t, re.I)
    if hit: return '是(' + hit.group(0) + ')'
    if S(r['外包嫌疑']): return S(r['外包嫌疑'])
    # 大厂低薪的开发岗：多为外包/驻场
    if (r['是否实习'] != '是' and size_up(r['规模']) >= 10000 and pd.notna(r['薪资上限K'])
            and r['薪资上限K'] <= 12 and re.search(r'开发|工程师', S(r['岗位']))):
        return '疑似(大厂低薪)'
    return ''
df['外包嫌疑'] = df.apply(outsourced, axis=1).replace('', pd.NA)

def cohort(r):
    t = S(r['岗位']) + S(r['经验'])
    m = re.search(r'(?:20)?(2[5-9])\s*(?:届|春招|秋招|校招)', t)
    if not m: return ''
    y = int(m.group(1))
    return '27届及以后校招(不可投)' if y >= 27 else '26届春招/校招(可能已结束)' if y == 26 else ''
df['届别'] = df.apply(cohort, axis=1)

df['赛道'] = df.apply(track, axis=1)
df['毕业生可投'] = df.apply(grad_flag, axis=1)
df['排除标记'] = df.apply(lambda r: '排除' if exclude(r) else '', axis=1)
df['录取可能(35)'] = df.apply(s_enter, axis=1)
df['对口(25)'] = df.apply(s_fit, axis=1)
df['稳定口碑(15)'] = df.apply(s_stable, axis=1)
df['薪资(15)'] = df.apply(s_pay, axis=1)
df['成长(10)'] = df.apply(s_grow, axis=1)
df['总分'] = df[['录取可能(35)', '对口(25)', '稳定口碑(15)', '薪资(15)', '成长(10)']].sum(axis=1)
df.loc[df['外包嫌疑'].notna(), '总分'] = df['总分'].clip(upper=69)
df.loc[df['录取可能(35)'] <= 3, '总分'] = df['总分'].clip(upper=44)
df.loc[df['排除标记'] == '排除', '总分'] = 0

def grade(r):
    if r['排除标记']: return '排除'
    t = r['总分']
    if S(r['外包嫌疑']): t = min(t, 69)
    if r['录取可能(35)'] <= 3: t = min(t, 44)
    return 'A' if t >= 75 else 'B' if t >= 60 else 'C' if t >= 45 else 'D'
df['评级'] = df.apply(grade, axis=1)
df = df.sort_values('总分', ascending=False)

def key(c):
    c = re.sub(r'[（(].*?[）)]', '', S(c))
    c = re.sub(r'^(杭州|浙江|上海|北京|深圳|大连|苏州)', '', c)
    return re.sub(r'(股份)?有限(责任)?公司|集团|科技|网络|技术|信息', '', c)
df['_k'] = df['公司'].map(key)

rows = []
for (k, tk), g in df.groupby(['_k', '赛道']):
    g = g.sort_values('总分', ascending=False); b = g.iloc[0]
    rows.append({'公司': b['公司'], '最佳岗位': b['岗位'], '最佳分': b['总分'],
                 '前3均分': round(g['总分'].head(3).mean(), 1), '在招岗位数': len(g),
                 '赛道': tk, '最高评级': b['评级'],
                 '外包岗数/总数': f"{int(g['外包嫌疑'].notna().sum())}/{len(g)}", '届别': b['届别'], '平台': '/'.join(sorted(set(g['平台']))),
                 '规模': b['规模'], '行业': b['行业'], '薪资': b['薪资'], '毕业生可投': b['毕业生可投'],
                 '最佳岗位链接': b['链接'], '待调研(口碑/风险)': ''})
co = pd.DataFrame(rows).sort_values(['最佳分', '在招岗位数'], ascending=False)

rules = pd.DataFrame({'维度': ['录取可能(35)', '对口(25)', '稳定口碑(15)', '薪资(15)', '成长(10)', '排除', '评级', '说明'], '规则': [
    '【录取可能性】经验不限/应届/在校/1年以内=35；高级/资深/专家/经理类封顶3；正式岗大厂(万人+)封顶14、千人级封顶22(竞争激烈)；1-3年=8；2年=3；3年+=0；硕士博士=0；统招要求-3；实习岗仅限在校生(学生可投/在校/27届等)最多6分',
    'Unity/游戏客户端=25；前端/Vue=22；实施/交付=20；C#/.NET=18；大模型/AI应用=14；其他开发=8；另按技能标签重合最多+3',
    '(最终×0.75)规模(万人+=14,千人=12,500+=10,100+=7,20+=4)；多平台在招+3；上市/不需融资/D轮+=+3；外包/派遣/OD嫌疑-12，且总分最高69(含"大厂月薪上限≤12K"的疑似外包)',
    '正式岗按月薪中值分档(<6K=3 … 14K+=15)；实习按元/天分档；面议7；无法解析5',
    '互联网/软件/游戏/AI等行业5；大厂或上市+3；技术栈含Vue/React/Unity/Python/大模型+2',
    '售前/销售/运营/客服/讲师/招生/地推/培训/审核/达人/主播 直接0分',
    'A>=75，B 60-74，C 45-59，D<45',
    '稳定口碑目前只是公开数据的估算，真实口碑和风险待调研后校正']})

out = BASE + '杭州岗位评分_v2_20261023.xlsx'
cols = ['评级', '总分', '赛道', '毕业生可投', '届别', '平台', '公司', '岗位', '薪资', '经验', '学历', '规模', '行业', '录取可能(35)', '对口(25)',
        '稳定口碑(15)', '薪资(15)', '成长(10)', '外包嫌疑', '排除标记', '区域', '技能标签', '链接']
with pd.ExcelWriter(out, engine='openpyxl') as w:
    df[cols].to_excel(w, sheet_name='岗位评分', index=False)
    co.to_excel(w, sheet_name='公司评分', index=False)
    rules.to_excel(w, sheet_name='评分规则', index=False)
wb = load_workbook(out)
for ws in wb:
    ws.freeze_panes = 'A2'
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F4E79')
    for i, col in enumerate(ws.columns, 1):
        ml = max((len(str(c.value)) if c.value is not None else 0) for c in list(col)[:80])
        ws.column_dimensions[get_column_letter(i)].width = min(max(8, ml * 1.6), 50)
    if ws.title != '评分规则': ws.auto_filter.ref = ws.dimensions
    else: ws.column_dimensions['B'].width = 120
wb.save(out)

print('评级分布(按赛道):'); print(pd.crosstab(df['赛道'], df['评级']))
print('实习岗毕业生可投:', df[df['赛道'] == '实习']['毕业生可投'].value_counts().to_dict())
show = ['公司', '最佳岗位', '薪资', '规模', '最佳分', '在招岗位数', '外包岗数/总数', '毕业生可投', '届别']
print('\n=== 正式岗 公司 TOP25'); print(co[co['赛道'] == '正式'][show].head(25).to_string(index=False))
print('\n=== 实习岗 公司 TOP15'); print(co[co['赛道'] == '实习'][show].head(15).to_string(index=False))
