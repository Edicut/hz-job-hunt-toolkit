import os,sys,re,json,glob; sys.stdout.reconfigure(encoding='utf-8')
import warnings; warnings.filterwarnings('ignore')
import pandas as pd
d=pd.read_excel(os.environ.get('JOBHUNT_HOME','./workspace')+'/data/杭州岗位评分_v2_20261023.xlsx',sheet_name='岗位评分')
S=lambda v:'' if pd.isna(v) else str(v)
print('全部：',len(d),'个岗位',d['公司'].nunique(),'家公司')
nontech=r'漫剧|短剧|视频|美工|导演|编剧|剪辑|主播|直播|运营|文案|设计师|制作师|生成师|BD|商务|销售|客服|培训|讲师|审核|标注|招聘|猎头|产品助理|售前|地推|招生|美术|特效|动效|原画|建模|策划'
x=d[d['排除标记'].isna()&d['外包嫌疑'].isna()]
x=x[~x['岗位'].str.contains(nontech,regex=True,na=False)]
x=x[~x['岗位'].str.contains(r'2[789]届|2[789]秋招|2[789]春招|日语|派遣|外派',regex=True,na=False)]
x=x[~x['学历'].fillna('').str.contains('硕士|博士')]
x=x[x['岗位'].str.contains(r'开发|工程师|研发|程序|全栈|前端|后端|实施|Unity|U3D|客户端|Agent|智能体|测试',regex=True,na=False)]
print('去掉：外包/非技术/27届/日语/硕博/无关岗位 后：',len(x),'个岗位',x['公司'].nunique(),'家')
exp=x['经验'].fillna('')
hard=exp.str.contains(r'3-5年|5-10年|10年以上|3年以上|5年以上')
x2=x[~hard]
print('再去掉：招聘页明确要3年以上 后：',len(x2),'个岗位',x2['公司'].nunique(),'家')
low=x2[x2['经验'].fillna('').str.contains('不限|应届|在校|1年以[下内]|实习|无')]
mid=x2[x2['经验'].fillna('').str.contains('1-3年')]
print('  其中招聘页写"不限/应届/1年以内/实习"：',len(low),'岗',low['公司'].nunique(),'家')
print('  其中写"1-3年"（小公司可海投）：',len(mid),'岗',mid['公司'].nunique(),'家')
def cat(t):
    if re.search(r'Unity|U3D|游戏',t,re.I): return 'Unity/游戏'
    if re.search(r'前端|Vue|Web|H5',t,re.I): return '前端'
    if re.search(r'实施|交付',t): return '实施'
    if re.search(r'AI|大模型|Agent|智能体',t,re.I): return 'AI'
    return '其他开发/测试'
x2=x2.copy(); x2['方向']=x2['岗位'].map(cat)
print(x2.groupby('方向').agg(岗位=('岗位','size'),公司=('公司','nunique')).to_string())
big=x2['规模'].fillna('').str.contains('10000人以上|5000-10000|1000-9999')
print('规模>=1000人的公司：',x2[big]['公司'].nunique(),'；100-999人：',x2[x2['规模'].fillna('').str.contains('100-299|100-499|300-499|500-999')]['公司'].nunique())
