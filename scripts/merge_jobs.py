# -*- coding: utf-8 -*-
"""合并 Boss / 智联 / 猎聘 杭州岗位原始数据 -> 杭州岗位汇总_20261023.xlsx
raw 目录只读；输出写到 数据/ 下。运行: python -I merge_jobs.py 。"""
import os
WORK = os.environ.get('JOBHUNT_HOME', './workspace')  # 工作目录：data/raw 放抓取原始数据，其余是输出
import sys, os, re, json, glob, statistics, collections
import pandas as pd
from openpyxl.utils import get_column_letter

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

BASE = WORK + "/data"
RAW = BASE + "/raw"
OUT = BASE + "/杭州岗位汇总_20261023.xlsx"
problems = []

# ---------------- 通用规范化 ----------------
def norm_salary(s):
    """返回 (下限K, 上限K)；元/天、面议、无法解析 -> (None, None)"""
    if not s:
        return None, None
    s = str(s).replace(" ", "")
    if "面议" in s or "/天" in s or "元/时" in s or "/小时" in s:
        return None, None
    core = s.split("·")[0]
    m = re.match(r"^([\d.]+)(?:-([\d.]+))?(k|K|千|万|元)?", core)
    if not m:
        problems.append(f"薪资无法解析: {s}")
        return None, None
    lo = float(m.group(1)); hi = float(m.group(2)) if m.group(2) else lo
    unit = m.group(3) or ""
    if unit == "万":
        lo, hi = lo * 10, hi * 10
    elif unit == "元" or (unit == "" and lo >= 1000):
        lo, hi = lo / 1000, hi / 1000
    return round(lo, 1), round(hi, 1)

COUNTY = {"建德": "建德市", "桐庐": "桐庐县", "淳安": "淳安县"}
def norm_area(d):
    d = (d or "").strip()
    if not d:
        return ""
    if d.endswith(("区", "市", "县")):
        return d
    return COUNTY.get(d, d + "区")

CO_PREFIX = r"^(杭州市?|浙江省?|中国|上海|北京|深圳)"
CO_SUFFIX = r"(股份有限公司|有限责任公司|有限公司|集团|股份|科技|网络|信息技术|技术|智能|软件|信息|服务)$"
def co_key(name):
    k = (name or "").strip()
    k = re.sub(r"[（(][^）)]*[）)]", "", k)          # 去括号 (杭州)/(深圳)
    for _ in range(4):
        k2 = re.sub(CO_SUFFIX, "", k)
        k2 = re.sub(CO_PREFIX, "", k2)
        if k2 == k:
            break
        k = k2
    k = re.sub(r"[\s|｜·.]", "", k).lower()
    return k or (name or "").strip().lower()

def title_key(t):
    t = (t or "").lower()
    t = re.sub(r"【[^】]*】|\[[^\]]*\]", "", t)
    t = re.sub(r"[（(]\s*j?\d+\s*[）)]|-?j\d{3,}", "", t)   # 岗位编号 (J11215) / -J15244
    t = re.sub(r"[\s/／\-—_（）()·,，、]", "", t)
    return t

OK_EXP = ("经验不限", "在校/应届", "应届", "1年以内", "1年以下", "无经验", "不限", "")
def exp_base(exp):
    e = (exp or "").strip()
    if e in OK_EXP or "应届" in e or "在校" in e:
        return 1
    m = re.match(r"(\d+)", e)
    if m:
        return 0 if int(m.group(1)) < 3 else -9
    return 0

MATCH_RE = re.compile(r"前端|vue|unity|c#|实施|游戏客户端|ai应用|大模型应用|agent", re.I)
BAD_RE = re.compile(r"销售|售前|运营|客服|讲师|培训|招生")
OUTS_RE = re.compile(r"外派|外包|人力|劳务")

def score(r):
    s = exp_base(r["经验"])
    if MATCH_RE.search(r["岗位"] + " " + r["技能标签"]):
        s += 1
    edu = r["学历"]
    if "本科" in edu or "大专" in edu or "不限" in edu:
        s += 1
    if r["是否实习"] == "是" or BAD_RE.search(r["岗位"]):
        s -= 9
    return max(0, min(3, s))

def row(**kw):
    base = dict(平台="", 搜索词="", 岗位="", 公司="", 薪资="", 经验="", 学历="", 规模="", 行业="",
                融资阶段="", 区域="", 技能标签="", 发布时间="", 链接="", _proxy=False, _intern=False)
    base.update(kw)
    return base

rows = []
# ---------------- Boss ----------------
boss = json.load(open(RAW + "/boss_hz_raw_20261023.json", encoding="utf-8"))
raw_by_jid = {x.get("encryptJobId"): x for v in boss.get("raw", {}).values() for x in v}
empty_q = [k for k, v in boss.get("raw", {}).items() if not v]
if empty_q:
    problems.append("Boss 只有『前端开发』有数据，以下搜索词为空: " + "、".join(empty_q))
for x in boss["flat"]:
    rx = raw_by_jid.get(x["jid"], {})
    parts = (x.get("area") or "").split("-")
    rows.append(row(平台="Boss直聘", 搜索词=x["q"], 岗位=x["job"], 公司=x["co"], 薪资=x["sal"],
                    经验=x["exp"], 学历=x["edu"], 规模=x["size"], 行业=x["ind"], 融资阶段=x["stage"],
                    区域=norm_area(parts[1] if len(parts) > 1 else ""), 技能标签=x["skills"],
                    链接=f"https://www.zhipin.com/job_detail/{x['jid']}.html",
                    _proxy=bool(rx.get("proxyJob"))))

# ---------------- 智联 ----------------
for f in sorted(glob.glob(RAW + "/zhaopin_hz_*_20261023.json")):
    if "slim_stash" in f:
        continue
    d = json.load(open(f, encoding="utf-8"))
    for x in d.get("list", []):
        tags = [t.get("value") for t in (x.get("skillLabel") or []) if t.get("value")]
        if not tags:
            tags = [t.get("name") for t in (x.get("jobSkillTags") or []) if t.get("name")]
        rows.append(row(平台="智联招聘", 搜索词=d["kw"], 岗位=x.get("name", ""), 公司=x.get("companyName", ""),
                        薪资=x.get("salary60", ""), 经验=x.get("workingExp", ""), 学历=x.get("education", ""),
                        规模=x.get("companySize", ""), 行业=x.get("industryName", ""),
                        融资阶段=x.get("financingStage") or "", 区域=norm_area(x.get("cityDistrict")),
                        技能标签="/".join(tags), 发布时间=x.get("publishTime", ""),
                        链接=x.get("positionURL") or x.get("positionUrl") or "",
                        _proxy=bool(x.get("rpoProxy") or x.get("rpoProxied")),
                        _intern=(x.get("workType") == "实习")))

# ---------------- 猎聘 ----------------
SAL_RE = re.compile(r"薪|[kK]|万|元|面议")
EDU_RE = re.compile(r"本科|大专|硕士|博士|学历不限|统招|高中|中专")
EXP_RE = re.compile(r"年|经验不限|实习|应届")
SIZE_RE = re.compile(r"\d+-\d+人|\d+人以上|\d+人以下")
STAGE_RE = re.compile(r"(沪深A股上市|创业板上市|科创板上市|港股上市|美股上市|北交所上市|新三板|已上市|"
                      r"天使轮|Pre-?A轮|A\+?轮|B\+?轮|C\+?轮|D轮及以上|战略融资|融资未公开|不需要融资|未融资)$")
for f in sorted(glob.glob(RAW + "/liepin_hz_*_20261023.json")):
    d = json.load(open(f, encoding="utf-8"))
    if "page" not in d:
        problems.append(f"猎聘 {os.path.basename(f)} 无 page 字段且 list 为空（早期空抓取），跳过")
    for it in d.get("list", []):
        L = [s.strip() for s in it.get("lines", []) if s.strip()]
        if len(L) < 7 or L[1] != "【":
            problems.append(f"猎聘卡片结构异常: {L}")
            continue
        title, area = L[0], L[2]
        rest = L[4:]  # 】之后
        sal = rest[0] if SAL_RE.search(rest[0]) else ""
        last, company = rest[-1], rest[-2]
        mids = rest[1 if sal else 0:-2]
        edu = next((s for s in mids if EDU_RE.search(s)), "")
        exps = [s for s in mids if EXP_RE.search(s) and s != edu]
        exp = "实习" if "实习" in exps else (exps[0] if exps else "")
        size_m = SIZE_RE.search(last)
        size = size_m.group(0) if size_m else ""
        rem = last[:size_m.start()] if size_m else last
        st_m = STAGE_RE.search(rem)
        stage = st_m.group(1) if st_m else ""
        ind = rem[:st_m.start()] if st_m else rem
        if not sal:
            problems.append(f"猎聘未识别薪资: {L}")
        ap = area.split("-")
        rows.append(row(平台="猎聘", 搜索词=d["kw"], 岗位=title, 公司=company, 薪资=sal, 经验=exp, 学历=edu,
                        规模=size, 行业=ind, 融资阶段=stage, 区域=norm_area(ap[1] if len(ap) > 1 else ""),
                        链接=it.get("href", ""), _intern=("实习" in exps)))

df = pd.DataFrame(rows)
before = df["平台"].value_counts().to_dict()

# ---------------- 派生列 ----------------
df[["薪资下限K", "薪资上限K"]] = df["薪资"].apply(lambda s: pd.Series(norm_salary(s)))
df["是否实习"] = df.apply(lambda r: "是" if (r["_intern"] or "实习" in r["岗位"] or "实习" in r["经验"]) else "否", axis=1)
df["适合度"] = df.apply(score, axis=1)
def outs_reason(r):
    why = []
    if "人力资源" in r["行业"]:
        why.append("人力资源行业")
    m = OUTS_RE.search(r["岗位"] + r["公司"])
    if m:
        why.append("含" + m.group(0))
    if r["_proxy"]:
        why.append("平台标记代招")
    return "是(" + "/".join(why) + ")" if why else ""
df["外包嫌疑"] = df.apply(outs_reason, axis=1)
df["_ck"] = df["公司"].map(co_key)
df["_tk"] = df["岗位"].map(title_key)

# 平台内去重：同 (公司key, 岗位key) 保留第一条，搜索词合并
agg_q = df.groupby(["平台", "_ck", "_tk"])["搜索词"].agg(lambda s: "/".join(dict.fromkeys(s)))
df = df.drop_duplicates(["平台", "_ck", "_tk"], keep="first").copy()
df["搜索词"] = [agg_q[(p, c, t)] for p, c, t in zip(df["平台"], df["_ck"], df["_tk"])]
after = df["平台"].value_counts().to_dict()
df["多平台"] = df.groupby(["_ck", "_tk"])["平台"].transform("nunique")

COLS = ["平台", "搜索词", "岗位", "公司", "薪资", "薪资下限K", "薪资上限K", "经验", "学历", "规模", "行业",
        "融资阶段", "区域", "技能标签", "发布时间", "链接", "是否实习", "适合度", "外包嫌疑", "多平台"]
detail = df.sort_values(["适合度", "多平台", "薪资上限K"], ascending=[False, False, False])[COLS]

# ---------------- 公司汇总 ----------------
def first_nonempty(s):
    for v in s:
        if v:
            return v
    return ""

comp = []
for ck, g in df.groupby("_ck"):
    names = sorted(set(g["公司"]), key=len, reverse=True)
    titles = "; ".join(dict.fromkeys(g["岗位"]))
    lo, hi = g["薪资下限K"].min(), g["薪资上限K"].max()
    comp.append({
        "公司": names[0] + (f"（又名: {'/'.join(names[1:])}）" if len(names) > 1 else ""),
        "在招岗位数": len(g),
        "平台列表": "/".join(dict.fromkeys(g["平台"])),
        "岗位列表": titles[:200] + ("…" if len(titles) > 200 else ""),
        "规模": first_nonempty(g["规模"]), "行业": first_nonempty(g["行业"]),
        "融资阶段": first_nonempty(g["融资阶段"]),
        "区域": "/".join(dict.fromkeys(a for a in g["区域"] if a)),
        "薪资范围": (f"{lo:g}-{hi:g}K" if pd.notna(lo) and pd.notna(hi) else ""),
        "最高适合度": int(g["适合度"].max()),
        "外包嫌疑": "是" if g["外包嫌疑"].str.startswith("是").any() else "",
    })
comp = pd.DataFrame(comp).sort_values(["最高适合度", "在招岗位数"], ascending=[False, False])

# ---------------- 统计 ----------------
st_q = df.assign(搜索词=df["搜索词"].str.split("/")).explode("搜索词").groupby("搜索词").agg(
    岗位数=("岗位", "size"), 薪资下限中位数=("薪资下限K", "median"), 薪资上限中位数=("薪资上限K", "median")
).reset_index().sort_values("岗位数", ascending=False)
tagc = collections.Counter(t.strip() for s in df["技能标签"] for t in re.split(r"[/、,，]", s) if t.strip())
st_tag = pd.DataFrame(tagc.most_common(30), columns=["技能标签", "频次"])
st_size = df["规模"].replace("", "(空)").value_counts().rename_axis("规模").reset_index(name="岗位数")
st_area = df["区域"].replace("", "(空)").value_counts().rename_axis("区域").reset_index(name="岗位数")
st_fit = df["适合度"].value_counts().sort_index(ascending=False).rename_axis("适合度").reset_index(name="岗位数")

# ---------------- 写 Excel ----------------
def autofit(ws, maxw=60):
    for col in ws.columns:
        w = 4
        for c in list(col)[:400]:
            if c.value is not None:
                v = str(c.value)
                w = max(w, sum(2 if ord(ch) > 127 else 1 for ch in v) + 2)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w, maxw)

with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
    detail.to_excel(xw, sheet_name="岗位明细", index=False)
    comp.to_excel(xw, sheet_name="公司汇总", index=False)
    r = 0
    blocks = [("按搜索词", st_q), ("适合度分布", st_fit), ("技能标签 Top30", st_tag),
              ("规模分布", st_size), ("区域分布", st_area)]
    # 横向排列各表，第 1 行标题、第 2 行表头
    c = 0
    for name, t in blocks:
        t.to_excel(xw, sheet_name="统计", index=False, startrow=1, startcol=c)
        xw.sheets["统计"].cell(row=1, column=c + 1, value=name)
        c += len(t.columns) + 1
    for name in ("岗位明细", "公司汇总"):
        ws = xw.sheets[name]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autofit(ws)
    ws = xw.sheets["统计"]
    ws.freeze_panes = "A3"
    autofit(ws, 30)

# ---------------- 报告 ----------------
print("去重前:", before)
print("去重后:", after)
print("公司数:", len(comp))
print("适合度分布:", df["适合度"].value_counts().sort_index().to_dict())
print("外包嫌疑岗位:", df["外包嫌疑"].str.startswith("是").sum(), " 跨平台重复(多平台>1)岗位行:", (df["多平台"] > 1).sum())
print("Top15 公司:")
print(comp.head(15)[["公司", "规模", "在招岗位数", "薪资范围", "最高适合度", "外包嫌疑"]].to_string(index=False))
print("问题:")
for p in dict.fromkeys(problems):
    print(" -", p)
print("输出:", OUT)
