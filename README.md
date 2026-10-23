# hz-job-hunt-toolkit

一次真实求职辅导的**方法与脚本**（已脱敏）：给一位应届计算机毕业生，在一座城市里做岗位调研 → 评分 → 读原文核验 → 建海投池 → 按方向做多版本简历 → 半自动投递。

> 目标不是"批量刷岗位"，而是：**不编造经历**、尽量多拿到真实面试机会、把时间花在能进的岗位上。

## 流程

```
抓取(多平台抽样) → 合并去重 → 机器评分 → 按方向分类 → 读岗位原文核验 → 查公司真实性 → 海投池(冲/稳/保) → 多版本简历 → 半自动投递+跟踪
```

示例漏斗（一次实测，数字仅供参考）：抓到约1000个岗位/700家公司 → 去掉外包、非技术岗、27届、日语、硕博等后约600 → 去掉要3年以上的约490 → 读原文再踢掉约40个特别差的 → 约200个岗位进海投池，分成每批25个。

## 目录

| 路径 | 作用 |
|---|---|
| `scripts/local_receiver.py` | 本机接收器：浏览器页面把抓到的数据 POST 到 `127.0.0.1`，避开浏览器的批量下载拦截 |
| `scripts/merge_jobs.py` | 把各平台原始数据合并、去重、解析薪资，输出 Excel |
| `scripts/score_jobs.py` | 岗位评分（录取可能性/对口/稳定口碑/薪资/成长），输出岗位与公司评分表 |
| `scripts/funnel_report.py` | 统计"从全部到可投"的漏斗和各方向数量 |
| `scripts/build_pool.py` | 构建海投池：过滤、分方向、分层（精投/海投/冲）、分批（每批25个）、带跟踪列 |
| `scripts/screen_pool_by_jd.py` | 读岗位原文，踢掉特别差的并写明原因 |
| `scripts/make_variants.js` | 从同一份数据生成多个方向的简历（docx） |
| `examples/data.example.js` | 简历示例数据（虚构） |
| `docs/scoring.md` | 评分权重与规则 |
| `docs/jd-screening-rules.md` | 岗位原文核验规则与公司真实性查法 |
| `docs/lessons.md` | 踩坑与经验（抓取、风控、验证码、数据可信度） |
| `docs/scraping-snippets.md` | 页面内抓取片段 |
| `docs/resume-variants.md` | 多版本简历策略 |
| `docs/learning-roadmap.md` | 面向"AI应用/前端/游戏客户端/实施"方向的自学路线 |

## 使用
```bash
export JOBHUNT_HOME=./workspace      # Windows: set JOBHUNT_HOME=.\workspace
python scripts/local_receiver.py     # 另开终端，页面脚本把数据发到这里
python scripts/merge_jobs.py
python scripts/score_jobs.py
python scripts/build_pool.py         # 可选: workspace/blacklist.txt、whitelist.txt（公司关键词，一行一个）
python scripts/screen_pool_by_jd.py  # 需要先用页面脚本抓好岗位原文
```
依赖：Python 3 + pandas + openpyxl；简历生成需要 Node.js + `docx`（`npm i docx`）。评分与过滤规则里写的是一个**示例画像**（应届、0年经验、游戏/前端/实施方向），请按自己的情况改。

## 注意
- 仅供个人求职调研；请遵守各招聘平台使用条款，不要高频请求、不要绕过验证码、不要自动投递。
- 仓库里**不含**任何真实个人信息、公司名单和抓取的数据。
- 机器评分只是第一轮筛选，真正的判断靠读原文和查公司。

License: MIT
