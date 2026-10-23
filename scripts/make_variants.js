// 生成 4 个简历版本：通用 / Unity游戏 / 前端 / AI实施(Agent)
// 事实全部来自 data.js；这里只调整：求职意向、技能顺序、项目顺序与详略、板块顺序、自我评价。
// 红字【】= 待他确认或待补充；没做出来的项目只留红色占位，不写成已完成。
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, LevelFormat, TabStopType,
} = require("docx");
const BASE = require("./data.example.js");  // 换成你自己的 data.js

const FONT = "微软雅黑", ACCENT = "1F4E79";
const clone = o => JSON.parse(JSON.stringify(o));
const skill = k => BASE.skills.find(s => s[0].startsWith(k));
const proj = k => BASE.projects.find(p => p.left.includes(k));
const pick = keys => keys.map(skill).filter(Boolean);

// ---------- 各版本 ----------
const V = {};

V["通用"] = clone(BASE);

// Unity / 游戏
V["Unity游戏"] = (() => {
  const d = clone(BASE);
  d.intent = "求职意向：Unity 客户端开发 / 游戏开发工程师  |  【毕业年份】  |  【城市】";
  d.skills = pick(["Unity", "编程基础", "前端", "后端", "AI", "工具"]);
  const tank = clone(proj("游戏"));
  tank.bullets = [
    ...tank.bullets,
    "【写 1～2 个实际用到的 Unity 技术点，例如：对象池 / 状态机 / 协程 / UGUI / 射线检测，只写真的用过的】",
    "【可演示：GitHub 仓库链接 + 30 秒游戏录屏】",
  ];
  const vote = clone(proj("Web"));
  vote.bullets = vote.bullets.slice(0, 2);
  d.projects = [tank, vote];
  d.other = d.other.filter(s => !s.includes("网络安全"));
  d.summary = "【按该版本方向写 2-3 句自我评价，只写真实的事】";
  return d;
})();

// 前端
V["前端"] = (() => {
  const d = clone(BASE);
  d.intent = "求职意向：前端开发工程师（实习或初级岗）  |  【毕业年份】  |  【城市】";
  d.skills = pick(["前端", "后端", "编程基础", "Unity", "AI", "工具"]);
  const vote = clone(proj("Web"));
  vote.bullets = [
    ...vote.bullets,
    "【前端部分：写你真正做的页面与交互，例如 Bootstrap 响应式布局、表单校验、结果页数据展示】",
  ];
  const tank = clone(proj("游戏"));
  tank.left = tank.left + "（交互与图形方向）";
  tank.bullets = tank.bullets.slice(0, 2);
  const slot = { left: "【Vue 个人项目槽位：做完再填，没做完不写】", right: "",
    bullets: ["计划：用 Vue3 + Vite 做一个带接口调用的小项目（如大模型对话页面），上传 GitHub"] };
  d.projects = [vote, slot, tank];
  d.summary = "【按该版本方向写 2-3 句自我评价，只写真实的事】";
  return d;
})();

// AI / 实施（需先做完 Agent 项目再投）
V["AI实施"] = (() => {
  const d = clone(BASE);
  d.warn = "【投递前删除本行：本版需先做完 Dify / Agent 项目并把它写进项目经历，没做完不要投】";
  d.intent = "求职意向：AI 应用开发 / AI 应用实施工程师（初级）  |  【毕业年份】  |  【城市】";
  d.skills = pick(["AI", "后端", "前端", "编程基础", "工具", "Unity"]);
  const agent = { left: "【Agent 项目：某领域知识问答 Agent（做完再填）】", right: "【2026.xx — 2026.xx】",
    bullets: [
      "【技术栈：Dify + DeepSeek 或通义千问 API + Vue3；整理某领域资料构建知识库】",
      "【设计 system prompt 限定回答范围并给出出处；建立 20 题测试集，调整切片与检索参数，准确率由 X% 提升到 Y%】",
      "【用 Docker 部署到云服务器，公网可访问；记录部署中踩过的坑与解决办法】",
    ] };
  const vote = clone(proj("Web"));
  vote.bullets = vote.bullets.slice(0, 2);
  d.projects = [agent, vote];
  d.order = ["education", "skills", "work", "projects", "other", "summary"]; // 实施岗：经历放项目前
  d.summary = "【按该版本方向写 2-3 句自我评价，只写真实的事】";
  return d;
})();

// ---------- 渲染 ----------
function runs(text, o = {}) {
  return text.split(/(【[^】]*】)/).filter(Boolean).map(t =>
    new TextRun({ text: t, font: FONT, size: o.size || 19, bold: o.bold,
      color: t.startsWith("【") ? "C00000" : (o.color || "222222") }));
}
const section = t => new Paragraph({ spacing: { before: 100, after: 40 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 1 } },
  children: [new TextRun({ text: t, font: FONT, size: 24, bold: true, color: ACCENT })] });
const headLine = (l, r) => new Paragraph({ spacing: { before: 60, after: 20 },
  tabStops: [{ type: TabStopType.RIGHT, position: 10466 }],
  children: [...runs(l, { bold: true, size: 21 }), new TextRun({ text: "\t", font: FONT }), ...runs(r, { size: 20, color: "555555" })] });
const bullet = t => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 20, line: 226 }, children: runs(t) });
const block = items => items.flatMap(i => [headLine(i.left, i.right), ...i.bullets.map(bullet)]);

function build(d) {
  const sec = {
    education: () => [section("教育背景"), ...block(d.education)],
    skills: () => [section("专业技能"), ...d.skills.map(([k, v]) => new Paragraph({ numbering: { reference: "b", level: 0 },
      spacing: { after: 20, line: 226 }, children: [...runs(k + "：", { bold: true }), ...runs(v)] }))],
    projects: () => [section("项目经历"), ...block(d.projects)],
    work: () => [section("实习经历"), ...block(d.work)],
    other: () => [section("其他经历"), ...d.other.map(bullet)],
    summary: () => [section("自我评价"), new Paragraph({ spacing: { line: 226 }, children: runs(d.summary) })],
  };
  const order = d.order || ["education", "skills", "projects", "work", "other", "summary"];
  const head = [
    ...(d.warn ? [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: runs(d.warn, { size: 20, bold: true }) })] : []),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: runs(d.name, { size: 40, bold: true, color: ACCENT }) }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: runs(d.contact, { size: 19, color: "444444" }) }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: runs(d.intent, { size: 20, bold: true }) }),
  ];
  return new Document({
    numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 300, hanging: 220 } } } }] }] },
    sections: [{ properties: { page: { margin: { top: 560, bottom: 520, left: 700, right: 700 } } },
      children: [...head, ...order.flatMap(k => sec[k]())] }],
  });
}

(async () => {
  fs.mkdirSync("./resume-output", { recursive: true });
  for (const [name, d] of Object.entries(V)) {
    const out = `./resume-output/简历-${name}版.docx`;
    try { fs.writeFileSync(out, await Packer.toBuffer(build(d))); console.log("ok", out); }
    catch (e) { console.log("FAIL", out, e.code); }
  }
})();
