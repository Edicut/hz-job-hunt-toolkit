# 抓取片段（在已登录的浏览器页面里运行，数据发给本机接收器）

> 仅供个人求职调研。遵守各平台的使用条款，不要高频请求，不要绕过验证码，不要自动投递。

先启动接收器：`python scripts/local_receiver.py`（监听 `127.0.0.1:8765`，POST 的 JSON 存到 `workspace/data/raw/`；`workspace/serve/` 里的文件可被页面 GET，用来下发网址清单）。

## 读取已抓取页面的岗位原文（通用）
```js
window.__fr = {done:false, out:[], saved:null};
(async () => {
  const L = await (await fetch('http://127.0.0.1:8765/urls.json')).json();   // [[公司, 岗位, 链接], ...]
  const sl = ms => new Promise(r => setTimeout(r, ms));
  for (const [co, job, u] of L) {
    try {
      const h = await (await fetch(u, {credentials:'include'})).text();
      const d = new DOMParser().parseFromString(h, 'text/html');
      d.querySelectorAll('script,style').forEach(e => e.remove());
      const t = d.body.textContent.replace(/\s+/g, ' ');
      let i = t.search(/职位描述|职位介绍|岗位职责|工作职责/); if (i < 0) i = 0;
      window.__fr.out.push({co, job, u, text: t.slice(i, i + 2600)});
    } catch (e) { window.__fr.out.push({co, job, u, err: String(e)}); }
    await sl(2200 + Math.random() * 2200);          // 随机间隔，别太快
  }
  window.__fr.done = true;
  const r = await fetch('http://127.0.0.1:8765/research_xxx.json',
    {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(window.__fr)});
  window.__fr.saved = r.status;
})();
```
- 单次脚本执行有时限（约45秒），**脚本在页面后台继续跑**，之后用 `window.__fr.out.length` 查进度。
- 同一个来源的页面才能带登录态 `fetch`，所以智联/猎聘要各开一个标签页分别跑。

## 招聘页搜索结果
- 前端状态型页面：找到页面框架的 store（如 Vue 的 `$store.state.<list>`），脚本里不断滚动加载，把列表深拷贝后 POST 出去。
- 服务端渲染型页面：按页码参数逐页 `fetch`，用 `DOMParser` 解析卡片文字。
- 注意控制每个关键词的页数和请求间隔，避免触发风控。
