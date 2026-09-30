# 真机验证脚手架（Playwright 加载未打包扩展）

不属于扩展产物，仅用于人工/自动回归验证。可随时删除。

## 构建：只用 `npm run dev`，别用 `wxt`（serve）

```bash
npm run dev      # = wxt build -m development + scripts/check-build.mjs（产物自检）
npm run build    # 生产产物 + 自检
npm run check    # 只跑自检
```

**不要**跑 `npx wxt` / `wxt dev`。serve 模式的 manifest **故意不含 `content_scripts`**
（改由它自己的 background 运行时注册），而 `wxt build` 正好相反。两者写进同一个
`.output/chrome-mv3-dev` 就会杂交出「manifest 说没有内容脚本、background 又不去注册」的产物 →
内容脚本永不执行，**页面上面板凭空消失、控制台一行日志都没有**。
（2026-09-25 真实踩过：dev server 在跑时跑了一次 `wxt build -m development`。
`wxt.config.ts` 里已有 `server:started` 钩子会告警，`scripts/check-build.mjs` 会拦截。）

## 一次性准备

```bash
cd "C:/Users/13027/.workbuddy/binaries/node/workspace"
npm init -y && npm install playwright-core
```

## 跑

```bash
cd "C:/Users/13027/.workbuddy/binaries/node/workspace"
cp "D:/soft/JetBrains/workspace/dianxiaomi/.dev-verify/"*.cjs .
cp "D:/soft/JetBrains/workspace/dianxiaomi/.dev-verify/"*.html .
DIST="D:/soft/JetBrains/workspace/dianxiaomi/.output/chrome-mv3-dev" node dxm_probe.cjs
```

| 场景 | 命令 |
| --- | --- |
| **最小诊断**：面板挂没挂 + 全部 console（排查「什么都没发生」先跑这个） | `node dxm_panel_diag.cjs` |
| A 顶层账号区（真实 URL） | `node dxm_probe.cjs` |
| B 账号区在同源 iframe | `MOCK=dxm_mock_iframe.html node dxm_probe.cjs` |
| C 完全没有账号区（应中止导出） | `NOACCOUNT=1 DL_TIMEOUT=12000 node dxm_probe.cjs` |
| 有头观察 | `HEADFUL=1 node dxm_probe.cjs` |
| 验证 prod 产物 | `DIST=.../.output/chrome-mv3 node dxm_probe.cjs` |

探针会依次断言：① 面板挂载 + 账号识别 → ② 面板上**没有**日期范围选择器
（`noDateRangeUI=true`，导出的是列表全部数据）→ ③ `✕` 只收起不删节点、且页面内
**没有**圆形开关（`ballGone=true`）→ ④ 弹窗里的开关（点一次收起、再点一次展开，并读回按钮文案）
→ ⑤ 导出下载（文件名 + 字节数）→ ⑥ 收尾快照（面板全量可见文案 `allText`、`diagEntryGone` 等）。

## 注意

- 必须带 `--disable-features=DisableLoadExtensionCommandLineSwitch`（否则 Chrome 忽略 `--load-extension`）。
- 别用 `chrome --headless --screenshot`，本机必崩（network service / GPU），只能用 Playwright。
- 改了代码后 **dev + prod 两份产物都要重建**；构建标记（`CONFIG.buildTag`，构建期自动生成，
  如 `B0925-1610`）用于确认浏览器里跑的是哪一版，**界面上不显示**，
  读 `#__dxm_export_panel` 的 `data-build`（探针输出里的 `buildTag` 字段）。
- 「面板不见了」先看 **content script 有没有打出第一行日志**：
  完全没日志 = 没被注入（查 manifest 的 content_scripts），有日志 = 运行时报错（查 stack）。
