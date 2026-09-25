# 店小秘发货导出助手（Chrome MV3 扩展）

在店小秘「发货成功列表」页面注入悬浮面板，按日期范围（当天 / 近两天 / 近三天）把列表导出成 Excel。
导出字段：**图片（嵌入单元格，1000×1000 高清）、尺寸、件数、材质、运单号**。
文件名 = `登录账号名_导出当天YYYYMMDD.xlsx`。

## 构建 / 加载

```bash
npm run dev      # 开发产物 → .output/chrome-mv3-dev（含产物自检）
npm run build    # 生产产物 → .output/chrome-mv3
npm run check    # 只跑产物自检
npm run compile  # vue-tsc 类型检查
```

然后在 `chrome://extensions` 打开开发者模式 → **加载已解压的扩展程序** → 选 `.output/chrome-mv3-dev`。
改了代码后：`npm run dev`，再点扩展卡片上的**刷新**按钮，然后 F5 重开页面。

> ⚠️ **不要跑 `npx wxt`（dev server / serve 模式）。**
> serve 模式的产物 manifest 里**没有 `content_scripts`**（改由它自己的 background 运行时注册），
> 而 `wxt build` 的产物正好相反。两者写进同一个输出目录会杂交出
> 「manifest 说没有内容脚本、background 又不去注册」的产物 ——
> 结果内容脚本永不执行，**页面上面板凭空消失、控制台一行日志都没有，极难排查**。
> `scripts/check-build.mjs` 会拦截这种产物，`wxt.config.ts` 的 `server:started` 钩子也会告警。

## 构建标记（界面不显示）

构建时会自动生成版本标记（形如 `B0925-1610` = 9月25日 16:10），但**不在面板上显示**。
扩展「改了没生效」时用它判断浏览器里是不是旧产物，三种看法：

1. Elements 里看面板宿主元素 `#__dxm_export_panel` 的 `data-build` 属性；
2. 页面控制台里 `[店小秘导出] 账号诊断报告 =` 的 `buildTag` 字段；
3. 构建终端里 `scripts/check-build.mjs` 打印的 `buildTag=…`。

## 面板开关（展开 / 收起）

面板右上角 `✕` 是**收起**（只是隐藏，不删节点）：选的日期范围、识别到的账号、手填的名字都保留，重新展开即原样。
重新展开的入口：

- 工具栏扩展图标 → 弹窗里的开关按钮（会给页面发 `DXM_TOGGLE_PANEL`，
  打开弹窗时先发 `DXM_PANEL_STATE` 问一次当前状态，按钮文案才对得上）；
- 刷新页面 —— 面板默认是展开的（收起状态不跨页面持久化）。

弹窗按 `*://*.dianxiaomi.com/*` 找标签页，不会去找别的站点。

## 账号识别排查（面板上没有入口）

面板不再提供「诊断账号识别」按钮。识别失败、或只拿到低可信的猜测值时，
脚本会把完整报告打进控制台：`[店小秘导出] 账号识别失败，诊断报告 = …`（F12 查看）。

## 商店清单的语言（别把 `__MSG_` 改回字面量）

manifest 里**名称与描述写的是 i18n 占位符**，真实文案在 `public/_locales/zh_CN/messages.json`：

```jsonc
// manifest（由 wxt.config.ts 生成）
"name": "__MSG_extensionName__",
"description": "__MSG_extensionDescription__",
"default_locale": "zh_CN"
```

**为什么不能写回中文字面量**：Microsoft Partner Center 判定商店清单语言的依据，
就是这两个 `__MSG_` 引用 —— 它拿引用的 key 去每个 `_locales/<locale>/messages.json` 里查，
查到哪个语言才为哪个语言建一行清单。写死字面量时它没有任何引用可跟，
**所有语言都被跳过，只剩默认的「英语(美国)」一行** —— 中文插件在商店里就显示成英语，
而且本地加载、构建、自检全都不会报错。`scripts/check-build.mjs` 已加断言卡住这种情况。

两条约束（浏览器强校验，违反会**直接拒绝加载扩展**）：

- `default_locale` 与 `_locales` 目录必须**同时存在**，只有其一不行；
- `__MSG_xxx__` 引用的 key 必须在 `_locales/<default_locale>/messages.json` 里真实存在。

**要加英文清单**：新建 `public/_locales/en/messages.json`，写同名 key（`extensionName` /
`extensionDescription`）的英文文案即可，重新打包后商店会多出一行英语清单，无需改代码。

`action.default_title`（工具栏悬停提示）**不在 `wxt.config.ts` 里设** ——
WXT 会用 `entrypoints/popup/index.html` 的 `<title>` 覆盖它，写在那里是死配置。
它也不参与商店语言判定，所以保持中文字面量即可。

## 目录

```
entrypoints/content.ts            # 核心：面板 UI、vxe-table 解析、账号识别、Excel 导出
entrypoints/background.ts         # 图片下载代理（跨域图片必须在这里抓）
entrypoints/popup/                # 工具栏弹窗（使用说明 + 页面面板开关）
public/_locales/zh_CN/messages.json  # 扩展名称/描述的文案（商店清单语言据此判定）
scripts/check-build.mjs           # 构建产物自检
.dev-verify/                      # Playwright 真机验证脚手架（非产物，可删）
wxt.config.ts                     # 非 ASCII 转义插件、构建标记注入、manifest
```
