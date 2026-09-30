# 店小秘发货导出助手（Chrome MV3 扩展）

在店小秘「发货成功列表」页面注入悬浮面板，把当前列表导出成 Excel。
导出字段：**图片（嵌入单元格，1000×1000 高清）、尺寸(CM)、件数、材质、运单号**。
材质列表页没有源数据，固定填默认值「水洗底」；
尺寸解析后按业务口径 **数值 ×30 换算成厘米、结果取整（不带小数点）**
（只乘能乘的数字，`A4` / `12cm` 这类字母数字粘连的值原样保留，见 `toCmSize`）。
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

## 图片导出（缺图排查）

列表商品图是**懒加载**的：没滚进过视口的行，DOM 里只有 1×1 占位图，真实地址根本不在页面上。
所以点「导出Excel 和面单」时扩展会先把列表**从上到下自动滚一遍**（滚完原样滚回原位，
期间面板显示「正在滚动列表并等待加载…」），然后**重新读一次行**再下载图片 —— 不滚就导出，那些行必然没有图。

导出结束面板会给出统计，缺图分两类：

- `N 行取不到图片地址`：页面本身没给地址（该行确实无图，或列表还没滚到）；
- `N 张图片下载失败`：有地址但下不来，已隔 0.8s 自动重试一轮，具体 URL 见控制台
  `[店小秘导出] N 张图片下载失败`；出现任一情况提示会变红并提示按 F12。

下载链路（`buildImageCandidates` + `entrypoints/background.ts`）按顺序试：
`HD 尺寸修饰符 → 去修饰符原图 → _SL500_ → 换 m.media-amazon.com 域名再走一遍 → 页面原 URL`，
全失败才算失败。图片统一转 1000×1000 白底 JPEG，并按真实 mime 注册 `extension`
（标成 jpeg 却塞 png，Excel 里就是空白格）。

## 面单下载（按钮「导出Excel 和面单」）

点导出按钮后**先**下当页全部订单的面单，再导 Excel —— 顺序不能反：
「开始打印」会 `window.open` 一个 PDF，而 `window.open` 只认**点按钮后约 5s 内**的用户激活，
滚动列表 + 下载几十张商品图会把这个窗口耗光，放后面就会被浏览器拦掉。

链路（`downloadLabels` + `entrypoints/background.ts`）：

1. 勾选当页全部订单（本来就全选则不动，导完**恢复原样的勾选状态**）；
2. 点「**批量打印**」展开菜单 → 点菜单里的「**打印面单**」
   （页面上没批量菜单时，才退回行内那个「打印面单」链接）；
3. 等「**打印管理**」弹窗（`.ant-modal-content`，标题即「打印管理」）：
   - 读出「当前选中 N 条数据」，N = 0 直接报错（说明没勾到）；
   - 把打印方式锁到第一档 **PDF生成打印（仅打印选中 [当前页]）** ——
     只有这档把当页勾选合成**一个 PDF**（第二档要打印驱动、第三档每包一个文件）；
   - 点「**开始打印**」→ 页面 `window.open` 出
     `https://print.dianxiaomi.com/<日期>/<uuid>.pdf`；
4. background 用 `tabs.onCreated / onUpdated` 捕获这个 `.pdf` 地址（不 hook 页面代码，避开 CSP），
   抓完顺手关掉这个一闪而过的标签页；
5. 存成 `账号_日期_面单.pdf`，点弹窗「取消」关掉，还原勾选。

一个 PDF 里含**所有勾选订单** —— 页面自己的批量打印就是合订的，扩展不再二次合并。
面单失败不阻断 Excel，但提示会标红并写明原因；没点到「开始打印」时会主动撤销监听
（`DXM_CANCEL_PDF`），免得用户随后自己开的 PDF 被当成面单关掉。

## 面板开关（展开 / 收起）

面板右上角 `✕` 是**收起**（只是隐藏，不删节点）：识别到的账号、手填的名字都保留，重新展开即原样。
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
