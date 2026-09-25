import { defineConfig } from 'wxt';
import type { Plugin } from 'vite';

/**
 * 把产物里所有非 ASCII 字符转义成 \uXXXX（非 BMP 用代理对）。
 *
 * 为什么需要：ExcelJS 内置的 XML 解析器含 U+EFFFF / U+10FFFF 这类 Unicode【非字符】
 * (noncharacter)。Chromium 加载扩展资源时用严格的 IsStringUTF8() 校验，会拒绝非字符，
 * 于是报「该文件采用的不是 UTF-8 编码」，内容脚本直接加载失败。
 *
 * 转义成 ASCII 后文件是纯 ASCII，必然通过校验，运行时语义完全等价（中文照常显示）。
 * 非 BMP 字符用代理对转义（\uDBFF\uDFFF），在字符串、模板串、正则里都合法。
 */
function escapeNonAsciiPlugin(): Plugin {
  return {
    name: 'dxm-escape-non-ascii',
    enforce: 'post',
    generateBundle(_options, bundle) {
      for (const [fileName, out] of Object.entries(bundle)) {
        if (out.type !== 'chunk') continue;
        if (!fileName.endsWith('.js')) continue;
        if (!/[^\x00-\x7F]/.test(out.code)) continue;

        out.code = out.code.replace(/[^\x00-\x7F]/g, (ch) => {
          const cp = ch.codePointAt(0) ?? 0;
          if (cp <= 0xffff) {
            return '\\u' + cp.toString(16).toUpperCase().padStart(4, '0');
          }
          const v = cp - 0x10000;
          const hi = 0xd800 + (v >> 10);
          const lo = 0xdc00 + (v & 0x3ff);
          return (
            '\\u' + hi.toString(16).toUpperCase() + '\\u' + lo.toString(16).toUpperCase()
          );
        });
      }
    },
  };
}

/**
 * 生成构建标记（形如 `B0925-1210` = 9月25日 12:10，本地时间）。
 *
 * 界面上不显示它，只挂在面板宿主元素的 data-build 属性上（并在控制台诊断报告里带上），
 * 用来「确认浏览器里加载的是哪一版」——
 * 排查「改了没生效」时，先看标记对不对，就能区分「代码没生效」和「浏览器里还是旧产物」。
 * 由构建期自动生成，避免手动维护时忘记改、导致误判。
 */
function makeBuildTag(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, '0');
  return `B${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
}

// See https://wxt.dev/api/config.html
export default defineConfig({
  modules: ['@wxt-dev/module-vue'],

  /**
   * serve 模式（`npx wxt` / `wxt dev`）的产物【不能用来加载】。
   *
   * 原因：serve 模式下 WXT 故意不把内容脚本写进 manifest，改由它自己的 background
   * 在运行时用 chrome.scripting.registerContentScripts 动态注册；
   * 而 `wxt build` 的产物正好相反（manifest 里有 content_scripts，background 里没有注册代码）。
   *
   * 两种产物写进同一个输出目录就会互相污染 —— 最典型的后果是
   * 「manifest 说没有内容脚本、background 又不去注册」，于是内容脚本永不执行，
   * 页面上的悬浮面板凭空消失、控制台连一行日志都没有，极难排查。
   *
   * 所以这里直接拦住，统一走 `npm run dev`（= wxt build -m development + 自检）。
   */
  hooks: {
    'server:started'(wxt) {
      wxt.logger.warn(
        '检测到 serve(wxt dev) 模式：该产物会覆盖 .output/chrome-mv3-dev 且缺少 content_scripts，' +
          '加载后页面上面板不会出现。请改用 `npm run dev`（静态构建，带产物自检）。',
      );
    },
  },

  vite: () => ({
    // 构建/版本标记（见 entrypoints/content.ts 的 CONFIG.buildTag，界面不显示）
    define: {
      __DXM_BUILD_TAG__: JSON.stringify(makeBuildTag()),
    },
    // 注意：Vite 8 默认用 oxc 压缩，output.charset 已被移除、esbuild 也未安装，
    // 所以只能自己加 post 钩子转义。
    build: {
      rollupOptions: {
        plugins: [escapeNonAsciiPlugin()],
      },
    },
  }),
  manifest: {
    /**
     * 名称 / 描述 / 工具栏提示写成 i18n 占位符（实际文案在 public/_locales/<locale>/messages.json）。
     *
     * 【为什么必须这么写】Microsoft Partner Center 判定商店清单语言的依据就是 manifest 里
     * name 与 description 的 __MSG_ 引用：它拿引用的 key 去每个 `_locales/<locale>/messages.json`
     * 里查，能查到哪个语言就为哪个语言建一行清单。
     * 写死中文字面量时它没有任何引用可跟 —— 官方文档原话是
     * 「Partner Center uses these message references to identify available languages.
     *   If these message references are missing, the language will be skipped.」
     * 于是所有语言都被跳过，只剩默认的「英语(美国)」一行，
     * 中文插件在商店里就被显示成英语了（这正是我们踩过的坑）。
     */
    name: '__MSG_extensionName__',
    description: '__MSG_extensionDescription__',
    /**
     * 默认语言。必须与 _locales 目录【同时存在】：只有其一浏览器会直接拒绝加载扩展。
     * 设成 zh_CN 后，浏览器界面语言是中文的用户直接看到中文文案；
     * 其他语言回退到它（所以本扩展只有一份中文文案也完全可用）。
     */
    default_locale: 'zh_CN',
    // 内容脚本只在店小秘站点注入（见 entrypoints/content.ts 的 matches）
    // *://*/* 是给 background 下载商品图用的：商品图在 Amazon CDN 等跨域地址，
    // 内容脚本直接 fetch 会被 CORS 拦截，必须带主机权限在 background 里抓。
    host_permissions: ['*://*.dianxiaomi.com/*', '*://*/*'],
    // 只申请真正用到的 API 权限。审核方会拒「申请了却不用」的权限：
    // - storage：记住上次用的用户名（账号自动识别失败时的兜底）
    // - 已删 activeTab / scripting：代码里都没用。内容脚本靠静态 content_scripts 注入，
    //   弹窗给标签页发消息靠上面 host_permissions 的域名权限即可，不需要这两个。
    permissions: ['storage'],
    /**
     * 注意：工具栏图标的悬停提示（action.default_title）不在这里设 ——
     * WXT 会把 `entrypoints/popup/index.html` 的 `<title>` 当作 defaultTitle
     * 并【覆盖】manifest.action.default_title（见 wxt 的 find-entrypoints.mjs + manifest.mjs），
     * 所以写在这里是死配置、改了不生效。要改提示文字就改那个 `<title>`。
     *
     * 它不参与商店语言判定（Partner Center 只读 name / description），
     * 因此保持中文字面量即可，不必做成 i18n 占位符。
     */
  },
});
