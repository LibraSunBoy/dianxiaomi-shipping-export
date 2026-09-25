#!/usr/bin/env node
/**
 * 构建产物自检 —— 每次构建后自动跑，防止「静默损坏」的产物被加载进浏览器。
 *
 * 背景（真实事故）：
 *   本项目同时存在两种 wxt 产物，它们的 manifest 与 background 是【互相配套】的：
 *
 *     wxt dev (serve)  → manifest 里【没有】content_scripts，
 *                        由它自己的 background 在运行时用
 *                        chrome.scripting.registerContentScripts 动态注册；
 *     wxt build (build) → manifest 里【有】content_scripts，
 *                        但 background 里【没有】动态注册代码。
 *
 *   两者写进同一个目录（.output/chrome-mv3-dev）时就会互相污染：
 *   manifest 说「没有内容脚本」，background 又不会去注册 ——
 *   结果内容脚本永远不执行，页面上的悬浮面板凭空消失，且控制台一行日志都没有。
 *   这个错误极其隐蔽（扩展能加载、能启用、background 也活着），所以必须用脚本卡住。
 *
 * 用法：node scripts/check-build.mjs [outDir...]
 *       不带参数时检查 .output/chrome-mv3-dev 与 .output/chrome-mv3（存在才查）。
 */
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';

const ROOT = process.cwd();
const DEFAULT_DIRS = ['.output/chrome-mv3-dev', '.output/chrome-mv3'];
const dirs = process.argv.slice(2).length ? process.argv.slice(2) : DEFAULT_DIRS;

/** 内容脚本必须匹配的域名（与 entrypoints/content.ts 的 matches 对应） */
const REQUIRED_MATCH = '.dianxiaomi.com';
/** 面板宿主元素 id（存在于产物说明 content.ts 真的被编进去了） */
const PANEL_ID = '__dxm_export_panel';
/** 工具栏弹窗的「页面面板开关」用到的消息名（content.ts 监听、App.vue 发送） */
const PANEL_MSG = 'DXM_TOGGLE_PANEL';

let failed = 0;
const problems = [];
const notes = [];

function fail(msg) {
  failed++;
  problems.push(msg);
}

for (const rel of dirs) {
  const dir = path.resolve(ROOT, rel);
  const tag = `[${rel}]`;

  if (!fs.existsSync(dir)) {
    notes.push(`${tag} 不存在，跳过`);
    continue;
  }

  // ---- 1) manifest 必须存在、可解析 ----
  const manifestPath = path.join(dir, 'manifest.json');
  if (!fs.existsSync(manifestPath)) {
    fail(`${tag} 缺少 manifest.json`);
    continue;
  }
  let manifest;
  try {
    manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  } catch (e) {
    fail(`${tag} manifest.json 不是合法 JSON：${e.message}`);
    continue;
  }

  // ---- 2) 关键：manifest 必须声明 content_scripts ----
  // 缺了它 = 内容脚本不会被注入（serve 模式产物的典型特征）
  const cs = manifest.content_scripts;
  if (!Array.isArray(cs) || cs.length === 0) {
    fail(
      `${tag} manifest.content_scripts 缺失！内容脚本不会被注入，页面上面板不会出现。` +
        `\n          常见原因：目录里的产物是 "wxt"(serve) 与 "wxt build" 混出来的。` +
        `\n          处理：删掉该目录后重新执行 npm run dev。`,
    );
  } else {
    const hit = cs.find((c) => (c.matches ?? []).some((m) => m.includes(REQUIRED_MATCH)));
    if (!hit) {
      fail(`${tag} content_scripts 存在，但没有匹配 ${REQUIRED_MATCH} 的条目：${JSON.stringify(cs)}`);
    } else {
      // content_scripts 里引用的 js 文件必须真实存在
      for (const js of hit.js ?? []) {
        if (!fs.existsSync(path.join(dir, js))) {
          fail(`${tag} content_scripts 引用的文件不存在：${js}`);
        }
      }
      if (!(hit.js ?? []).length) fail(`${tag} content_scripts 条目的 js 列表为空`);
    }

    // serve 模式的产物会带 host_permissions 但无 content_scripts；
    // 反向识别：有 content_scripts 却又带 dev server 标志 → 也是混出来的。
    if (manifest.commands?.['wxt:reload-extension']) {
      fail(
        `${tag} 同时存在 content_scripts 和 wxt:reload-extension 命令 —— ` +
          `这是 serve / build 产物混用的确凿特征，请删目录重建。`,
      );
    }
  }

  // ---- 3) background service worker ----
  const sw = manifest.background?.service_worker;
  if (!sw) {
    fail(`${tag} manifest 缺少 background.service_worker`);
  } else if (!fs.existsSync(path.join(dir, sw))) {
    fail(`${tag} background.service_worker 指向的文件不存在：${sw}`);
  }

  // ---- 4) 内容脚本：可解析 + 纯 ASCII + 含面板代码 ----
  for (const js of cs?.[0]?.js ?? []) {
    const p = path.join(dir, js);
    if (!fs.existsSync(p)) continue;

    // 4a) 必须能被解析（防止被中断的构建留下半截文件）。
    // 用 vm.Script 在进程内编译，不 spawn 子进程 —— 后者在 Windows 上会偶发
    // EBUSY（node.exe 被占用），把好产物误判成语法错误。
    const src = fs.readFileSync(p, 'utf8');
    try {
      new vm.Script(src, { filename: p });
    } catch (e) {
      fail(`${tag} ${js} 语法错误（构建可能被中断，文件不完整）：${String(e.message).slice(0, 200)}`);
      continue;
    }

    const buf = Buffer.from(src, 'utf8');
    // 4b) 必须是纯 ASCII —— Chromium 对扩展资源做严格 UTF-8 校验，
    //     ExcelJS 里 U+EFFFF / U+10FFFF 这类 Unicode 非字符会导致
    //     「该文件采用的不是 UTF-8 编码」而拒绝加载。
    let nonAscii = 0;
    for (const b of buf) if (b > 127) nonAscii++;
    if (nonAscii > 0) {
      fail(
        `${tag} ${js} 含 ${nonAscii} 个非 ASCII 字节，Chromium 可能拒绝加载。` +
          `\n          检查 wxt.config.ts 的 escapeNonAsciiPlugin 是否生效。`,
      );
    }

    // 4c) 面板代码必须在里面
    if (!buf.includes(PANEL_ID)) {
      fail(`${tag} ${js} 里没有面板宿主 id "${PANEL_ID}" —— content.ts 可能没被编进去`);
    }

    // 4d) 面板开关的消息监听也必须在（收起后靠工具栏弹窗的按钮重新展开）
    if (!buf.includes(PANEL_MSG)) {
      fail(
        `${tag} ${js} 里没有 "${PANEL_MSG}" —— ` +
          `面板收起后弹窗里的开关将无法展开它（检查 content.ts 的 runtime.onMessage）`,
      );
    }
  }

  // ---- 5) i18n：商店清单能不能被识别成中文，全靠这一段 ----
  // 背景：Microsoft Partner Center 判定商店清单语言的依据，是 manifest 的 name / description
  // 里的 __MSG_ 引用 —— 它拿 key 去 _locales/<locale>/messages.json 里查，查到才为该语言建清单。
  // 一旦这里写回中文字面量，Partner Center 会把所有语言跳过、只留默认的「英语(美国)」，
  // 中文插件在商店里就显示成英语，而且【本地与构建都毫无报错】，极易被忽略。故用脚本卡住。
  const dloc = manifest.default_locale;
  const localesDir = path.join(dir, '_locales');
  if (dloc) {
    if (!fs.existsSync(localesDir)) {
      fail(
        `${tag} manifest 声明了 default_locale="${dloc}"，但产物里没有 _locales 目录。` +
          `\n          两者必须同时存在，只有其一浏览器会直接拒绝加载该扩展。`,
      );
    } else {
      const msgPath = path.join(localesDir, dloc, 'messages.json');
      if (!fs.existsSync(msgPath)) {
        fail(`${tag} 缺少 _locales/${dloc}/messages.json —— 默认语言的文案无处可取，浏览器会拒绝加载`);
      } else {
        let msgs = null;
        try {
          msgs = JSON.parse(fs.readFileSync(msgPath, 'utf8'));
        } catch (e) {
          fail(`${tag} _locales/${dloc}/messages.json 不是合法 JSON：${e.message}`);
        }
        if (msgs) {
          // name / description 必须是 __MSG_ 引用，且引用的 key 在 messages.json 里真实存在。
          // 缺任意一条 → Partner Center 跳过该语言 → 商店清单退化成「英语(美国)」。
          for (const field of ['name', 'description']) {
            const v = manifest[field];
            const m = typeof v === 'string' ? /^__MSG_(.+)__$/.exec(v) : null;
            if (!m) {
              fail(
                `${tag} manifest.${field} 是字面量「${v}」而不是 __MSG_ 引用。` +
                  `\n          Partner Center 将无法识别任何语言，商店清单会退化成「英语(美国)」。` +
                  `\n          修法：manifest 里写 __MSG_xxx__，文案放进 _locales/<locale>/messages.json。`,
              );
            } else if (!msgs[m[1]]) {
              fail(
                `${tag} manifest.${field} 引用了 __MSG_${m[1]}__，` +
                  `但 _locales/${dloc}/messages.json 里没有 key「${m[1]}」`,
              );
            }
          }
        }
      }
    }
  } else if (fs.existsSync(localesDir)) {
    fail(
      `${tag} 有 _locales 目录但 manifest 没有 default_locale —— 浏览器会直接拒绝加载该扩展。`,
    );
  }

  const firstJs = cs?.[0]?.js?.[0];
  notes.push(
    `${tag} OK — content_scripts=${cs?.length ?? 0} 条，` +
      `background=${sw ?? '缺失'}，` +
      `buildTag=${firstJs ? describeBuildTag(path.join(dir, firstJs)) : '(无内容脚本)'}`,
  );
}

/** 从产物里抠出构建标记（CONFIG.buildTag），用于确认浏览器里跑的是哪一版 */
function describeBuildTag(p) {
  if (!p) return '(无内容脚本)';
  try {
    const s = fs.readFileSync(p, 'utf8');
    const m = s.match(/buildTag:\s*["'`]([^"'`]+)["'`]/);
    return m ? m[1] : '(未找到)';
  } catch {
    return '(读取失败)';
  }
}

console.log('\n===== 构建产物自检 =====');
for (const n of notes) console.log('  · ' + n);
if (problems.length) {
  console.log('\n  ✗ 发现 ' + problems.length + ' 个问题：');
  for (const p of problems) console.log('    - ' + p);
  console.log('\n  自检未通过：这份产物加载进浏览器后很可能没有反应，先修好再加载。\n');
  process.exit(1);
}
console.log('\n  ✓ 自检通过：产物结构与编码均正常，可以直接加载。\n');
