/**
 * 店小秘导出扩展 —— 真机端到端验证探针
 *
 * 思路：用 Playwright 加载未打包的 MV3 扩展，把 *.dianxiaomi.com 的请求
 * fulfill 成本地模拟页面（URL 仍是真实域名，内容脚本照常注入），
 * 然后走完整流程：识别账号 → 生成文件名 → 导出 → 捕获下载 → 校验产物。
 *
 * 用法：
 *   node dxm_probe.cjs                              # 场景A：账号区在顶层，异步挂载
 *   MOCK=dxm_mock_iframe.html node dxm_probe.cjs    # 场景B：账号区在同源 iframe 内
 *   NOACCOUNT=1 node dxm_probe.cjs                  # 场景C：完全没有账号区（应中止导出）
 *   HEADFUL=1 node dxm_probe.cjs                    # 有头模式观察
 */
const { chromium } = require('playwright-core');
const fs = require('fs');
const path = require('path');
const os = require('os');

const CHROME = 'C:/Users/13027/AppData/Local/ms-playwright/chromium-1228/chrome-win64/chrome.exe';
const DIST = process.env.DIST || 'D:/soft/JetBrains/workspace/dianxiaomi/.output/chrome-mv3-dev';
const MOCK_FILE = path.join(__dirname, process.env.MOCK || 'dxm_mock.html');
const HEADER_FILE = path.join(__dirname, 'dxm_header.html');
const HEADLESS = process.env.HEADFUL !== '1';
const DL_TIMEOUT = Number(process.env.DL_TIMEOUT || 30000);

(async () => {
  let MOCK = fs.readFileSync(MOCK_FILE, 'utf8');
  const HEADER = fs.readFileSync(HEADER_FILE, 'utf8');
  if (process.env.NOACCOUNT === '1') {
    // 抹掉页面里所有账号区（顶层模板 + iframe 注入脚本）
    MOCK = MOCK.replace(/<template id="hdr-tpl">[\s\S]*?<\/template>/, '')
      .replace(/setTimeout\(function \(\) \{[\s\S]*?\}, 2500\);/, '');
  }
  console.log('扩展目录:', DIST, fs.existsSync(DIST) ? '存在' : '!! 不存在');
  console.log('模拟页面:', MOCK_FILE, MOCK.length, '字节');

  const userDataDir = path.join(os.tmpdir(), 'dxm-probe-' + Date.now());
  const ctx = await chromium.launchPersistentContext(userDataDir, {
    executablePath: CHROME,
    headless: HEADLESS,
    acceptDownloads: true,
    args: [
      `--disable-extensions-except=${DIST}`,
      `--load-extension=${DIST}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-features=DisableLoadExtensionCommandLineSwitch',
    ],
  });

  // 确认扩展真的加载了（MV3 的 background service worker）
  await new Promise((r) => setTimeout(r, 2500));
  const sws = ctx.serviceWorkers();
  console.log('\n=== Service Worker ===');
  console.log(sws.length ? sws.map((s) => s.url()).join('\n') : '(未发现，扩展可能没加载)');

  const page = ctx.pages()[0] || (await ctx.newPage());
  const logs = [];
  page.on('console', (m) => logs.push(`[${m.type()}] ${m.text()}`));
  page.on('pageerror', (e) => logs.push(`[pageerror] ${e.message}`));

  await page.route('**/*', (route) => {
    const url = route.request().url();
    if (url.includes('dianxiaomi.com')) {
      // /frame-header 用于场景 B：账号区所在的同源 iframe
      const body = /\/frame-header/.test(url) ? HEADER : MOCK;
      return route.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body });
    }
    // 放行 Amazon 图片等外部请求，让高清图能真实下载
    return route.continue();
  });

  // 使用真实页面 URL（从 Chrome 历史里读到的发货成功列表地址）
  await page.goto('https://www.dianxiaomi.com/web/order/shipped/success?go=m10403', { waitUntil: 'load' });
  // 账号区模拟为 2.5s 后异步挂载，这里多等一会儿让轮询有机会命中。
  // WAIT=21000 可等到「轮询用尽」的收尾分支（15 × 1.2s ≈ 18s）。
  await page.waitForTimeout(Number(process.env.WAIT || 6500));

  const snap = (stage) =>
    page.evaluate((s) => {
      const host = document.getElementById('__dxm_export_panel');
      const sr = host && host.shadowRoot;
      const t = (id) => {
        const e = sr && sr.getElementById(id);
        return e ? (e.textContent || '').replace(/\s+/g, ' ').trim() : '(missing)';
      };
      const un = sr && sr.getElementById('uname');
      const inp = sr && sr.getElementById('un');
      // 构建标记已不在界面上显示，改为读宿主元素的 data-build
      const buildTag = host ? host.dataset.build || '(no tag)' : '(no host)';
      const box = sr && sr.querySelector('.box');
      return {
        stage: s,
        panelMounted: !!host,
        buildTag,
        // 面板不应再有「诊断账号识别」入口（含其输出区）
        diagEntryGone: sr ? !sr.querySelector('#diag') && !sr.querySelector('#diagout') : false,
        // 面板显示状态 + 圆形开关应已移除（ballGone 必须为 true）
        boxVisible: sr ? sr.querySelector('.box')?.style.display !== 'none' : false,
        ballGone: sr ? !sr.getElementById('ball') : false,
        // 面板上所有可见文案（用来核对界面没有多余文案 / 标记）
        allText: box ? (box.textContent || '').replace(/\s+/g, ' ').trim() : '(no box)',
        labels: sr ? Array.from(sr.querySelectorAll('.lbl')).map((e) => e.textContent.trim()) : [],
        uname: t('uname'),
        unameSource: un ? un.title : '',
        hiddenInput: inp ? inp.value : '(no input)',
        inputVisible: inp ? inp.style.display !== 'none' : false,
        fname: t('fname'),
        tip: t('tip'),
      };
    }, stage);

  const results = [];
  results.push(await snap('① 页面加载后'));

  // 选「近三天」，保证数据落在筛选窗口内
  await page.evaluate(() => {
    const sr = document.getElementById('__dxm_export_panel').shadowRoot;
    sr.querySelectorAll('.opt').forEach((o) => {
      if (o.getAttribute('data-r') === '3d') o.click();
    });
  });
  await page.waitForTimeout(400);
  results.push(await snap('② 选择近三天后'));

  // ③ 圆形开关已移除：✕ 只做「收起」（面板不消失），页面内不再有开关按钮
  results.push({
    stage: '③ ✕ 收起（页面内无圆形开关）',
    ...(await page.evaluate(() => {
      const sr = document.getElementById('__dxm_export_panel').shadowRoot;
      const box = sr.querySelector('.box');
      const state = () => ({
        boxVisible: box.style.display !== 'none',
        ballGone: !sr.getElementById('ball'),
      });
      const initial = state();
      sr.querySelector('.hd .x').click(); // 收起
      const afterX = { ...state(), panelStillMounted: !!sr.querySelector('.box') };
      return { initial, afterX };
    })),
  });

  // ④ 弹窗里的开关：点一次收起、再点一次展开（无圆形按钮后唯一页面内展开入口）
  const extId = sws.length ? new URL(sws[0].url()).host : '';
  if (extId) {
    const panelState = () =>
      page.evaluate(() => {
        const sr = document.getElementById('__dxm_export_panel').shadowRoot;
        return {
          boxVisible: sr.querySelector('.box').style.display !== 'none',
        };
      });

    const popupPage = await ctx.newPage();
    await popupPage.goto(`chrome-extension://${extId}/popup.html`);
    const btn = popupPage.locator('#panel-toggle');
    await btn.waitFor({ state: 'visible', timeout: 8000 });
    await popupPage.waitForTimeout(400); // 等 onMounted 问到状态、按钮文案对上

    const before = { label: (await btn.textContent()).trim(), ...(await panelState()) };
    await btn.click();
    await popupPage.waitForTimeout(800);
    const afterFirst = {
      label: (await btn.textContent()).trim(),
      hint: ((await popupPage.locator('.hint').textContent().catch(() => '')) || '').trim(),
      ...(await panelState()),
    };
    await btn.click();
    await popupPage.waitForTimeout(800);
    const afterSecond = {
      label: (await btn.textContent()).trim(),
      hint: ((await popupPage.locator('.hint').textContent().catch(() => '')) || '').trim(),
      ...(await panelState()),
    };

    // i18n 实测：确认 manifest 里的 __MSG_ 占位符在真实浏览器里能解析出中文。
    // 这是「商店清单为什么显示成英语」那类问题的正面证据 ——
    // 只要 _locales 目录缺失 / default_locale 写错 / key 拼错，这里就会拿到空串或占位符原文。
    const i18n = await popupPage.evaluate(() => ({
      uiLanguage: chrome.i18n.getUILanguage(),
      // 解析后的文案（应为中文）
      extensionName: chrome.i18n.getMessage('extensionName'),
      extensionDescription: chrome.i18n.getMessage('extensionDescription'),
      // getManifest() 拿到的是 Chrome 已解析过的 manifest：key 找得到就是中文文案，
      // 找不到（_locales 缺失 / default_locale 写错 / key 拼错）就会原样露出 __MSG_xxx__
      rawName: chrome.runtime.getManifest().name,
      defaultLocale: chrome.runtime.getManifest().default_locale,
    }));

    results.push({ stage: '④ 弹窗开关（收起/展开）', i18n, before, afterFirst, afterSecond });
    await popupPage.close();
  } else {
    results.push({ stage: '④ 弹窗开关（收起/展开）', skipped: '拿不到扩展 id' });
  }

  // 触发导出并等待下载（先把面板恢复为展开，贴近真实操作）
  const dlPromise = page.waitForEvent('download', { timeout: DL_TIMEOUT });
  await page.evaluate(() => {
    const sr = document.getElementById('__dxm_export_panel').shadowRoot;
    sr.querySelector('.box').style.display = 'block';
    sr.getElementById('exp').click();
  });

  try {
    const dl = await dlPromise;
    const saved = path.join(__dirname, 'dxm_out.xlsx');
    fs.rmSync(saved, { force: true });
    await dl.saveAs(saved);
    results.push({
      stage: '⑤ 下载完成',
      suggestedFilename: dl.suggestedFilename(),
      bytes: fs.statSync(saved).size,
    });
  } catch (e) {
    results.push({
      stage: '⑤ 未触发下载',
      timeoutMs: DL_TIMEOUT,
      note: process.env.NOACCOUNT === '1'
        ? '(场景C预期如此：账号未知时会中止导出并要求手填一次)'
        : '(非预期，说明导出流程被中断)',
      error: String(e).slice(0, 120),
    });
  }

  results.push(await snap('⑥ 导出结束后'));

  const interesting = logs.filter(
    (l) =>
      l.includes('店小秘') ||
      l.includes('pageerror') ||
      l.includes('WARN') ||
      l.toLowerCase().includes('error'),
  );

  console.log('\n===RESULT===');
  console.log(JSON.stringify({ results, logs: interesting.slice(0, 40) }, null, 1));
  console.log('===END===');

  await ctx.close();
})().catch((e) => {
  console.error('探针异常:', e);
  process.exit(1);
});
