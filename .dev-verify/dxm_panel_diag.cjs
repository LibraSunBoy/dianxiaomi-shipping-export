/**
 * 只做一件事：加载扩展 + 模拟页，把全部 console / pageerror 原样打出来。
 * 用于定位「面板没挂上」这类启动即崩溃的问题。
 */
const { chromium } = require('playwright-core');
const fs = require('fs');
const path = require('path');
const os = require('os');

const CHROME = 'C:/Users/13027/AppData/Local/ms-playwright/chromium-1228/chrome-win64/chrome.exe';
const DIST = process.env.DIST || 'D:/soft/JetBrains/workspace/dianxiaomi/.output/chrome-mv3-dev';
const MOCK = fs.readFileSync(path.join(__dirname, process.env.MOCK || 'dxm_mock.html'), 'utf8');

(async () => {
  const userDataDir = path.join(os.tmpdir(), 'dxm-diag-' + Date.now());
  const ctx = await chromium.launchPersistentContext(userDataDir, {
    executablePath: CHROME,
    headless: true,
    args: [
      `--disable-extensions-except=${DIST}`,
      `--load-extension=${DIST}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--disable-features=DisableLoadExtensionCommandLineSwitch',
    ],
  });

  await new Promise((r) => setTimeout(r, 2000));

  const page = ctx.pages()[0] || (await ctx.newPage());
  const logs = [];
  page.on('console', (m) => logs.push(`[${m.type()}] ${m.text()}`));
  page.on('pageerror', (e) => logs.push(`[PAGEERROR] ${e.message}\n${(e.stack || '').split('\n').slice(0, 6).join('\n')}`));

  await page.route('**/*', (route) => {
    if (route.request().url().includes('dianxiaomi.com')) {
      return route.fulfill({ status: 200, contentType: 'text/html; charset=utf-8', body: MOCK });
    }
    return route.continue();
  });

  await page.goto('https://www.dianxiaomi.com/web/order/shipped/success?go=m10403', { waitUntil: 'load' });
  await page.waitForTimeout(6000);

  const probe = await page.evaluate(() => ({
    panel: !!document.getElementById('__dxm_export_panel'),
    bodyChildren: Array.from(document.body.children).map((e) => e.tagName + (e.id ? '#' + e.id : '') + (e.className ? '.' + String(e.className).split(' ')[0] : '')),
    // 手动执行一次内容脚本的入口，看它会不会抛错
    hasWindowApi: typeof window !== 'undefined',
  }));

  console.log('=== 面板状态 ===');
  console.log(JSON.stringify(probe, null, 1));
  console.log('\n=== 控制台（全部） ===');
  console.log(logs.length ? logs.join('\n') : '(空)');

  await ctx.close();
})().catch((e) => {
  console.error('诊断脚本异常:', e);
  process.exit(1);
});
