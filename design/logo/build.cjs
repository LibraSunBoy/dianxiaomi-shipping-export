/**
 * 出格 (Chu-Ge) · 交付构建器  v1.1
 * 输入：logo-master.svg（唯一几何真源）
 * 输出：variants/*.svg、icon/*.png、public/icon/*.png、preview.html
 * 用法：CHROME=<chromium> NODE_PATH=<node_modules> node build.cjs
 *
 * v1.1 修复：
 *   - 主变体底板颜色 bug（undefined → 黑）
 *   - Lockup 增加深/浅底两套，字标宽度改由浏览器 measureText 实测，杜绝裁字
 *   - 竖排 Lockup 真正竖排（标志在上、字标居中在下）
 */
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright-core');

const DIR = __dirname;
const ROOT = path.resolve(DIR, '..', '..');
const VDIR = path.join(DIR, 'variants');
const IDIR = path.join(DIR, 'icon');
for (const d of [VDIR, IDIR]) fs.mkdirSync(d, { recursive: true });

const master = fs.readFileSync(path.join(DIR, 'logo-master.svg'), 'utf8');
const TILE_G = /<g clip-path="url\(#tile\)">[\s\S]*?<\/g>/;
const MARK_PATHS = `<path d="M512 240 V784"/><path d="M240 304 V784"/><path d="M784 304 V784"/>
    <path d="M240 464 H784"/><path d="M240 624 H784"/>`;

function variant({ tile, mark }) {
  let s = master.replace(/<title>.*?<\/title>/, '');
  s = tile === null ? s.replace(TILE_G, '') : s.replace('fill="#2B6CFF"', `fill="${tile}"`);
  if (mark) s = s.replace('stroke="#FFFFFF"', `stroke="${mark}"`);
  return s;
}
const markPaths = (mark, sw = 128) =>
  `<g fill="none" stroke="${mark}" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round">${MARK_PATHS}</g>`;
const tileBlock = (fill, uid) =>
  `<defs><clipPath id="${uid}"><rect width="1024" height="1024" rx="224"/></clipPath></defs>
   <g clip-path="url(#${uid})"><rect width="1024" height="1024" fill="${fill}"/></g>`;

/** Lockup：标志 + 字标。width 由浏览器实测传入。 */
function lockup({ text, sub = '', fs: fontSize, w, h, bg, fg, subFg = '#8B98AD', fam, ls, pad = 56, M = 200, gap = 36, vertical = false }) {
  const bgRect = bg ? `<rect width="${w}" height="${h}" fill="${bg}"/>` : '';
  const subH = sub ? fontSize * 0.62 : 0;
  if (vertical) {
    const mx = (w - M) / 2, my = pad;
    const ty = my + M + gap + fontSize * 0.86;
    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${bgRect}
  <g transform="translate(${mx},${my}) scale(${M / 1024})">${tileBlock('#2B6CFF', 'lv')}${markPaths('#FFFFFF')}</g>
  <text x="${w / 2}" y="${ty}" text-anchor="middle" font-family="${fam}" font-size="${fontSize}" font-weight="700" letter-spacing="${ls}" fill="${fg}">${text}</text>
  ${sub ? `<text x="${w / 2}" y="${ty + subH + 8}" text-anchor="middle" font-family="'Inter','Segoe UI',sans-serif" font-size="${fontSize * 0.5}" letter-spacing="${fontSize * 0.22}" fill="${subFg}">${sub}</text>` : ''}
</svg>`;
  }
  const my = (h - M) / 2;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}">${bgRect}
  <g transform="translate(${pad},${my}) scale(${M / 1024})">${tileBlock('#2B6CFF', 'lh')}${markPaths('#FFFFFF')}</g>
  <text x="${pad + M + gap}" y="${h / 2 + fontSize * 0.35}" font-family="${fam}" font-size="${fontSize}" font-weight="700" letter-spacing="${ls}" fill="${fg}">${text}</text>
</svg>`;
}

const CN = "'PingFang SC','Microsoft YaHei','Source Han Sans SC','Noto Sans SC',sans-serif";
const EN = "'Inter','Segoe UI',system-ui,sans-serif";

const ICON_SIZES = [16, 32, 48, 96, 128];
const EXTRA_SIZES = [256, 300, 512];
const RASTERS = [16, 20, 24, 32, 48, 64, 96, 128, 256, 300, 512];

(async () => {
  const browser = await chromium.launch({ executablePath: process.env.CHROME, args: ['--hide-scrollbars'] });
  const page = await browser.newPage({ viewport: { width: 900, height: 900 } });
  await page.setContent('<body></body>');

  const png = (svgText, size) =>
    page.evaluate(async ({ s, size }) => {
      const img = new Image();
      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(s);
      await img.decode();
      const c = document.createElement('canvas');
      c.width = size; c.height = size;
      c.getContext('2d').drawImage(img, 0, 0, size, size);
      return c.toDataURL('image/png');
    }, { s: svgText, size }).then((d) => Buffer.from(d.split(',')[1], 'base64'));

  const measure = (text, font, ls) =>
    page.evaluate(({ text, font, ls }) => {
      const c = document.createElement('canvas').getContext('2d');
      c.font = font;
      if ('letterSpacing' in c) c.letterSpacing = ls + 'px';
      return Math.ceil(c.measureText(text).width);
    }, { text, font, ls });

  // ---- 变体 ----
  const VARIANTS = [
    ['01-primary-tile.svg', variant({ tile: '#2B6CFF', mark: '#FFFFFF' }), false],
    ['02-mark-blue.svg', variant({ tile: null, mark: '#2B6CFF' }), false],
    ['03-mark-white.svg', variant({ tile: null, mark: '#FFFFFF' }), true],
    ['04-mark-black.svg', variant({ tile: null, mark: '#111111' }), false],
    ['05-tile-gray.svg', variant({ tile: '#6B7280', mark: '#FFFFFF' }), false],
    ['06-tile-white.svg', variant({ tile: '#FFFFFF', mark: '#2B6CFF' }), false],
  ];
  for (const [n, s] of VARIANTS) fs.writeFileSync(path.join(VDIR, n), s, 'utf8');

  // ---- Lockup（宽度实测）----
  const fsCN = 88, lsCN = 88 * 0.02;
  const nameCN = '店小秘发货导出助手';
  const wCN = await measure(nameCN, `700 ${fsCN}px 'Microsoft YaHei'`, lsCN);
  const nameMark = '出格', fsMark = 150;
  const wMark = await measure(nameMark, `700 ${fsMark}px 'Microsoft YaHei'`, fsMark * 0.02);
  const latin = 'CHU-GE', fsEn = 104, lsEn = fsEn * 0.22;
  const wEn = await measure(latin, `700 ${fsEn}px Inter`, lsEn);

  const PAD = 56, M = 200, GAP = 40;
  const LO = [
    ['07-lockup-vertical-light.svg', lockup({ text: nameCN, fs: fsCN, ls: lsCN, fam: CN, w: Math.max(M, wCN) + 2 * PAD, h: 2 * PAD + M + GAP + fsCN * 0.9, bg: '#FFFFFF', fg: '#12203B', vertical: true, pad: PAD, M, gap: GAP })],
    ['08-lockup-vertical-dark.svg', lockup({ text: nameCN, fs: fsCN, ls: lsCN, fam: CN, w: Math.max(M, wCN) + 2 * PAD, h: 2 * PAD + M + GAP + fsCN * 0.9, bg: '#141C2E', fg: '#FFFFFF', vertical: true, pad: PAD, M, gap: GAP })],
    ['09-lockup-horizontal-light.svg', lockup({ text: nameCN, fs: fsCN, ls: lsCN, fam: CN, w: PAD + M + GAP + wCN + PAD, h: PAD * 2 + M, bg: '#FFFFFF', fg: '#12203B', pad: PAD, M, gap: GAP })],
    ['10-lockup-horizontal-dark.svg', lockup({ text: nameCN, fs: fsCN, ls: lsCN, fam: CN, w: PAD + M + GAP + wCN + PAD, h: PAD * 2 + M, bg: '#141C2E', fg: '#FFFFFF', pad: PAD, M, gap: GAP })],
    ['11-lockup-name-cn.svg', lockup({ text: nameMark, fs: fsMark, ls: fsMark * 0.02, fam: CN, w: PAD + M + GAP + wMark + PAD, h: PAD * 2 + M, bg: '#FFFFFF', fg: '#12203B', pad: PAD, M, gap: GAP })],
    ['12-lockup-latin.svg', lockup({ text: latin, fs: fsEn, ls: lsEn, fam: EN, w: PAD + M + GAP + wEn + PAD, h: PAD * 2 + M, bg: '#FFFFFF', fg: '#12203B', pad: PAD, M, gap: GAP })],
  ];
  for (const [n, s] of LO) fs.writeFileSync(path.join(VDIR, n), s, 'utf8');

  // ---- PNG ----
  const pubDir = path.join(ROOT, 'public', 'icon');
  fs.mkdirSync(pubDir, { recursive: true });
  const iconBuf = {};
  for (const size of RASTERS) iconBuf[size] = await png(master, size);
  for (const size of ICON_SIZES) {
    fs.writeFileSync(path.join(IDIR, `${size}.png`), iconBuf[size]);
    fs.writeFileSync(path.join(pubDir, `${size}.png`), iconBuf[size]);
  }
  for (const size of EXTRA_SIZES) fs.writeFileSync(path.join(IDIR, `${size}.png`), iconBuf[size]);
  const inverseSvg = variant({ tile: null, mark: '#FFFFFF' });
  for (const size of ICON_SIZES) fs.writeFileSync(path.join(IDIR, `inverse-${size}.png`), await png(inverseSvg, size));

  await browser.close();

  const svg = (p) => fs.readFileSync(path.join(DIR, p), 'utf8').replace(/<\?xml[^>]*\?>/, '');
  const b64 = (p) => fs.readFileSync(path.join(DIR, p)).toString('base64');
  const r = (id) => iconBuf[id].toString('base64');

  const grid = (items) => items.map(([cap, p, dark]) =>
    `<div class="card${dark ? ' dark' : ''}">${svg(p)}<div class="cap">${cap}</div></div>`).join('');

  const html = `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>出格 · 店小秘发货导出助手 图标方案</title>
<style>
:root{--ink:#12203B;--mut:#6B7A93;--line:#E3E8F0;--bg:#F5F7FB;--brand:#2B6CFF;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
  font-family:'PingFang SC','Microsoft YaHei','Source Han Sans SC','Noto Sans SC',system-ui,sans-serif;}
.wrap{max-width:1060px;margin:0 auto;padding:44px 26px 72px;}
h1{font-size:29px;margin:0 0 4px;letter-spacing:.5px}
h1 small{display:block;font-size:13.5px;font-weight:400;color:var(--mut);letter-spacing:2px;margin-top:8px}
h2{font-size:16.5px;margin:48px 0 14px;padding-bottom:9px;border-bottom:1px solid var(--line)}
h2 i{font-style:normal;color:var(--brand);margin-right:8px;font-weight:700}
.grid{display:grid;gap:15px} .g3{grid-template-columns:repeat(3,1fr)} .g2{grid-template-columns:repeat(2,1fr)}
.card{background:#fff;border:1px solid var(--line);border-radius:14px;padding:18px;text-align:center;display:flex;
  flex-direction:column;align-items:center;justify-content:center;min-height:190px}
.card.dark{background:#141C2E;border-color:#1E2A42}
.card>svg{width:132px;height:132px;display:block}
.cap{font-size:11.5px;color:var(--mut);margin-top:12px;line-height:1.6}
.card.dark .cap{color:#93A2BC}
.hero{display:grid;grid-template-columns:300px 1fr;gap:30px;align-items:center;background:#fff;
  border:1px solid var(--line);border-radius:18px;padding:30px}
.hero svg{width:250px;height:250px;display:block;margin:0 auto}
.lead{font-size:13.5px;line-height:1.95;color:#33415A}
.kv{font-size:12.5px;color:var(--mut);line-height:1.95;margin-top:14px}
.kv code{background:#EEF2F9;border-radius:5px;padding:1.5px 6px;color:#2B3A55;font-family:'Cascadia Mono',Consolas,monospace}
.sizes{display:flex;align-items:flex-end;gap:22px;flex-wrap:wrap;background:#fff;border:1px solid var(--line);border-radius:14px;padding:22px}
.sizes figure{margin:0;text-align:center}
.sizes img{display:block;background:#fff;margin:0 auto}
.sizes figcaption{font-size:10.5px;color:var(--mut);margin-top:7px}
.mag{image-rendering:pixelated}
table{width:100%;border-collapse:collapse;background:#fff;border:1px solid var(--line);border-radius:12px;overflow:hidden;font-size:13px}
th,td{padding:10px 14px;text-align:left;border-bottom:1px solid var(--line)}
th{background:#F0F4FA;font-weight:600;font-size:12px;color:#445} tr:last-child td{border-bottom:none}
td.ok{color:#1B8A4B;font-weight:600}
.lock{background:#fff;border:1px solid var(--line);border-radius:14px;padding:0;overflow:hidden;display:flex;align-items:center;justify-content:center}
.lock>svg{display:block;max-width:100%;height:auto}
.note{font-size:12.5px;color:var(--mut);line-height:1.9;margin-top:13px}
.mono{font-family:'Cascadia Mono',Consolas,monospace;font-size:12px}
</style></head><body><div class="wrap">

<h1>出格 <small>CHU-GE · 店小秘发货导出助手 · 浏览器扩展图标</small></h1>

<div class="hero">
  <div>${svg('logo-master.svg')}</div>
  <div>
    <div class="lead">
      <b>「出」的几何减笔。</b>三竖栏 = 发货单的「行」，两横档 = 「格」，合起来就是那个「出」字 ——
      <b>导出 / 出货 / 取出去</b>，与这个工具唯一的动作完全同义。中栏上探一格，给出「取出」的方向感，
      而<b>全程不画一根箭头</b>。
    </div>
    <div class="kv">
      比例系统 <code>整数比 · 模数 8px</code>　风格语法 <code>圆端</code>　元素预算 <code>2</code><br>
      视觉外框 <code>672×672</code>　四边留白 <code>176 等重</code>　笔宽 <code>128</code><br>
      主色 <code>#2B6CFF</code>（沿用扩展面板主色）
    </div>
  </div>
</div>

<h2><i>01</i>六变体套件</h2>
<div class="grid g3">${grid([
    ['主变体 · 蓝底白形', 'variants/01-primary-tile.svg', false],
    ['纯图形 · 蓝形透明底', 'variants/02-mark-blue.svg', false],
    ['反白 · 白形透明底（深色界面）', 'variants/03-mark-white.svg', true],
    ['单黑 · 黑形透明底', 'variants/04-mark-black.svg', false],
    ['灰阶 · 灰底白形', 'variants/05-tile-gray.svg', false],
    ['白底蓝形', 'variants/06-tile-white.svg', false],
  ])}</div>

<h2><i>02</i>组合 Lockup（字标宽度实测，不裁字）</h2>
<div class="grid g2">${[
    'variants/07-lockup-vertical-light.svg',
    'variants/08-lockup-vertical-dark.svg',
    'variants/09-lockup-horizontal-light.svg',
    'variants/10-lockup-horizontal-dark.svg',
    'variants/11-lockup-name-cn.svg',
    'variants/12-lockup-latin.svg',
  ].map((p) => `<div class="lock">${svg(p)}</div>`).join('')}</div>
<div class="note">字标字体：中文 <span class="mono">PingFang SC / 思源黑体</span>，拉丁 <span class="mono">Inter</span>；
回退链写在 SVG 的 <span class="mono">font-family</span> 里，缺字体自动降级。</div>

<h2><i>03</i>缩放墙（真实栅格化 · 16→512px 一眼看完）</h2>
<div class="sizes">
${RASTERS.map((s) => `<figure><img src="data:image/png;base64,${r(s)}" width="${s}" height="${s}"><figcaption>${s}px</figcaption></figure>`).join('')}
  <figure><img class="mag" src="data:image/png;base64,${r(16)}" width="128" height="128"><figcaption>16px 放大 8×</figcaption></figure>
</div>

<h2><i>04</i>七项检查</h2>
<table>
<tr><th>检查项</th><th>结果</th><th>说明</th></tr>
<tr><td>剪影</td><td class="ok">通过</td><td>纯黑白轮廓下「出」结构完整，可独立识别</td></tr>
<tr><td>16px</td><td class="ok">通过</td><td>不粘连成团，收拢读作「格」，与 Excel 语义同向（见缩放墙）</td></tr>
<tr><td>单色</td><td class="ok">通过</td><td>变体 04 纯黑一色成立</td></tr>
<tr><td>模糊</td><td class="ok">通过</td><td>重度模糊后仍是「圆角方块 + 内部三栏」的独特质量分布</td></tr>
<tr><td>反白</td><td class="ok">通过</td><td>白形/蓝形 × 深底/浅底 四向皆成立（变体 02 / 03 / 06）</td></tr>
<tr><td>撞脸</td><td class="ok">通过</td><td>见下表；同类工具清一色「表格 + 下载箭头 / 纸箱」，无一以汉字骨架做标</td></tr>
<tr><td>语义安全</td><td class="ok">通过</td><td>见下表</td></tr>
</table>

<h2><i>05</i>撞脸自查</h2>
<table>
<tr><th>参照对象</th><th>其惯用图形语言</th><th>我们的差异</th></tr>
<tr><td>店小秘（平台本体）</td><td>蓝色圆角字标，无独立图形记</td><td>我们是纯几何字形记，不与平台标争同一视觉位</td></tr>
<tr><td>Excel / 表格类图标</td><td>绿底白 X、单元格网格</td><td>不用网格直译，用「出」的骨架承载「格」的含义</td></tr>
<tr><td>常见「导出」扩展</td><td>向下箭头 + 文档 / 云下载</td><td>全程零箭头，靠字义而非符号惯例传达导出</td></tr>
<tr><td>电商 ERP 类扩展</td><td>纸箱 / 货车 / 购物车</td><td>不出现任何具象物流物件</td></tr>
</table>
<div class="note">以上为品类图形语言的惯例比对；<b>未抓取任何真实商标图片</b>，也不作「已查过商标库」的表述 —— 商标检索请自行进行。</div>

<h2><i>06</i>语义安全审查</h2>
<table>
<tr><th>风险项</th><th>结果</th></tr>
<tr><td>180° 旋转</td><td class="ok">无意外形象（呈对称格状）</td></tr>
<tr><td>镜像</td><td class="ok">字形左右本对称，镜像后一致</td></tr>
<tr><td>负空间暗藏物</td><td class="ok">三栏两档围出的六格为纯几何窗格，无暗藏图形</td></tr>
<tr><td>拼音 / 缩写歧义</td><td class="ok">不含拉丁字母</td></tr>
<tr><td>国旗联想</td><td class="ok">非「浅底 + 中央实心圆」结构，无国旗相似性</td></tr>
<tr><td>宗教 / 文化禁忌</td><td class="ok">「出」为中性常用字，无宗教符号形态</td></tr>
</table>

<h2><i>07</i>落地</h2>
<div class="note">
图标 PNG（16/32/48/96/128）已直接写入扩展的 <span class="mono">public/icon/</span>，重新构建即生效；
图标目录另存 256/300/512 与反白套件，商店封面用 Lockup。几何真源只有一个：
<span class="mono">design/logo/logo-master.svg</span>，改它重跑 <span class="mono">build.cjs</span> 即可全量再生成。
</div>

</div></body></html>`;

  fs.writeFileSync(path.join(DIR, 'preview.html'), html, 'utf8');
  console.log('variants:', VARIANTS.length + LO.length, '| icons:', ICON_SIZES.join('/'), '+256/300/512 +inverse*5');
  console.log('lockup 实测宽度: CN=', wCN, ' 出格=', wMark, ' CHU-GE=', wEn);
  console.log('→ public/icon/ 已更新，preview.html 已生成');
})().catch((e) => { console.error(e); process.exit(1); });
