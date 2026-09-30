/**
 * 店小秘发货成功列表导出插件 —— Content Script
 *
 * 功能：在 https://www.dianxiaomi.com/ 的「发货成功列表」页面注入一个悬浮面板，
 *      把当前列表数据按固定字段（图片、尺寸(CM)、件数、材质、运单号）导出为 Excel，
 *      文件名取当前登录账号用户名。
 *
 * 数据来源（店小秘 vxe-table 真实结构）：
 *   - 图片 / 尺寸 / 件数  ← 「商品信息」列（图片取 <img>，尺寸取 size，件数取 x N）
 *   - 运单号              ← 「物流方式」列里的单号（如 「YT2626700708781963」）
 *   - 材质                ← 店小秘发货成功列表页【无此字段】，统一填默认值「水洗底」
 *   - 尺寸                ← 解析后按业务口径 ×30 换算成厘米（见 toCmSize）
 *
 * 说明：本页面表格是 vxe-table，表头比数据行多 1 列（首列选择框），且存在固定列，
 *       因此不能简单按「表头索引 == 数据行索引」读取。这里用每个单元格自带的
 *       colid 属性做表头↔数据行关联，彻底规避列偏移问题。
 */

// 用 ExcelJS 而不是 SheetJS：SheetJS 社区版不支持往单元格里嵌图片，
// ExcelJS 支持 addImage，能把商品图真正嵌入 Excel（图片本体在文件里，不是链接）。
import ExcelJS from 'exceljs';

/**
 * 构建标记，由 wxt.config.ts 的 `define` 在构建时替换成字符串字面量（如 "B0925-1210"）。
 * 这里只做类型声明，值来自构建期注入。
 */
declare const __DXM_BUILD_TAG__: string;

// ============================ 可配置项 ============================

const CONFIG = {
  /** 导出目标列（顺序即 Excel 列顺序），与参考表 张攀7-10发货表.xlsx 一致 */
  targetColumns: ['图片', '尺寸(CM)', '件数', '材质', '运单号'] as const,

  /** 材质列默认值 —— 店小秘发货成功列表页没有这个字段，按业务口径固定填这个 */
  defaultMaterial: '水洗底',

  /**
   * 店小秘发货成功列表真实列标题 -> 关键词（命中其一即用该列）。
   * 商品信息 / 物流方式 是解析目标字段的数据来源；
   * 时间 不再参与导出，只用来把数据行与 vxe 的展开详情行区分开（见 readRows）。
   */
  sourceColumns: {
    商品信息: ['商品信息'],
    物流方式: ['物流方式'],
    时间: ['时间'],
  },

  /**
   * 探测当前登录用户名的选择器（按顺序尝试，取第一个非空文本）。
   *
   * 用【整页 HTML】核实过真实结构（4.1MB 的 body 快照里 `user-name` 只出现 1 次）：
   *   body.order-layout-active
   *     └ #app
   *       └ .in-body.order-layout
   *         └ .nav-index
   *           └ .fr.head-nav-right
   *             └ .li-hover.user.relative   ← 13 个 li-hover 里唯一带 `user` 的
   *               └ .btn-user
   *                 └ .user-name.fl.p-left10[title="jia_yangdong"]
   *
   * 注意两点：
   *   1) 账号栏在主文档里（不在 iframe），且位于 `#app` 内部 —— 见 candidateDocs 的兼容处理；
   *   2) 账号名同时写在 `title` 属性里，优先读 title（不受角标/子元素拼接干扰）。
   */
  usernameSelectors: [
    // —— 真实结构（最精确 → 最宽松）——
    '.user-name',
    '.btn-user .user-name',
    '.li-hover.user .user-name',
    '.head-nav-right .user-name',
    '.nav-index .user-name',
    // 只认结构、不认类名的兜底：容器内任意带 title 的节点（账号名一定写在 title 里）
    '.btn-user [title]',
    '.li-hover.user [title]',
    '.head-nav-right [title]',
    // —— 历代版本 / 其它可能的写法 ——
    '.user-info .name',
    '.user-name-text',
    '#userName',
    '.account-name',
    '.header-user .name',
    '.top-user .name',
    '.login-name',
    '.shop-name',
    '[class*="userName"]',
    '[class*="user-name"]',
    '[class*="account"] [class*="name"]',
  ],

  /**
   * 顶部区域里一定不是用户名的词（导航/功能文案）。
   * 启发式扫描时用来排除干扰项。
   *
   * 注意：这里是**整串精确匹配**。下面的词来自真实页面 `title` 属性的全量枚举
   * （整页只有 17 个不同的 title，除了账号名全是这些导航词）。
   */
  usernameStopWords: [
    '首页', '订单', '商品', '产品', '采购', '库存', '仓库', '物流', '发货', '客服',
    '财务', '数据', '报表', '设置', '系统', '工具', '帮助', '退出', '登录', '注册',
    '注销', '消息', '通知', '搜索', '全部', '更多', '刷新', '返回', '工作台', '店小秘',
    '上一页', '下一页', '末页', '收起', '展开', '排序方式', '发货时间', '物流方式',
    '平台渠道', '店铺账号', '国家区域', '帮助文档', '智能客服', '回到顶部', '打单类型',
    '点击查看物流追踪', '加载中', '暂无数据', '托管', '广告', '服务', '财务',
    'home', 'logout', 'login', 'sign', 'search', 'setting', 'help', 'menu',
  ],

  panelTitle: '店小秘发货导出助手',
  /**
   * 构建标记：**界面不显示**，只挂在面板宿主元素的 data-build 属性上。
   * 用途是「确认浏览器里加载的是哪一版」——扩展改了没生效时，
   * 在 Elements 里看 #__dxm_export_panel 的 data-build，或看控制台的诊断报告。
   *
   * 值由 wxt.config.ts 的 define 在【每次构建时自动生成】（月日-时分），
   * 这样不会出现「忘了手动改标记、误判成旧产物」的情况。
   */
  buildTag: __DXM_BUILD_TAG__,

  /** 上一次用过的用户名，存起来做兜底（页面探测失败时不用每次手填） */
  storageKey: 'dxm_last_username',

  // ---- 图片嵌入相关 ----
  /**
   * 高清化目标尺寸。
   *
   * 列表里 <img> 的 src 是 Amazon 缩略图（`..._SL100.jpg` = 高度 100px），
   * 直接抓下来嵌进 Excel 就是糊的。Amazon 的 URL 支持 `_SL<n>_` 尺寸修饰符，
   * 换成 `_SL1000_` 即可拿到 1000px 的大图（比鼠标悬浮放大那张还大）。
   * 这里优先请求 HD_SIZE，失败再回退原图 / 中等图 / 页面原 URL（见 buildImageCandidates）。
   */
  hdSize: 1000,
  /** 图片统一转成 JPEG（白底），并等比缩到最长边不超过这个像素（控制 xlsx 体积） */
  imgMaxPx: 1000,
  /** JPEG 质量 */
  imgQuality: 0.92,
  /** Excel 里图片显示的边长（像素） */
  imgPx: 110,
  /** 「图片」列列宽（Excel 字符单位） */
  imgColWidth: 18,
  /** 数据行行高（点，约 1px=0.75pt；110px 图约需 82.5pt） */
  imgRowHeight: 84,
};

type TargetColumn = (typeof CONFIG.targetColumns)[number];

// ============================ 工具函数 ============================

/** 取单元格的 colid（vxe-table 用它与表头关联；没有则退化到 class 里的 col_N） */
function colidOf(cell: Element | null | undefined): string | null {
  if (!cell) return null;
  const a = cell.getAttribute('colid');
  if (a) return a;
  const m = (cell.className || '').match(/col_(\d+)/);
  return m ? 'col_' + m[1] : null;
}

/** 判断文本是否命中任一关键词（大小写/空白不敏感） */
function matchAny(text: string, keywords: string[]): boolean {
  const t = (text ?? '').trim().toLowerCase();
  if (!t) return false;
  return keywords.some((k) => t.includes(k.trim().toLowerCase()));
}

/** 把相对地址转绝对地址 */
function absUrl(src: string): string {
  try {
    return new URL(src, location.href).href;
  } catch {
    return src;
  }
}

/** 等待 ms 毫秒（滚动触发懒加载、图片收尾时用） */
function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * `data:` 地址是不是「还没加载出来的占位图」。
 * 占位图就是 1×1 透明 GIF，百来个字节；真正内嵌在页面里的图片（base64）远比这长。
 */
function isInlinePlaceholder(src: string): boolean {
  return src.startsWith('data:') && src.length < 1024;
}

/**
 * 从一个 <img> 取真实图片地址；取不到返回 ''（= 这张图还没加载出来）。
 *
 * 图片是懒加载的：没进过视口的行，`src` 只是 1×1 占位 dataURL（data:image/gif...），
 * 真实地址此时不在 DOM 里 —— 直接拿它导出就是空图。所以：
 *   1) src 已是真实地址 → 用它；
 *   2) src 是占位 → 按常见懒加载属性（data-src / data-original / ...）兜底；
 *   3) 都没有 → 返回 ''，由调用方决定是「滚一遍列表再取」还是留空。
 */
function imgUrl(im: Element | null | undefined): string {
  if (!im) return '';
  for (const attr of ['src', 'data-src', 'data-original', 'data-lazy', 'data-url', 'data-img']) {
    const v = im.getAttribute(attr);
    if (!v || v.startsWith('blob:')) continue;
    if (isInlinePlaceholder(v)) continue; // 1×1 占位 GIF
    return absUrl(v);
  }
  return '';
}

/**
 * 图片是否仍是「未加载」状态（值得再滚一遍 / 等一下才能取到地址）：
 *   - 没 src、blob:、1×1 占位 GIF → 未加载；
 *   - 真实地址但还没下载完（`!complete`）→ 也算，读行前等它下载完能少留空格；
 *   - 解码失败（naturalWidth 为 0）和 1×1 → 未加载。
 */
function isPendingImg(im: Element): boolean {
  const src = im.getAttribute('src') || '';
  if (!src || src.startsWith('blob:') || isInlinePlaceholder(src)) return true;
  const el = im as HTMLImageElement;
  if (!el.complete) return true;
  return el.naturalWidth <= 1;
}

/**
 * 由页面上的缩略图 URL 推出「高清图」候选链（按优先级）。
 *
 * Amazon 图床的 URL 形如：
 *   https://ecx.images-amazon.com/images/I/81q+Af25mUL._SL100.jpg
 *                                          └──────┬─────┘
 *                                          images/I/<图片ID>._<尺寸修饰符>.<扩展名>
 * 其中 `_SL100` 表示「高度压到 100px」，这正是导出的图糊成一片的原因。
 * 把修饰符整体换成 `_SL1000_` 就能拿到 1000px 版本（比鼠标悬浮放大那张还大）。
 *
 * 候选链顺序：
 *   1) 换 HD 尺寸修饰符（首选，清晰且体积可控）
 *   2) 去掉修饰符 → 原图（最清晰，但可能几 MB）
 *   3) 换 `_SL500_`（保险的中等尺寸）
 *   4) 以上三条再换到 `m.media-amazon.com` 走一遍 —— 老图床 `ecx.` 对部分图片
 *      /部分尺寸修饰符会 404，同一张图换个域名往往就活了
 *   5) 页面原始 URL（最终兜底，即使前面都 404 也不至于整格空白）
 *
 * 非 Amazon 图床（URL 里没有尺寸修饰符）不会被改动，只返回原 URL。
 */
function buildImageCandidates(url: string): string[] {
  const out: string[] = [];
  const m = url.match(/^(.+?)\._[^./]+_?\.(jpg|jpeg|png|webp|gif)(\?.*)?$/i);
  if (m) {
    const base = m[1]!;
    const ext = m[2]!;
    const qs = m[3] ?? '';
    const bases = [base];
    const alt = base.replace(
      /^https?:\/\/(?:ecx|images|media)\.images-amazon\.com/i,
      'https://m.media-amazon.com',
    );
    if (alt !== base) bases.push(alt);
    for (const b of bases) {
      out.push(`${b}._SL${CONFIG.hdSize}_.${ext}${qs}`);
      out.push(`${b}.${ext}${qs}`);
      out.push(`${b}._SL500_.${ext}${qs}`);
    }
  }
  out.push(url);
  return Array.from(new Set(out));
}

/** 补零两位数 */
function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

/** YYYYMMDD */
function ymd(d: Date): string {
  return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}`;
}

/** 本地日期（当天 00:00:00） */
function startOfDay(d = new Date()): Date {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

// ============================ 表格探测与读取 ============================

interface ScanResult {
  /** 表头 colid -> 标题 */
  titleByColid: Map<string, string>;
  /** 表头索引 -> 标题（colid 缺失时兜底） */
  titleByIndex: string[];
  /** 表头原始标题（调试用） */
  headers: string[];
}

/**
 * 定位 vxe 表格。
 *
 * 关键：店小秘把「表头」和「数据行」拆成了【两个独立的 <table>】——
 *   - table.vxe-table--header：只有 <thead>，一行数据都没有
 *   - table.vxe-table--body  ：只有 <tbody>，没有表头
 * 所以绝不能只挑一个 table 就读，必须分别取，再靠 colid 关联。
 */
function findTables(): { headerTable: HTMLTableElement | null; bodyTable: HTMLTableElement | null } {
  const tables = Array.from(document.querySelectorAll('table')) as HTMLTableElement[];
  const headerTable =
    tables.find((t) => t.querySelector('.vxe-header--column')) ??
    tables.find((t) => t.querySelector('thead')) ??
    null;
  const bodyTable =
    tables.find((t) => t.querySelector('.vxe-body--row')) ??
    tables.find((t) => t.querySelector('tbody tr')) ??
    null;

  // 单表场景（表头与数据在同一个 table）：两者都指向它
  if (headerTable && !bodyTable && headerTable.querySelector('tbody tr')) {
    return { headerTable, bodyTable: headerTable };
  }
  if (!headerTable && bodyTable && bodyTable.querySelector('thead')) {
    return { headerTable: bodyTable, bodyTable };
  }
  return { headerTable, bodyTable };
}

/** 读取表头，建立 colid->标题 映射（用于和数据行单元格关联） */
function buildHeaderMap(tbl: HTMLTableElement): ScanResult {
  const headerRow = tbl.querySelector('thead tr') ?? tbl.querySelector('tr');
  const headerCells = headerRow ? Array.from(headerRow.querySelectorAll('th, td')) : [];
  const titleByColid = new Map<string, string>();
  const titleByIndex: string[] = [];
  for (const c of headerCells) {
    const k = colidOf(c);
    const t = (c.textContent ?? '').trim();
    if (k) titleByColid.set(k, t);
    titleByIndex.push(t);
  }
  return { titleByColid, titleByIndex, headers: titleByIndex };
}

/** 根据关键词在数据行里找对应单元格（按 colid 关联的标题匹配） */
function pickCell(rowMap: Map<string, Element>, keywords: string[]): Element | undefined {
  for (const [title, el] of rowMap) {
    if (matchAny(title, keywords)) return el;
  }
  return undefined;
}

/** 读取数据行：每行返回 { 列标题 -> 单元格元素 }，并用 colid 关联表头 */
function readRows(tbl: HTMLTableElement, map: ScanResult): Array<Map<string, Element>> {
  const bodyRows = Array.from(tbl.querySelectorAll('tbody tr')).filter(
    (tr) => tr.className.includes('vxe-body--row') && !tr.className.includes('expanded'),
  );
  const rows: Array<Map<string, Element>> = [];
  for (const tr of bodyRows) {
    const cells = Array.from(tr.querySelectorAll('td'));
    if (cells.length === 0) continue;
    const rowMap = new Map<string, Element>();
    cells.forEach((cell, i) => {
      const k = colidOf(cell);
      const title = (k && map.titleByColid.get(k)) || map.titleByIndex[i] || '';
      if (title) rowMap.set(title, cell);
    });
    // 跳过 vxe 的展开详情行（没有「物流方式/时间」列，只是合并单元格的明细）
    const isDetail =
      !pickCell(rowMap, CONFIG.sourceColumns['物流方式']) && !pickCell(rowMap, CONFIG.sourceColumns['时间']);
    if (!isDetail && rowMap.size > 0) rows.push(rowMap);
  }
  // 兜底：未识别到 vxe 行时，退化为直接读所有 tr 的数据行
  if (rows.length === 0) {
    const allTr = Array.from(tbl.querySelectorAll('tr'));
    const startFrom = tbl.querySelector('thead') ? 1 : 0;
    for (let i = startFrom; i < allTr.length; i++) {
      const cells = Array.from(allTr[i]!.querySelectorAll('th, td'));
      if (cells.length === 0) continue;
      const rowMap = new Map<string, Element>();
      cells.forEach((cell, idx) => {
        const t = map.titleByIndex[idx] || '';
        if (t) rowMap.set(t, cell);
      });
      if (rowMap.size > 0) rows.push(rowMap);
    }
  }
  return rows;
}

// ============================ 懒加载图片：计数与触发 ============================

/** 数据表里还有多少张商品图没加载出来（占位 dataURL / 空 src） */
function countPendingImages(tbl: HTMLTableElement | null): number {
  if (!tbl) return 0;
  let n = 0;
  for (const im of tbl.querySelectorAll('tbody img')) if (isPendingImg(im)) n++;
  return n;
}

/** 滚动器抽象 —— 元素容器与整页滚动用同一套逻辑 */
interface Scroller {
  get: () => number;
  set: (v: number) => void;
  /** 可滚动总高度 */
  extent: () => number;
  /** 视口高度 */
  viewport: () => number;
}

function elScroller(el: HTMLElement): Scroller {
  return {
    get: () => el.scrollTop,
    set: (v) => {
      el.scrollTop = v;
    },
    extent: () => el.scrollHeight,
    viewport: () => el.clientHeight || window.innerHeight,
  };
}

/**
 * 找到能带动列表滚动的东西，按优先级：
 *   1) vxe 自己的 `.vxe-table--body-wrapper`（列表滚容器，首选）
 *   2) 向上第一个真正 overflow 滚动的祖先
 *   3) 整页滚动（列表一屏放得下、但图片还在页面折叠区的情况）
 */
function findScroller(tbl: HTMLTableElement): Scroller | null {
  const vxe = tbl.closest('.vxe-table--body-wrapper') as HTMLElement | null;
  if (vxe && vxe.scrollHeight > vxe.clientHeight + 4) return elScroller(vxe);

  let cur = tbl.parentElement;
  while (cur && cur !== document.documentElement) {
    if (cur.scrollHeight > cur.clientHeight + 4) {
      const oy = getComputedStyle(cur).overflowY;
      if (oy === 'auto' || oy === 'scroll' || oy === 'overlay') return elScroller(cur as HTMLElement);
    }
    cur = cur.parentElement;
  }

  const de = document.scrollingElement;
  if (de && de.scrollHeight > window.innerHeight + 4) {
    return {
      get: () => de.scrollTop,
      set: (v) => de.scrollTo(0, v),
      extent: () => de.scrollHeight,
      viewport: () => window.innerHeight,
    };
  }
  return null;
}

/**
 * 把列表从上到下滚一遍，触发懒加载把商品图真正加载出来，最后滚回原位。
 * 返回仍处于未加载状态的图片数（0 = 全部就绪）。
 *
 * 为什么必须滚动：没进过视口的行，<img> 的 src 只是 1×1 占位图，
 * 真实地址根本不在 DOM 里 —— 不滚就导出，那些行必然没有图片
 * （这正是「有时候图片下不进 Excel」最常见的原因）。
 */
async function ensureImagesLoaded(tbl: HTMLTableElement | null): Promise<number> {
  if (!tbl) return 0;
  if (countPendingImages(tbl) === 0) return 0;
  const scroller = findScroller(tbl);
  if (!scroller) return countPendingImages(tbl);

  const origin = scroller.get();
  const step = Math.max(240, Math.floor(scroller.viewport() * 0.8));
  let guard = 0;
  try {
    for (let y = 0; y < scroller.extent() && guard < 100; y += step, guard++) {
      scroller.set(y);
      await sleep(150);
      if (countPendingImages(tbl) === 0) break;
    }
    // 收尾：等最后一屏的图走完网络（懒加载触发 ≠ 已经下载完）。
    // 连续 3 次数量不降就认定「剩下的本来就不会加载」（页面自身加载失败的图），
    // 不再干等 —— 免得面板卡在提示上好几秒。
    let last = countPendingImages(tbl);
    let stuck = 0;
    for (let i = 0; i < 25 && countPendingImages(tbl) > 0; i++) {
      await sleep(200);
      const now = countPendingImages(tbl);
      stuck = now >= last ? stuck + 1 : 0;
      last = now;
      if (stuck >= 3) break;
    }
  } finally {
    scroller.set(origin);
  }
  await sleep(120);
  return countPendingImages(tbl);
}

// ============================ 字段解析 ============================

interface Product {
  图片: string;
  尺寸: string;
  件数: string;
}

/** 从一段商品文本解析尺寸（形如 "... color ：Purple size ：2' x 6'"） */
function parseSize(seg: string): string {
  // 注意：size 在【最后】，其后没有 color，不能假设「size 后面跟着 color」
  const sizeM = seg.match(/size\s*[：:]\s*(.+)/i);
  if (!sizeM) return '';
  let v = sizeM[1] ?? '';
  const cIdx = v.search(/\s+color\s*[：:]/i);
  if (cIdx >= 0) v = v.slice(0, cIdx);
  return v.trim();
}

/**
 * 尺寸换算成厘米：数值 × 30 并取整（业务口径，1 英尺 ≈ 30cm），表头写作「尺寸(CM)」。
 *
 * 规则 = **只乘能乘的，其余原样**：
 *   1) 先去掉紧跟数字的英制单位记号（' / ′ / " / ft / feet / inch / inches / in）——
 *      换算完就该是厘米了，再留着英尺引号是自相矛盾；
 *      ⚠️ 这里不能用 `\bft\b`：`2ft` 的 `2` 与 `f` 都是单词字符，边界不成立，整段会漏掉；
 *   2) `x` / `×` 是分隔符（`2x7ft` = 2 x 7 ft），先把贴着数字的 x 拆成独立的 ` x `，
 *      否则它会被当成单位、把前面的数字一起跳过；
 *   3) 逐个**独立**数字 × 30 并取整（2.54 → 76，0.7 → 21）；
 *      字母与数字粘在一起的整段不动（`A4`、`2XL`、`12cm`、`12.5cm` 是型号/带单位的值）；
 *   4) 整段一个可乘的数字都没有 → 原样返回，绝不写进猜出来的值。
 *
 * 例：`2ft` → `60`；`2x7ft` → `60 x 210`；`2' x 6'` → `60 x 180`；`20*15*10` → `600*450*300`。
 */
function toCmSize(raw: string): string {
  const s = (raw ?? '').trim();
  if (!s || !/\d/.test(s)) return s;

  // 1) 去英制单位（长的放前面，避免 inch 被 in 吃掉半个；负向断言只挂在字母单位上，
  //    否则 `2'x6'` 里的引号会被后面的 x 挡住、删不掉）
  const noUnit = s
    .replace(
      /(\d)\s*(?:(?:'|′|”|")|(?:inches|feet|foot|inch|ft|in)(?![a-z]))/gi,
      '$1',
    )
    .trim();

  // 2) 拆开贴着数字的分隔符 x
  const sep = noUnit.replace(/(\d)\s*[x×]\s*(?=\d)/g, '$1 x ');

  // 3) 逐个数字换算；含字母的 token（型号 / 带单位）原样保留
  return sep.replace(
    /\d+(?:\.\d+)?[A-Za-z]+|[A-Za-z]+\d+(?:\.\d+)?|\d+(?:\.\d+)?/g,
    (t) => {
      if (/[A-Za-z]/.test(t)) return t;
      return String(Math.round(Number(t) * 30));
    },
  );
}

/**
 * 从一段商品文本解析件数（形如 "6W-W7WR-L9I5 x 1 USD 49.99"）。
 *
 * 注意：绝不能简单取第一个 "x N"——SKU 自身可能含 "X3"（如 0L-VJC9-X3H4），
 * 会被误读成件数 3，而真实件数是它后面那个 "x 1"。
 * 因此必须锚定「SKU 之后」的 x N。
 */
function parseQty(seg: string): string {
  // 优先：SKU（形如 XX-XXXX-XXXX）后紧跟的 x N
  const m = seg.match(/[A-Z0-9]{2,4}-[A-Z0-9]{4}-[A-Z0-9]{4}\s*[x×]\s*(\d+)/i);
  if (m) return m[1] ?? '';
  // 兜底：货币金额前的 x N（如 "x 1 USD 49.99"）
  const m2 = seg.match(/[x×]\s*(\d+)\s*(?=USD|RMB|CNY|EUR|GBP|JPY|￥|\$)/i);
  if (m2) return m2[1] ?? '';
  return '';
}

/**
 * 从「商品信息」单元格解析出一个或多个商品。
 *
 * 首选按【商品块】逐个解析：每个商品在页面里是一个独立的 `.order-sku`
 * （图片在 `.order-sku__image`，文本在 `.order-sku__info`），
 * 图片与文本天然一一对应 —— 绝不会出现「图片错位 / 多商品丢图」。
 *
 * 兜底（页面没有 .order-sku 结构时）：按渲染出的 "!" 拆段，一段=一个商品。
 * 这里不再「先过滤占位图再按下标对齐」—— 那样会让占位图把下标挤歪，
 * 导致第 1 段拿到第 3 张的图。改为按下标原位取，取不到就留空。
 *
 * 图片为懒加载：未加载时 src 是 1x1 占位 base64（data:image/gif...），
 * 此时真实地址不在 DOM 里，imgUrl 返回 ''（见其注释）。
 */
function parseProducts(cell: Element | undefined): Product[] {
  if (!cell) return [{ 图片: '', 尺寸: '', 件数: '' }];

  // ---- 1) 按商品块解析（真实结构） ----
  const blocks = Array.from(cell.querySelectorAll('.order-sku'));
  if (blocks.length > 0) {
    return blocks.map((b) => {
      const text = (b.textContent ?? '').replace(/\s+/g, ' ').trim();
      const im =
        b.querySelector('.order-sku__image img') ?? b.querySelector('.imageContainer img') ?? b.querySelector('img');
      return { 图片: imgUrl(im), 尺寸: parseSize(text), 件数: parseQty(text) };
    });
  }

  // ---- 2) 兜底：按 "!" 拆段，图片按下标原位对齐 ----
  const imgs = Array.from(cell.querySelectorAll('img')).map((im) => imgUrl(im));
  const text = (cell.textContent ?? '').replace(/\s+/g, ' ').trim();
  const segs = text
    .split(/\s*!\s*/)
    .map((s) => s.trim())
    .filter(Boolean);
  const list = segs.length > 0 ? segs : [text];

  return list.map((seg, i) => ({
    图片: imgs[i] ?? (list.length === 1 ? (imgs[0] ?? '') : ''),
    尺寸: parseSize(seg),
    件数: parseQty(seg),
  }));
}

/** 从「物流方式」单元格解析 运单号（如 「YT2626700708781963」） */
function parseTracking(text: string): string {
  const bracket = text.match(/[「『]\s*([A-Za-z0-9]+)\s*[」』]/);
  if (bracket) return bracket[1] ?? '';
  const m = text.match(/\b([A-Z]{2}\d{8,})\b/);
  return m ? (m[1] ?? '') : '';
}

/**
 * 把一个数据行 Map<标题,单元格> 转成【一条或多条】导出记录。
 * 一个订单可能含多个商品 → 每个商品导出一行（共享同一运单号）。
 */
function toRecords(rowMap: Map<string, Element>): Array<Record<string, string>> {
  const goodsCell = pickCell(rowMap, CONFIG.sourceColumns['商品信息']);
  const logisticsCell = pickCell(rowMap, CONFIG.sourceColumns['物流方式']);

  const products = parseProducts(goodsCell);
  const 运单号 = parseTracking((logisticsCell?.textContent ?? '').trim());

  return products.map((p) => ({
    图片: p.图片,
    尺寸: toCmSize(p.尺寸),
    件数: p.件数,
    材质: CONFIG.defaultMaterial, // 页面没有「材质」字段，统一填默认值
    运单号,
  }));
}

// ============================ 用户名探测 ============================

interface UsernameHit {
  name: string;
  /** 命中来源，便于排查（选择器字面量 / storage 键 / heuristic） */
  source: string;
  /**
   * 是否为「高可信」命中。
   *
   * 只有命中**已知的账号名 DOM 结构**（`.user-name` 那一族选择器）才算高可信。
   * storage / cookie / 顶部文本启发式都是**猜**——它们很容易命中
   * 会话 ID（如 `952C1DB87F4EDF20`）、子账号标识之类的东西。
   *
   * 这个标记决定了轮询是否停止：账号区是 Vue 异步渲染的，早期只能用低可信来源兜底，
   * 但绝不能因此停止重试 —— 否则真正的 `.user-name` 渲染出来后也不会再被检查。
   */
  strong: boolean;
}

/** 清理并校验一个候选用户名（去掉「你好，」之类前缀、换行、多余空白） */
function cleanUsername(raw: string): string {
  let s = (raw ?? '').replace(/\s+/g, ' ').trim();
  // 去掉常见的问候/标签前缀，只留名字本体
  s = s.replace(/^(你好|您好|hi|hello|welcome|欢迎|账号|用户名|用户|店铺|hi,|hello,)[\s,，:：]*/i, '');
  return s.trim();
}

/**
 * 判断候选是否「像」一个用户名。
 *
 * @param strict 严格模式 —— 用于 storage / cookie / 启发式这类**猜测**来源。
 *   它们最容易捞出会话 ID、设备码之类的脏值（真实踩坑：`952C1DB87F4EDF20`），
 *   所以额外拒绝「十六进制串」「长串大写字母+数字」。
 *   DOM 选择器命中的元素本身就是账号名容器，不需要这么严（仍做基本校验）。
 */
function looksLikeUsername(s: string, strict = false): boolean {
  if (!s) return false;
  if (s.length < 1 || s.length > 24) return false;
  // 含明显非名字符号（路径、括号、多个标点）的直接排除
  if (/[\/\\<>{}()[\]|]/.test(s)) return false;
  // 短纯数字是角标/数量（真实页面里通知角标就是 `title="15"`），不是账号名；
  // 但 6 位以上的纯数字保留 —— 店小秘允许手机号登录，账号名可能就是手机号。
  if (/^\d+$/.test(s) && s.length < 6) return false;
  // 至少要有一个字母/汉字，或是一串 ≥6 位的数字；否则视为纯符号
  if (!/[A-Za-z\u4e00-\u9fa5]/.test(s) && !/^\d{6,}$/.test(s)) return false;
  if (strict) {
    // 纯十六进制且够长 → 设备号/会话 token（真实案例 `952C1DB87F4EDF20`）
    if (/^[0-9A-F]{12,}$/i.test(s)) return false;
    // 没有任何小写字母、没有汉字，且长度 ≥ 10 → 几乎都是 ID 而非账号名
    // （`jia_yangdong` / `张攀` / `abc123` 都能过，`952C1DB87F4EDF20` 过不了）
    if (s.length >= 10 && !/[a-z\u4e00-\u9fa5]/.test(s)) return false;
  }
  const low = s.toLowerCase();
  if (CONFIG.usernameStopWords.some((w) => low === w.toLowerCase())) return false;
  return true;
}

/**
 * 多路探测当前登录用户名，按可靠性从高到低：
 *   1) 明确的 DOM 选择器
 *   2) localStorage / sessionStorage / cookie 里 key 含 user|account|name 的短值
 *   3) 顶部区域（header/nav 等）内「像名字」的文本启发式
 * 每一条命中都会记进日志，方便定位店小秘改版后的真实位置。
 */
/**
 * 需要探测的文档列表：主文档 + 所有【同源】iframe。
 *
 * 店小秘是典型的多框架后台：顶部账号栏（head-nav-right / btn-user / user-name）
 * 和业务列表有可能不在同一个 document 里。content script 默认只注入主框架，
 * 只查 `document` 就会「看得见列表、却看不到账号」，账号便一直识别不到。
 * 同源 iframe 可以直接拿 contentDocument；跨域会抛错，catch 掉即可。
 */
function candidateDocs(): Document[] {
  const docs: Document[] = [document];
  try {
    for (const f of Array.from(document.querySelectorAll('iframe'))) {
      try {
        const d = f.contentDocument;
        if (d && d !== document) docs.push(d);
      } catch {
        /* 跨域 iframe，忽略 */
      }
    }
  } catch {
    /* ignore */
  }
  return docs;
}

/** 文档的显示名，用于日志 / 诊断报告（top / iframe#1 ...） */
function docLabel(i: number): string {
  return i === 0 ? 'top' : `iframe#${i}`;
}

function detectUsername(): UsernameHit {
  const tried: string[] = [];
  // 记录引用，诊断报告可直接读到「看过但被否掉」的候选
  lastTried = tried;
  const docs = candidateDocs();
  if (docs.length > 1) {
    console.log(`[店小秘导出] 检测到 ${docs.length - 1} 个同源 iframe，账号名会在其中一起查找`);
  }

  // ---- 1) DOM 选择器（主文档 + 同源 iframe） ----
  for (let di = 0; di < docs.length; di++) {
    const doc = docs[di]!;
    const where = docLabel(di);
    for (const sel of CONFIG.usernameSelectors) {
      try {
        const el = doc.querySelector(sel);
        if (!el) continue;
        // 店小秘把用户名同时写进 title 属性（实测 `title="jia_yangdong"`）。
        // title 比 textContent 干净：不受角标/子元素拼接影响，所以优先采信。
        const fromTitle = cleanUsername(el.getAttribute('title') ?? '');
        const fromText = cleanUsername((el.textContent ?? '').trim());
        const hit = looksLikeUsername(fromTitle) ? fromTitle : fromText;
        if (hit && looksLikeUsername(hit)) {
          console.log('[店小秘导出] 用户名命中（选择器，高可信）', `${where} ${sel}`, '=', hit);
          // 命中已知结构 = 高可信：这是真正代表账号名的元素，可以停止重试
          return { name: hit, source: `selector:${sel}@${where}`, strong: true };
        }
        tried.push(`${where} selector ${sel} -> "${fromText}"(不像用户名)`);
      } catch {
        /* ignore invalid selector */
      }
    }
  }

  // ---- 2) storage / cookie（每个文档各自的 storage 都扫） ----
  //
  // 注意：这一路是**低可信猜测**。真实踩坑——账号区还没渲染时降级到这里，
  // 从某个 `user*` 键里读出了会话/设备 ID（`952C1DB87F4EDF20`），长度与字符集
  // 都通过了宽松校验，于是被当成用户名用了。
  // 因此：key 名排除更多「ID 语义」词 + 值走严格校验 + 标为低可信（不停止重试）。
  const keyRe = /(user|account|login|nick|shop|owner)/i;
  // key 名里只要出现这些语义片段，就是 ID/token 而非账号名
  const keyDeny =
    /(id|token|time|expire|avatar|url|path|level|type|status|count|session|sid|uuid|guid|hash|key|secret|sign|device|machine|code|sn|no|random|salt|verif|ticket|auth)/i;
  for (let di = 0; di < docs.length; di++) {
    const doc = docs[di]!;
    const where = docLabel(di);
    const stores: Array<[string, Storage | null]> = [];
    try {
      stores.push([`${where}.localStorage`, doc.defaultView?.localStorage ?? null]);
      stores.push([`${where}.sessionStorage`, doc.defaultView?.sessionStorage ?? null]);
    } catch {
      /* 某些环境下禁止访问 storage */
    }
    for (const [label, store] of stores) {
      if (!store) continue;
      try {
        for (let i = 0; i < store.length; i++) {
          const k = store.key(i) ?? '';
          if (!keyRe.test(k)) continue;
          if (keyDeny.test(k)) continue;
          const v = store.getItem(k) ?? '';
          if (!v || v.startsWith('{') || v.startsWith('[')) continue;
          const c = cleanUsername(v);
          if (looksLikeUsername(c, true)) {
            console.log('[店小秘导出] 用户名命中（storage，低可信）', label, k, '=', c);
            return { name: c, source: `${label}:${k}`, strong: false };
          }
          tried.push(`${label} ${k} -> "${v.slice(0, 30)}"(不像)`);
        }
      } catch {
        /* ignore */
      }
    }
    try {
      for (const pair of doc.cookie.split(';')) {
        const eq = pair.indexOf('=');
        if (eq < 0) continue;
        const k = pair.slice(0, eq).trim();
        if (!keyRe.test(k)) continue;
        if (keyDeny.test(k)) continue;
        const v = decodeURIComponent(pair.slice(eq + 1).trim());
        if (!v || v.startsWith('{') || v.length > 24) continue;
        const c = cleanUsername(v);
        if (looksLikeUsername(c, true)) {
          console.log('[店小秘导出] 用户名命中（cookie，低可信）', k, '=', c);
          return { name: c, source: `cookie:${k}`, strong: false };
        }
      }
    } catch {
      /* ignore */
    }
  }

  // ---- 3) 顶部区域启发式 ----
  for (let di = 0; di < docs.length; di++) {
    const doc = docs[di]!;
    const containers = Array.from(
      doc.querySelectorAll('header, [class*="header"], [class*="topbar"], [class*="top-bar"], [class*="nav"]'),
    ).slice(0, 6);
    for (const c of containers) {
      const nodes = Array.from(c.querySelectorAll('span, div, a, p, b, em')).slice(0, 400);
      for (const n of nodes) {
        if (n.children.length > 0) continue; // 只要叶子节点，避免拿到整块导航文案
        const t = cleanUsername((n.textContent ?? '').trim());
        if (!t || !looksLikeUsername(t, true)) continue;
        // 名字旁边通常有头像/用户图标
        const near = n.parentElement?.querySelector('img, [class*="avatar"], [class*="icon-user"], [class*="user"]');
        if (near) {
          console.log('[店小秘导出] 用户名命中（启发式，低可信）', t, n.className);
          return { name: t, source: `heuristic:${t}`, strong: false };
        }
        tried.push(`heuristic "${t}"(无头像相邻)`);
      }
    }
  }

  console.log('[店小秘导出] 用户名未识别。已尝试：', tried);
  return { name: '', source: '', strong: false };
}

/** 读取上次记住的用户名（探测失败时的兜底，避免每次手填） */
async function loadSavedUsername(): Promise<string> {
  try {
    const r = await browser.storage?.local?.get(CONFIG.storageKey);
    const v = r?.[CONFIG.storageKey];
    return typeof v === 'string' ? v : '';
  } catch {
    return '';
  }
}

/** 记住本次使用的用户名 */
async function saveUsername(name: string): Promise<void> {
  if (!name) return;
  try {
    await browser.storage?.local?.set({ [CONFIG.storageKey]: name });
  } catch {
    /* storage 不可用时静默忽略 */
  }
}

/**
 * 采集一份「账号名到底藏在哪里」的诊断报告。
 * 店小秘改版后选择器会失效，靠这份报告可以一次性定位真实位置。
 */
function buildUsernameReport(): Record<string, unknown> {
  const docs = candidateDocs();
  const frames: Array<Record<string, unknown>> = [];

  for (let di = 0; di < docs.length; di++) {
    const doc = docs[di]!;
    const where = docLabel(di);

    const selectors = CONFIG.usernameSelectors
      .map((sel) => {
        try {
          const el = doc.querySelector(sel);
          if (!el) return null;
          return {
            sel,
            title: el.getAttribute('title') ?? null,
            text: (el.textContent ?? '').trim().slice(0, 40),
          };
        } catch {
          return { sel, error: 'invalid selector' };
        }
      })
      .filter(Boolean);

    const topTexts: string[] = [];
    const containers = Array.from(
      doc.querySelectorAll('header, [class*="header"], [class*="topbar"], [class*="nav"], [class*="user"]'),
    ).slice(0, 5);
    for (const c of containers) {
      for (const n of Array.from(c.querySelectorAll('span, div, a, p, b, em')).slice(0, 120)) {
        if (n.children.length > 0) continue;
        const t = (n.textContent ?? '').trim();
        if (t && t.length <= 24 && !topTexts.includes(t)) topTexts.push(t);
        if (topTexts.length >= 60) break;
      }
    }

    let href = '';
    try {
      href = doc.location?.href ?? '';
    } catch {
      href = '(跨域不可读)';
    }

    frames.push({ where, url: href, selectors, topTexts });
  }

  // storage 只取主文档的（够定位问题，避免报告过大）
  const storageDump: Record<string, string> = {};
  try {
    for (let i = 0; i < localStorage.length && i < 40; i++) {
      const k = localStorage.key(i) ?? '';
      storageDump[k] = (localStorage.getItem(k) ?? '').slice(0, 80);
    }
  } catch {
    /* ignore */
  }

  return {
    buildTag: CONFIG.buildTag,
    pageUrl: location.href,
    detected: detectedUsername,
    detectedSource,
    detectedStrong,
    detectedFromMemory,
    /** 探测过程中被拒绝的候选（含值），用于解释「为什么识别成了别的值」 */
    rejected: lastTried.slice(-40),
    frameCount: docs.length,
    frames,
    localStorage: storageDump,
    cookieNames: document.cookie
      .split(';')
      .map((s) => s.split('=')[0]?.trim())
      .filter(Boolean),
  };
}

/** 面板当前生效的自动识别账号（手工填写不算）；面板与诊断报告共用同一份状态 */
let detectedUsername = '';
let detectedSource = '';
let detectedStrong = false;
/** 该账号名来自「上次导出时记住的」——用户此前已经确认过一次，可不再要求重复确认 */
let detectedFromMemory = false;
/**
 * 最近一次探测中「看过但被否掉」的候选。
 * 这里存的是 detectUsername 内部那个数组的**引用**，所以后续 push 会实时反映出来。
 * 诊断时最有用的就是它 —— 能直接看出账号名被什么值/什么来源抢先占用了。
 * （真实案例：`top.localStorage userKey -> "952C1DB87F4EDF20"(不像)` 就能一眼看出问题。）
 */
let lastTried: string[] = [];

// ============================ 导出 ============================

/** base64 -> Uint8Array（ExcelJS addImage 需要 buffer） */
function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

/** 触发浏览器下载 */
function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

/** 一张抓下来并处理好的图片 */
interface PreparedImage {
  dataUrl: string;
  mime: string;
}

/**
 * 经 background 下载图片；失败返回 null。
 * （content script 直接 fetch 会被 CORS 拦，必须走 background 的主机权限）
 *
 * 会一并把「高清候选链」带过去，由 background 依次尝试，取第一个真正抓到的。
 * 这里**不打日志** —— 失败要在导出结束时统一汇报（同一张图可能被多行引用，
 * 逐次 warn 会刷屏），见 exportToExcel 的统计。
 */
async function fetchImageViaBackground(url: string): Promise<PreparedImage | null> {
  try {
    const resp = (await browser.runtime.sendMessage({
      type: 'DXM_FETCH_IMAGE',
      url,
      candidates: buildImageCandidates(url),
    })) as
      | { ok: boolean; base64?: string; mime?: string; error?: string; url?: string }
      | undefined;
    if (resp && resp.ok && resp.base64) {
      if (resp.url && resp.url !== url) {
        console.log('[店小秘导出] 已换高清图:', url, '->', resp.url);
      }
      const mime = resp.mime || 'image/jpeg';
      return { dataUrl: `data:${mime};base64,${resp.base64}`, mime };
    }
    return null;
  } catch {
    return null;
  }
}

/** 从 dataURL 里读出 mime（`data:image/png;base64,...` → `image/png`） */
function mimeOfDataUrl(dataUrl: string): string {
  const m = /^data:([^;,]+)/.exec(dataUrl);
  return m?.[1] || 'image/jpeg';
}

/**
 * ExcelJS addImage 的 extension 必须与图片真实格式一致，
 * 标成 jpeg 却塞 png 数据，Excel 打开就是空白/黑块。
 */
function extOfMime(mime: string): 'jpeg' | 'png' | 'gif' {
  const m = (mime || '').toLowerCase();
  if (m.includes('png')) return 'png';
  if (m.includes('gif')) return 'gif';
  return 'jpeg';
}

/**
 * 把图片等比缩到最长边 IMG_MAX_PX，并统一转成【白底 JPEG】。
 *
 * 为什么是 JPEG 而不是 PNG：
 *   Excel 只认得 png/jpeg/gif。转 PNG 是无损的，但一张 900px 商品图 PNG 动辄
 *   500KB~1MB，几十行就把 xlsx 撑到几十 MB；同样清晰度的 JPEG 只要 100KB 左右。
 *   商品图基本是白底照片，先铺白再画，避免透明区域被压成黑块。
 *
 * 转换失败（图片解码不了 / canvas 拿不到）时**原样返回**，并保留原始 mime ——
 * 调用方必须按这个 mime 注册 extension，否则格式对不上，Excel 里不显示。
 */
function toJpegDataUrl(dataUrl: string, mime?: string): Promise<PreparedImage> {
  const original: PreparedImage = { dataUrl, mime: mime ?? mimeOfDataUrl(dataUrl) };
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      try {
        const w = img.naturalWidth || CONFIG.imgMaxPx;
        const h = img.naturalHeight || CONFIG.imgMaxPx;
        const scale = Math.min(1, CONFIG.imgMaxPx / Math.max(w, h));
        const c = document.createElement('canvas');
        c.width = Math.max(1, Math.round(w * scale));
        c.height = Math.max(1, Math.round(h * scale));
        const ctx = c.getContext('2d');
        if (!ctx) return resolve(original);
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, c.width, c.height);
        ctx.drawImage(img, 0, 0, c.width, c.height);
        // dataURL 是同源的，canvas 不会被污染
        return resolve({
          dataUrl: c.toDataURL('image/jpeg', CONFIG.imgQuality),
          mime: 'image/jpeg',
        });
      } catch {
        return resolve(original);
      }
    };
    img.onerror = () => resolve(original);
    img.src = dataUrl;
  });
}

/** 导出结果统计 —— 图片没进 Excel 时必须有迹可循，不能只留一句「导出成功」 */
interface ExportStats {
  rows: number;
  /** 成功嵌入图片的行数 */
  withImage: number;
  /** 页面里就没拿到图片地址的行数（懒加载没触发 / 页面本身无图） */
  noUrl: number;
  /** 拿到地址但最终没下载成功的图片数（按唯一 URL 计） */
  failed: number;
  failedUrls: string[];
}

/**
 * 导出 Excel：表头 + 当前列表的全部数据行，并把商品图【嵌入】到「图片」列。
 *
 * filename 由调用方算好（用户名 + 导出当天日期），这里不再关心命名规则。
 * 返回统计（含图/无地址/下载失败），由调用方展示在面板提示里。
 */
async function exportToExcel(
  rows: Array<Record<string, string>>,
  filename: string,
  onProgress?: (done: number, total: number) => void,
): Promise<ExportStats> {
  // ---- 1) 下载用到的图片（同一 URL 只下一次） ----
  const urls: string[] = Array.from(
    new Set(rows.map((r) => r.图片).filter((u): u is string => !!u)),
  );
  const imgCache = new Map<string, PreparedImage>();

  const loadOne = async (u: string): Promise<void> => {
    const raw = await fetchImageViaBackground(u);
    if (raw) imgCache.set(u, await toJpegDataUrl(raw.dataUrl, raw.mime));
  };

  for (let i = 0; i < urls.length; i++) {
    onProgress?.(i + 1, urls.length);
    await loadOne(urls[i]!);
  }

  // 网络抖动 / CDN 限流会偶发失败：隔一会儿再试一轮（最多 20 张，免得整体卡太久）
  const retryList = urls.filter((u) => !imgCache.has(u)).slice(0, 20);
  if (retryList.length) {
    await sleep(800);
    for (const u of retryList) await loadOne(u);
  }

  const failedUrls = urls.filter((u) => !imgCache.has(u));
  if (failedUrls.length) {
    console.warn(`[店小秘导出] ${failedUrls.length} 张图片下载失败（已重试一次）:`, failedUrls);
  }

  // ---- 2) 建工作簿 ----
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('发货成功列表');

  const head = ws.addRow([...CONFIG.targetColumns]);
  head.font = { bold: true };
  head.height = 22;

  // 列宽：图片列留宽一些给缩略图
  ws.getColumn(1).width = CONFIG.imgColWidth;
  ws.getColumn(2).width = 16; // 尺寸
  ws.getColumn(3).width = 8; // 件数
  ws.getColumn(4).width = 12; // 材质
  ws.getColumn(5).width = 26; // 运单号

  // ---- 3) 逐行写数据 + 插图 ----
  let withImage = 0;
  let noUrl = 0;
  rows.forEach((r, i) => {
    const rowNo = i + 2; // 1-based 行号（表头占第 1 行）
    ws.addRow(['', r.尺寸, r.件数, r.材质, r.运单号]);
    ws.getRow(rowNo).height = CONFIG.imgRowHeight;

    if (!r.图片) {
      noUrl++;
      return;
    }
    const prepared = imgCache.get(r.图片);
    if (!prepared) return; // 下载失败，格子留空（failed 已统计）
    const b64 = prepared.dataUrl.slice(prepared.dataUrl.indexOf(',') + 1);
    try {
      // ExcelJS 的类型声明写的是 Node 的 Buffer，浏览器环境实际传 Uint8Array，故整体断言
      const imageId = wb.addImage({
        buffer: base64ToBytes(b64),
        extension: extOfMime(prepared.mime),
      } as unknown as Parameters<typeof wb.addImage>[0]);
      ws.addImage(imageId, {
        tl: { col: 0, row: rowNo - 1 }, // tl 是 0-based
        ext: { width: CONFIG.imgPx, height: CONFIG.imgPx },
        editAs: 'oneCell',
      });
      withImage++;
    } catch (e) {
      console.warn('[店小秘导出] 插图失败:', r.图片, e);
    }
  });

  // ---- 4) 写出并下载 ----
  const buf = await wb.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
    filename,
  );

  return {
    rows: rows.length,
    withImage,
    noUrl,
    failed: failedUrls.length,
    failedUrls,
  };
}

/**
 * 生成导出文件名：`用户名_导出当天年月日.xlsx`，例如 `jia_yangdong_20260925.xlsx`。
 *
 * 日期固定取【执行导出的当天】，标记的是「这批数据是什么时候导的」。
 * 非法文件名字符（\ / : * ? " < > |）替换成下划线。
 */
function buildFileName(username: string): string {
  const base = (username || '店小秘').replace(/[\\/:*?"<>|]/g, '_').trim() || '店小秘';
  return `${base}_${ymd(startOfDay())}.xlsx`;
}

// ============================ 面单下载 ============================

/** 面单文件名：`用户名_当天年月日_面单.pdf`（和 Excel 同账号同日期，方便配对） */
function buildLabelFileName(username: string): string {
  const base = (username || '店小秘').replace(/[\\/:*?"<>|]/g, '_').trim() || '店小秘';
  return `${base}_${ymd(startOfDay())}_面单.pdf`;
}

/** background 回传的面单 PDF */
interface PdfResp {
  ok: boolean;
  url?: string;
  base64?: string;
  mime?: string;
  error?: string;
}

function isVisible(el: Element): boolean {
  const r = el.getBoundingClientRect();
  if (r.width <= 0 || r.height <= 0) return false;
  const st = getComputedStyle(el);
  return st.display !== 'none' && st.visibility !== 'hidden';
}

/** 页面上文本正好等于 `text` 且可见的元素，按可信度排序 */
function visibleByText(text: string): HTMLElement[] {
  const sel = 'a, button, .ant-btn, li, .ant-dropdown-menu-item, span, div';
  const found = Array.from(document.querySelectorAll<HTMLElement>(sel)).filter(
    (el) => el.textContent?.trim() === text && isVisible(el),
  );
  // 下拉菜单是 portal（挂在 body 末尾），但结构里未必带 .ant-dropdown 类，
  // 所以按「在菜单里 > 不在表格行里 > 其它」排序，保证不会先点到行内同名链接。
  const inMenu = found.filter((el) =>
    el.closest('.ant-dropdown, .ant-dropdown-menu, [class*="dropdown"], [class*="popover"]'),
  );
  const outsideTable = found.filter((el) => !el.closest('tbody') && !inMenu.includes(el));
  const rest = found.filter((el) => !inMenu.includes(el) && !outsideTable.includes(el));
  return [...inMenu, ...outsideTable, ...rest];
}

/**
 * 打开面单入口，按真实流程：**点「批量打印」→ 点菜单里的「打印面单」**。
 * 菜单已经展开时直接点那一项；连批量菜单都没有时，退回页面上直接写着
 * 「打印面单」的入口（每行操作列里那个）。
 *
 * 返回是否点到了东西。
 */
async function clickPrintEntry(): Promise<boolean> {
  // 菜单/工具栏上的那个（**不碰行内链接**，行内是兜底）
  const menuHit = (): HTMLElement | null =>
    visibleByText('打印面单').find((el) => !el.closest('tbody')) ?? null;

  // 1) 菜单已经展开 → 直接点
  const direct = menuHit();
  if (direct) {
    direct.click();
    return true;
  }

  // 2) 点「批量打印」把菜单展开。
  //    触发器是 `<button class="ant-btn ant-btn-primary ant-dropdown-trigger">`，
  //    antd 下拉可能是 hover 触发，所以 hover 事件和 click 都发；
  //    菜单可能晚一帧才挂出来，循环等，但**只 click 一次**（重复 click 会把它点关）。
  const isTrigger = (t: string): boolean => t === '批量打印' || (t.startsWith('批量打印') && t.length <= 12);
  const findTrigger = (sel: string): HTMLElement | undefined =>
    Array.from(document.querySelectorAll<HTMLElement>(sel))
      .filter((el) => isVisible(el))
      .map((el) => ({ el, t: (el.textContent ?? '').replace(/\s+/g, ' ').trim() }))
      .find(({ t }) => isTrigger(t))?.el;
  const trigger =
    findTrigger('a, button, .ant-btn') ?? findTrigger('.ant-dropdown-trigger') ?? findTrigger('span');

  if (trigger) {
    const hover = () => {
      for (const type of ['mouseenter', 'mouseover'] as const) {
        trigger.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window }));
      }
    };
    hover();
    trigger.click();
    for (let i = 0; i < 5; i++) {
      const item = menuHit();
      if (item) {
        item.click();
        return true;
      }
      hover();
      await sleep(500);
    }
  }

  // 3) 兜底：菜单没出来，直接点行内那个「打印面单」
  //    （打印管理弹窗按「勾选数量」出单，所以勾了全部时效果一样）
  const fallback = visibleByText('打印面单')[0];
  if (fallback) {
    fallback.click();
    return true;
  }
  return false;
}

/**
 * 勾选当页全部订单 —— 打印管理弹窗打的是「勾选的那批」，不勾就是空的/只有当前行。
 * 返回还原函数，把复选框恢复成点之前的样子（用户可能自己选了几行）。
 */
async function selectAllOrders(bodyTable: HTMLTableElement): Promise<() => void> {
  const scope = bodyTable.closest('.vxe-table') ?? bodyTable.parentElement ?? document.body;
  const collect = (): HTMLInputElement[] =>
    Array.from(scope.querySelectorAll<HTMLInputElement>('input[type="checkbox"]'));
  const boxes = collect();
  if (boxes.length < 2) return () => undefined; // 没有行级复选框，谈不上全选

  const before = boxes.slice(1).map((b) => b.checked);
  if (before.every(Boolean)) return () => undefined; // 本来就全选，别动

  const allBox = boxes[0];
  if (allBox && !allBox.checked) allBox.click();
  await sleep(600);

  // 兜底：表头全选没生效就逐行点（点多少算多少，至少别是空的）
  for (const b of collect().slice(1)) if (!b.checked) b.click();

  return () => {
    const now = collect().slice(1);
    if (now.length !== before.length) return; // 行数变了，硬还原可能点错行
    if (before.every((v) => !v)) {
      // 原本一行都没勾 → 点一下表头全选清空，比逐行点快
      const all = collect()[0];
      if (all?.checked) {
        all.click();
        return;
      }
    }
    now.forEach((b, i) => {
      if (b.checked !== before[i]) b.click();
    });
  };
}

/** 「打印管理」弹窗：.ant-modal-content，标题写死是「打印管理」 */
function findPrintModal(): HTMLElement | null {
  return (
    Array.from(document.querySelectorAll<HTMLElement>('.ant-modal-content')).find(
      (m) =>
        isVisible(m) &&
        (m.querySelector('.ant-modal-title')?.textContent ?? '')
          .replace(/\s+/g, '')
          .includes('打印管理'),
    ) ?? null
  );
}

async function waitPrintModal(timeoutMs = 10000): Promise<HTMLElement | null> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const m = findPrintModal();
    if (m) return m;
    await sleep(200);
  }
  return null;
}

/** 弹窗里那句「当前选中 N 条数据」—— 确认确实勾到了东西。读不到返回 -1（不武断判 0） */
function readSelectedCount(modal: HTMLElement): number {
  const txt = (modal.querySelector('.ant-modal-body')?.textContent ?? '').replace(/\s+/g, '');
  const m = /当前选中(\d+)条/.exec(txt);
  return m ? Number(m[1]) : -1;
}

/**
 * 打印方式三档，必须保证落在第一档「PDF生成打印 (仅打印选中 [当前页])」：
 * 只有它把当页勾选的订单合成**一个 PDF**；
 * 第二档要打印驱动，第三档每包一个文件（会变成 N 个 PDF）。
 */
function ensurePdfCombineMode(modal: HTMLElement): boolean {
  const radios = Array.from(modal.querySelectorAll<HTMLInputElement>('input[type="radio"]'));
  const target = radios.find((r) => (r.closest('label')?.textContent ?? '').includes('PDF生成打印'));
  if (!target) return false;
  if (!target.checked) (target.closest('label') ?? target).click();
  return true;
}

function findStartPrintBtn(modal?: HTMLElement | null): HTMLElement | null {
  const norm = (s: string | null | undefined): string => (s ?? '').replace(/\s+/g, '');
  const scopes: Element[] = modal
    ? [modal]
    : Array.from(
        document.querySelectorAll<HTMLElement>('.ant-modal, .ant-drawer, .el-dialog, [role="dialog"]'),
      );
  for (const scope of scopes) {
    for (const el of scope.querySelectorAll<HTMLElement>('button, .ant-btn, a, span, div')) {
      if (norm(el.textContent) === '开始打印' && isVisible(el)) return el;
    }
  }
  // 兜底：结构不认识时按按钮级元素全页找
  for (const el of document.querySelectorAll<HTMLElement>('button, a')) {
    if (norm(el.textContent) === '开始打印' && isVisible(el)) return el;
  }
  return null;
}

/** 等「开始打印」按钮出现（只在打印管理弹窗里找） */
async function waitStartPrintBtn(modal: HTMLElement, timeoutMs = 10000): Promise<HTMLElement | null> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const btn = findStartPrintBtn(modal);
    if (btn) return btn;
    await sleep(200);
  }
  return null;
}

/** 关掉打印管理弹窗：先点「取消」，再退右上角 ×，最后发一次 ESC */
function closePrintDialog(): void {
  const norm = (s: string | null | undefined): string => (s ?? '').replace(/\s+/g, '');
  const modal = findPrintModal();
  const cancel = modal
    ? Array.from(modal.querySelectorAll<HTMLElement>('button')).find(
        (b) => norm(b.textContent) === '取消' && isVisible(b),
      )
    : null;
  if (cancel) {
    cancel.click();
    return;
  }
  const close =
    modal?.querySelector<HTMLElement>('.ant-modal-close') ??
    Array.from(document.querySelectorAll<HTMLElement>('.ant-modal-close, .ant-drawer-close')).find(
      isVisible,
    ) ??
    null;
  if (close) {
    close.click();
    return;
  }
  document.dispatchEvent(
    new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true }),
  );
}

/**
 * 下载当页全部订单的面单（**一个** PDF，含所有勾选订单）。真实流程：
 *   勾选全部 → 点「批量打印」→ 点菜单里的「打印面单」→ 弹出「打印管理」
 *   → 确认打印方式 = 「PDF生成打印（仅打印选中 [当前页]）」→ 点「开始打印」
 *   → window.open 出 .pdf → background 抓下来 → 存 `账号_日期_面单.pdf`
 *   → 关弹窗 → 还原勾选。
 *
 * 为什么放在导出最前面：`window.open` 要点按钮后约 5s 内的用户激活，
 * 滚动列表 + 下载几十张商品图会把窗口耗光，放后面会被浏览器拦掉。
 */
async function downloadLabels(
  bodyTable: HTMLTableElement,
  username: string,
): Promise<{ ok: true; file: string } | { ok: false; error: string }> {
  const restore = await selectAllOrders(bodyTable);
  // 先让 background 开始等地址，再去点入口 —— PDF 地址可能几百毫秒就抛出来
  const waiter = browser.runtime.sendMessage({ type: 'DXM_AWAIT_PDF', timeout: 30000 }) as Promise<PdfResp>;
  const cancelWait = () => void browser.runtime.sendMessage({ type: 'DXM_CANCEL_PDF' });
  try {
    if (!(await clickPrintEntry())) {
      cancelWait();
      return { ok: false, error: '没找到「批量打印 → 打印面单」入口' };
    }
    const modal = await waitPrintModal(10000);
    if (!modal) {
      cancelWait();
      return { ok: false, error: '没弹出「打印管理」弹窗' };
    }

    const n = readSelectedCount(modal);
    if (n === 0) {
      cancelWait();
      closePrintDialog();
      return { ok: false, error: '弹窗显示「当前选中 0 条数据」，没勾到订单' };
    }
    ensurePdfCombineMode(modal);
    console.log('[店小秘导出] 打印管理：选中', n, '条；打印方式已设为「PDF生成打印」');

    const btn = await waitStartPrintBtn(modal, 10000);
    if (!btn) {
      cancelWait();
      closePrintDialog();
      return { ok: false, error: '弹窗里没找到「开始打印」按钮' };
    }
    btn.click();

    const resp = await waiter;
    if (!resp || !resp.ok || !resp.base64) {
      return { ok: false, error: resp?.error ?? '面单 PDF 下载失败' };
    }
    const bytes = base64ToBytes(resp.base64);
    const file = buildLabelFileName(username);
    downloadBlob(new Blob([bytes.buffer as ArrayBuffer], { type: resp.mime || 'application/pdf' }), file);
    closePrintDialog();
    return { ok: true, file };
  } finally {
    restore();
  }
}

// ============================ 面板 UI（Shadow DOM 隔离样式） ============================

function buildPanel(): HTMLElement {
  const host = document.createElement('div');
  host.id = '__dxm_export_panel';
  // 构建标记只作为**属性**挂在宿主元素上（界面不显示任何文案）。
  // 用途：确认浏览器里跑的是哪一版 —— Elements 里看 host，或探针读 host.dataset.build。
  host.dataset.build = CONFIG.buildTag;

  // 宿主元素在页面 light DOM 里，会被站点 CSS 影响（例如 `div { display:none !important }`）。
  // 用 inline !important 提升优先级：inline !important 高于作者样式表里的 !important，
  // 这样站点再怎么覆盖也无法把面板隐藏/挤走。
  const guard: Array<[string, string]> = [
    ['all', 'initial'],
    ['position', 'fixed'],
    ['right', '16px'],
    ['bottom', '16px'],
    ['left', 'auto'],
    ['top', 'auto'],
    ['z-index', '2147483647'],
    // 右对齐贴右下角（保持 flex 布局，方便面板靠右对齐）
    ['display', 'flex'],
    ['flex-direction', 'column'],
    ['align-items', 'flex-end'],
    ['gap', '8px'],
    ['visibility', 'visible'],
    ['opacity', '1'],
    ['transform', 'none'],
    ['filter', 'none'],
    ['pointer-events', 'auto'],
    ['margin', '0'],
    ['padding', '0'],
    ['width', 'auto'],
    ['height', 'auto'],
    ['max-width', 'none'],
    ['max-height', 'none'],
    ['overflow', 'visible'],
    ['clip', 'auto'],
    ['clip-path', 'none'],
    ['contain', 'none'],
    ['isolation', 'isolate'],
  ];
  for (const [prop, val] of guard) host.style.setProperty(prop, val, 'important');

  const shadow = host.attachShadow({ mode: 'open' });
  const style = document.createElement('style');
  style.textContent = `
    * { box-sizing: border-box; font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif; }
    .box { width: 260px; background: #fff; border: 1px solid #e3e3e3; border-radius: 10px;
           box-shadow: 0 6px 24px rgba(0,0,0,.16); overflow: hidden; }
    .hd { background: #2b6cff; color: #fff; padding: 10px 12px; font-size: 13px; font-weight: 600;
          display: flex; justify-content: space-between; align-items: center; }
    .hd .x { cursor: pointer; opacity: .85; font-weight: 400; }
    .bd { padding: 12px; font-size: 12px; color: #333; }
    .row { margin-bottom: 10px; }
    .lbl { display: block; margin-bottom: 4px; color: #666; }
    input.un { width: 100%; padding: 6px 8px; border: 1px solid #d8d8d8; border-radius: 6px; font-size: 12px; }
    .unbox { display: flex; align-items: center; gap: 6px; }
    .uname { flex: 1; min-width: 0; padding: 5px 8px; border: 1px dashed #dfe3ea; border-radius: 6px;
             background: #f7f9fc; color: #2b3a55; font-weight: 600;
             overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .uname.none { color: #e54545; font-weight: 400; border-style: dashed; }
    .uedit { flex: none; font-size: 11px; color: #2b6cff; cursor: pointer; text-decoration: underline; }
    .fname { margin-top: 6px; font-size: 11px; color: #5b6b85; word-break: break-all; }
    .fname b { color: #2b3a55; }
    .fname.err { color: #e54545; }
    .btn { width: 100%; padding: 9px 0; border: none; border-radius: 6px; background: #2b6cff; color: #fff;
           font-size: 13px; font-weight: 600; cursor: pointer; }
    .btn:disabled { background: #b9c8ee; cursor: not-allowed; }
    .tip { margin-top: 8px; font-size: 11px; color: #999; line-height: 1.5; min-height: 16px; }
    .tip.err { color: #e54545; }
    .rescan { display: inline-block; margin-top: 6px; font-size: 11px; color: #2b6cff;
              cursor: pointer; text-decoration: underline; }
  `;

  const box = document.createElement('div');
  box.className = 'box';
  box.innerHTML = `
    <div class="hd"><span>${CONFIG.panelTitle}</span><span class="x" title="收起面板">✕</span></div>
    <div class="bd">
      <div class="row">
        <span class="lbl">登录账号（自动识别）</span>
        <div class="unbox">
          <span class="uname" id="uname">识别中…</span>
          <span class="uedit" id="uedit" title="识别不到时可手动改">修改</span>
        </div>
        <input class="un" id="un" placeholder="请输入店小秘账号名" style="display:none" />
        <div class="fname" id="fname"></div>
      </div>
      <button class="btn" id="exp">导出Excel 和面单</button>
      <div class="tip" id="tip">正在扫描页面…</div>
      <span class="rescan" id="rescan">重新扫描</span>
    </div>
  `;

  shadow.appendChild(style);
  shadow.appendChild(box);
  return host;
}

// ============================ 主逻辑 ============================

export default defineContentScript({
  // 覆盖所有子域名与 http/https，仅限 dianxiaomi.com 域名
  matches: ['*://*.dianxiaomi.com/*'],
  main() {
    console.log('[店小秘导出] content script 已注入，当前 URL =', location.href);
    if (document.getElementById('__dxm_export_panel')) return;

    const panel = buildPanel();
    // 挂到 body 上（比 documentElement 更稳），并带兜底
    (document.body ?? document.documentElement).appendChild(panel);
    const shadow = panel.shadowRoot as ShadowRoot;

    const tip = shadow.getElementById('tip') as HTMLElement;
    const unInput = shadow.getElementById('un') as HTMLInputElement;
    const unameEl = shadow.getElementById('uname') as HTMLElement;
    const ueditEl = shadow.getElementById('uedit') as HTMLElement;
    const fnameEl = shadow.getElementById('fname') as HTMLElement;
    const expBtn = shadow.getElementById('exp') as HTMLButtonElement;
    const rescanBtn = shadow.getElementById('rescan') as HTMLElement;
    const boxEl = shadow.querySelector('.box') as HTMLElement;

    /**
     * 面板显示/收起只切显示，不删节点 ——
     * 识别到的账号、手填的名字全部保留，重新展开就是原样。
     * 收起靠面板右上角的 ✕；重新展开靠工具栏弹窗里的按钮或刷新页面。
     */
    function isPanelVisible(): boolean {
      return boxEl.style.display !== 'none';
    }

    function setPanelVisible(visible: boolean): void {
      boxEl.style.display = visible ? 'block' : 'none';
    }

    /** 收起 / 展开来回切；重新展开时顺手重扫一次（SPA 切页后数据可能变了） */
    function togglePanel(): boolean {
      const next = !isPanelVisible();
      setPanelVisible(next);
      if (next) scan();
      return next;
    }

    // 注意：detectedUsername / detectedSource / detectedStrong
    // 都是**模块级**变量（见 buildUsernameReport 附近），这里刻意不重复声明 ——
    // 否则局部变量会遮蔽它们，诊断报告里永远是空值。
    let headerTableRef: HTMLTableElement | null = null;
    let bodyTableRef: HTMLTableElement | null = null;
    let headerMap: ScanResult | null = null;
    /** 用户是否手动改过用户名（改过就不再被自动识别结果覆盖） */
    let usernameEdited = false;
    /** 用户是否已手动操作（导出/重新扫描）；已操作则不再自动重试覆盖提示 */
    let userActed = false;

    // ✕ 收起面板（面板不会消失，可从工具栏弹窗重新展开，或刷新页面）
    shadow.querySelector('.hd .x')?.addEventListener('click', () => setPanelVisible(false));

    // 重新展开入口：工具栏弹窗里的按钮走这条逻辑（需要 sendResponse 回包，
    // 因为 Chrome 不会把监听器的普通返回值当作响应；背景脚本也是这个写法）。
    browser.runtime.onMessage.addListener((msg: unknown, _sender, sendResponse) => {
      const type = (msg as { type?: string } | null)?.type;
      if (type === 'DXM_PANEL_STATE') {
        sendResponse({ ok: true, visible: isPanelVisible() });
        return true;
      }
      if (type !== 'DXM_TOGGLE_PANEL') return false;
      sendResponse({ ok: true, visible: togglePanel() });
      return true;
    });

    /** 当前生效的用户名：手填优先，其次自动识别 */
    function currentUsername(): string {
      return (unInput.value || detectedUsername).trim();
    }

    /**
     * 这个账号名是否还需要用户确认一次。
     *
     * 只有三种情况可以放心直接用：
     *   1) 用户亲手填/确认过（usernameEdited）
     *   2) 命中了已知的账号名 DOM 结构（detectedStrong）
     *   3) 是上次导出时记下来的（用户此前已确认过）
     * 其余（从 storage/cookie/顶部文字**猜**出来的）一律要确认 —— 宁可多点一下，
     * 也不要导出一个文件名莫名其妙的 Excel。
     */
    function needsNameConfirm(): boolean {
      if (usernameEdited && unInput.value.trim()) return false;
      if (detectedStrong) return false;
      if (detectedFromMemory) return false;
      return true;
    }

    /** 把识别来源变成人话（用在提示行 / 悬浮说明里，方便反馈问题时截图说明） */
    function sourceLabel(src: string): string {
      const i = src.indexOf(':');
      const kind = i < 0 ? src : src.slice(0, i);
      const rest = i < 0 ? '' : src.slice(i + 1);
      if (kind === 'selector') return `页面账号区 ${rest.replace(/@(top|iframe#\d+)$/, '')}`;
      if (kind === 'heuristic') return '页面顶部文字（推测）';
      if (kind === 'cookie') return `Cookie ${rest}`;
      if (/localStorage|sessionStorage/.test(kind)) return `浏览器存储 ${rest}`;
      return src;
    }

    /** 刷新「账号名」与「将导出为 xxx.xlsx」的实时预览 */
    function updateFilePreview(): void {
      const name = currentUsername();
      unameEl.textContent = name || '未识别到账号';
      unameEl.classList.toggle('none', !name);

      fnameEl.textContent = '';
      if (!name) {
        fnameEl.classList.add('err');
        fnameEl.textContent = '未识别到账号名，导出时会请你填一次（之后自动记住）';
        return;
      }
      fnameEl.classList.remove('err');
      fnameEl.appendChild(document.createTextNode('将导出为：'));
      const b = document.createElement('b');
      b.textContent = buildFileName(name);
      fnameEl.appendChild(b);
    }

    /** 自动识别登录账号（用户手动改过就不覆盖） */
    function refreshUsername(): void {
      // detectUsername 内部已逐条 try/catch，但这里再包一层：
      // 万一它整体抛错，UI 也不能停在「识别中…」这个中间态上——那样用户只会看到
      // 一个永远不变的占位符，误以为「里面的文本不对」。
      try {
        const hit = detectUsername();
        if (hit.name) {
          // 高可信命中随时可以覆盖；低可信命中不得覆盖已确认的高可信结果。
          // （账号区异步渲染：早期只能靠 storage 兜底，晚期拿到真结构必须能纠正过来）
          if (hit.strong || !detectedStrong) {
            detectedUsername = hit.name;
            detectedSource = hit.source;
            detectedStrong = hit.strong;
            // 这是本次页面探测的实时结果，不再是「记住的值」
            detectedFromMemory = false;
          }
        }
        unameEl.title = detectedUsername
          ? `识别来源：${detectedSource}${detectedStrong ? '' : '（低可信）'}`
          : '没读到账号名，可点右侧「修改」手动填';
      } catch (e) {
        console.warn('[店小秘导出] 账号识别异常', e);
        unameEl.title = '识别过程出错，可点右侧「修改」手动填';
      }
      updateFilePreview();
    }

    /**
     * 轮询用尽后的收尾。
     *
     * - 完全没结果 → 展开输入框，请用户填一次（填完会记住）
     * - 只有低可信结果（猜的）→ 保留但明确标黄警告，让用户核对/改正。
     *   绝不能默默把一个猜测值当账号名用掉 —— 那会导出文件名莫名其妙的 Excel。
     */
    function onDetectExhausted(): void {
      if (!currentUsername()) {
        // 面板上已不放诊断入口：识别彻底失败时，把完整报告打进控制台（F12）便于排查
        console.warn('[店小秘导出] 账号识别失败，诊断报告 =', buildUsernameReport());
        unInput.style.display = 'block';
        ueditEl.textContent = '完成';
        unameEl.classList.add('none');
        unameEl.title = '自动识别失败，请在下方手填一次（之后会记住）';
        updateFilePreview();
        tip.textContent = '没读到页面账号区。已展开输入框，填一次账号名即可（之后会记住）。';
        tip.classList.add('err');
        return;
      }
      if (!detectedStrong && !usernameEdited) {
        console.warn('[店小秘导出] 仅拿到低可信账号名 =', detectedUsername, '来源 =', detectedSource);
        tip.textContent = `账号名是从「${sourceLabel(detectedSource)}」猜的，不一定是登录账号。请核对上方，不对就点「修改」手填。`;
        tip.classList.add('err');
      }
    }

    // 「修改」：识别不到（或想换名字）时的兜底，点开才出现输入框
    ueditEl.addEventListener('click', () => {
      const editing = unInput.style.display !== 'none';
      unInput.style.display = editing ? 'none' : 'block';
      ueditEl.textContent = editing ? '修改' : '完成';
      if (!editing) {
        unInput.focus();
        unInput.select();
      }
      usernameEdited = !!unInput.value.trim();
      updateFilePreview();
    });
    unInput.addEventListener('input', () => {
      usernameEdited = !!unInput.value.trim();
      updateFilePreview();
    });

    refreshUsername();
    // 页面探测不到账号时，用上次记住的用户名兜底（用户手动改过则不接管）
    loadSavedUsername().then((saved) => {
      if (saved && !detectedUsername && !usernameEdited) {
        detectedUsername = saved;
        detectedSource = '上次导出时记住的账号';
        detectedStrong = false;
        detectedFromMemory = true;
        unameEl.title = '来源：上次导出时记住的账号';
        updateFilePreview();
      }
    });

    // 顶部账号区是 Vue 异步渲染的（content script 注入时可能还没出来），
    // 所以未识别到就轻量轮询。
    //
    // 停止条件只是「高可信命中」或「用户手填」——**低可信命中不能停**。
    // 真实踩坑：早期 DOM 里没有账号区，降级到 storage 捞出了会话 ID
    // `952C1DB87F4EDF20` 并当成用户名，轮询随即停止；等真正的 `.user-name`
    // 渲染出来时已经不再检查，于是面板永远显示那个 ID。
    let usernameRetries = 0;
    const usernameTimer = setInterval(() => {
      if (detectedStrong || usernameEdited) {
        clearInterval(usernameTimer);
        return;
      }
      if (++usernameRetries > 15) {
        clearInterval(usernameTimer);
        // 轮询用尽仍没拿到高可信结果 → 该手填就手填，不要把面板停在中间态
        onDetectExhausted();
        return;
      }
      refreshUsername();
    }, 1200);

    // 站点脚本若替换了 body（SPA 极端情况），把面板重新挂上，避免面板“凭空消失”
    const keepAlive = new MutationObserver(() => {
      if (!panel.isConnected) (document.body ?? document.documentElement).appendChild(panel);
    });
    keepAlive.observe(document.documentElement, { childList: true });

    function scan() {
      try {
        const { headerTable, bodyTable } = findTables();
        if (!headerTable || !bodyTable) {
          headerTableRef = null;
          bodyTableRef = null;
          headerMap = null;
          expBtn.disabled = true;
          tip.classList.add('err');
          tip.textContent = '未检测到列表表格（表头/数据表未能同时定位）。请确认在「发货成功列表」页面，点「重新扫描」或刷新后重试。';
          return;
        }
        headerTableRef = headerTable;
        bodyTableRef = bodyTable;
        headerMap = buildHeaderMap(headerTable);
        expBtn.disabled = false;

        const rows = readRows(bodyTable, headerMap);
        console.log(
          '[店小秘导出] 扫描到数据行 =',
          rows.length,
          '；数据表 tbody 总行数 =',
          bodyTable.querySelectorAll('tbody tr').length,
          '；表头 =',
          headerMap.headers,
        );
        if (rows.length === 0) {
          tip.classList.add('err');
          tip.textContent = `已找到表格，但没读到数据行（共 ${rows.length} 行）。可能是列表还没加载完，点「重新扫描」或刷新。`;
          console.log('[店小秘导出] 表头 =', headerMap.headers);
          return;
        }
        const foundGoods = !!pickCell(rows[0]!, CONFIG.sourceColumns['商品信息']);
        const foundLogi = !!pickCell(rows[0]!, CONFIG.sourceColumns['物流方式']);
        const foundTime = !!pickCell(rows[0]!, CONFIG.sourceColumns['时间']);
        const missing: string[] = [];
        if (!foundGoods) missing.push('商品信息');
        if (!foundLogi) missing.push('物流方式');
        if (!foundTime) missing.push('时间');
        if (missing.length) {
          tip.classList.add('err');
          tip.textContent = `已读 ${rows.length} 行，但未匹配到列：${missing.join('、')}。表头：${headerMap.headers.join(' / ')}`;
        } else {
          const pending = countPendingImages(bodyTable);
          tip.classList.remove('err');
          tip.textContent =
            `已就绪，读到 ${rows.length} 个订单。` +
            (pending
              ? `还有 ${pending} 张商品图未就绪（点「导出」会自动滚动列表并等它们加载完）。`
              : '') +
            '「材质」列本页无数据，统一填「水洗底」。';
        }
      } catch (e) {
        tip.classList.add('err');
        tip.textContent = '扫描出错：' + (e instanceof Error ? e.message : String(e));
      }
    }

    expBtn.addEventListener('click', async () => {
      // 用户已手动操作，停掉自动重试扫描，避免它把导出结果提示覆盖掉
      userActed = true;
      if (!headerTableRef || !bodyTableRef || !headerMap) {
        tip.classList.add('err');
        tip.textContent = '请先点「重新扫描」定位列表表格。';
        return;
      }
      // 防重复点击（下载图片期间按钮禁用）
      expBtn.disabled = true;
      const restore = () => {
        expBtn.disabled = false;
      };
      try {
        // 先读一次行：确认有数据、并留出「账号名待确认」的中止机会
        let rows = readRows(bodyTableRef, headerMap).flatMap(toRecords);
        if (rows.length === 0) {
          restore();
          tip.classList.add('err');
          tip.textContent = '没有可导出的数据行。';
          return;
        }
        tip.classList.remove('err');
        // 账号区域可能异步渲染，导出前再补一次识别（含同源 iframe）
        if (!currentUsername() || needsNameConfirm()) refreshUsername();
        if (needsNameConfirm()) {
          // 中止本次导出，让用户确认/填写一次：
          //  - 完全没识别到 → 空框让他填
          //  - 只是**猜**到的值（storage/cookie/顶部文字）→ 预填进去让他核对，
          //    因为猜出来的值很容易是会话 ID 之类的东西（真实踩坑 `952C1DB87F4EDF20`）
          // 两种情况都要点「完成」确认，之后会记住。
          restore();
          unInput.style.display = 'block';
          ueditEl.textContent = '完成';
          if (!unInput.value && detectedUsername) unInput.value = detectedUsername;
          unInput.focus();
          unInput.select();
          tip.classList.add('err');
          tip.textContent = currentUsername()
            ? `账号名是从「${sourceLabel(detectedSource)}」推断的，未必是登录账号。请核对上方，无误就点「完成」再导出；不对就直接改成正确账号名。`
            : '未能自动识别账号名。已在上方展开输入框，填一次即可（会记住）；填完再点「导出Excel 和面单」。';
          console.log('[店小秘导出] 账号名待确认，诊断报告 =', buildUsernameReport());
          return;
        }

        // ---- 面单（必须放在最前面）----
        // 「开始打印」会 window.open 一个 PDF，而 window.open 要求用户点按钮后 ~5s 内的
        // 激活态；滚动列表 + 下载几十张商品图会把这个窗口耗光，放后面就会被浏览器拦掉。
        // 面单失败不阻断 Excel：继续导，只是最后的提示会标红并写明原因。
        let labelNote = '';
        let labelFailed = false;
        tip.textContent = '正在生成面单（当页全部订单）…';
        try {
          const lr = await downloadLabels(bodyTableRef, currentUsername());
          if (lr.ok) {
            labelNote = `；面单 ${lr.file}`;
            console.log('[店小秘导出] 面单已下载:', lr.file);
          } else {
            labelFailed = true;
            labelNote = `；面单下载失败：${lr.error}`;
            console.warn('[店小秘导出] 面单下载失败:', lr.error);
          }
        } catch (e) {
          labelFailed = true;
          const msg = e instanceof Error ? e.message : String(e);
          labelNote = `；面单下载失败：${msg}`;
          console.warn('[店小秘导出] 面单下载异常:', msg);
        }

        // 懒加载图片：没进过视口的行，DOM 里只有 1×1 占位图、真实地址根本不在 ——
        // 不滚一遍就导出，那些行必然没有图片。这里自动滚一遍（最后会滚回原位）再重读。
        const pendingBefore = countPendingImages(bodyTableRef);
        if (pendingBefore > 0) {
          tip.textContent = `有 ${pendingBefore} 张商品图未就绪，正在滚动列表并等待加载…`;
          const left = await ensureImagesLoaded(bodyTableRef);
          if (left > 0) {
            console.warn(
              `[店小秘导出] 滚动后仍有 ${left} 张图未加载（页面自身也没加载出来），对应行导出后没有图片`,
            );
          }
          // 图片 src 变了，必须重新读一次行（防页面刚好在重渲染，读空就还用原来这份）
          const refreshed = readRows(bodyTableRef, headerMap).flatMap(toRecords);
          if (refreshed.length > 0) rows = refreshed;
        }

        const filename = buildFileName(currentUsername());
        const imgCount = new Set(rows.map((r) => r.图片).filter(Boolean)).size;
        tip.textContent = `导出 ${rows.length} 行，正在下载图片 0/${imgCount}…`;

        const st = await exportToExcel(rows, filename, (done, total) => {
          tip.textContent = `导出 ${rows.length} 行，正在下载高清图 ${done}/${total}…`;
        });

        console.log('[店小秘导出] 导出完成，文件名 =', filename, '；账号 =', currentUsername(), '；统计 =', st);
        const parts = [`${st.rows} 行`];
        if (st.withImage) parts.push(`${st.withImage} 行含图`);
        if (st.noUrl) parts.push(`${st.noUrl} 行取不到图片地址`);
        if (st.failed) parts.push(`${st.failed} 张图片下载失败`);
        const partial = st.noUrl > 0 || st.failed > 0 || labelFailed;
        tip.classList.toggle('err', partial);
        tip.textContent =
          `导出成功：${filename}（${parts.join('，')}）${labelNote}` +
          (partial ? '。详情按 F12 看控制台' : '');
        // 记住这次用的用户名，下次页面探测不到也能直接用
        void saveUsername(currentUsername());
        restore();
      } catch (e) {
        restore();
        tip.classList.add('err');
        tip.textContent = '导出失败：' + (e instanceof Error ? e.message : String(e));
      }
    });

    rescanBtn.addEventListener('click', () => {
      userActed = true;
      // SPA 切页后账号区可能变了，重新识别一次（用户手改过则不覆盖）
      if (!usernameEdited) refreshUsername();
      scan();
    });

    // 账号识别排查：面板上不再放入口，失败时把完整报告打进控制台（见 onDetectExhausted）
    // 和「账号名待确认」分支。需要现场排查就开 F12 看 `[店小秘导出] 账号诊断报告 =`。

    scan();
    // 应对 SPA 延迟渲染：短暂自动重试，但用户一旦操作就停止
    const autoScan = () => {
      if (!userActed) scan();
    };
    setTimeout(autoScan, 1500);
    setTimeout(autoScan, 4000);
  },
});
