/**
 * 店小秘发货成功列表导出插件 —— Content Script
 *
 * 功能：在 https://www.dianxiaomi.com/ 的「发货成功列表」页面注入一个悬浮面板，
 *      用户选择日期范围（当天 / 近两天 / 近三天）后，将列表数据按固定字段
 *      （图片、尺寸、件数、材质、运单号）导出为 Excel，文件名取当前登录账号用户名。
 *
 * 数据来源（店小秘 vxe-table 真实结构）：
 *   - 图片 / 尺寸 / 件数  ← 「商品信息」列（图片取 <img>，尺寸取 size，件数取 x N）
 *   - 运单号              ← 「物流方式」列里的单号（如 「YT2626700708781963」）
 *   - 材质                ← 店小秘发货成功列表页【无此字段】，留空（见下方说明）
 *   - 日期筛选            ← 「时间」列里的「发货：YYYY-MM-DD HH:mm」
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
  targetColumns: ['图片', '尺寸', '件数', '材质', '运单号'] as const,

  /**
   * 店小秘发货成功列表真实列标题 -> 关键词（命中其一即用该列）。
   * 这些列是解析 5 个目标字段的数据来源。
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
 *   4) 页面原始 URL（最终兜底，即使前面都 404 也不至于整格空白）
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
    out.push(`${base}._SL${CONFIG.hdSize}_.${ext}${qs}`);
    out.push(`${base}.${ext}${qs}`);
    out.push(`${base}._SL500_.${ext}${qs}`);
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

/**
 * 从文本解析发货日期为 Date；失败返回 null。
 * 兼容以下格式（店小秘线上常见写法都覆盖）：
 *   - 2026-09-24 18:00 / 2026-09-24 18:00:00
 *   - 2026/09/24 18:00（斜杠分隔）
 *   - 09-24 18:00（缺年份，按当前年补齐）
 *   - 今天 18:00 / 昨天 / 前天（相对日期，按当前日期回推）
 *   - 发货：2026-09-24 18:00（带前缀）
 */
function parseDate(raw: string): Date | null {
  if (!raw) return null;
  const s = raw.trim();
  if (!s) return null;
  const now = new Date();

  // 1) 相对日期：今天 / 昨天 / 前天（含「昨日」写法）
  const relMap: Record<string, number> = { 今天: 0, 今日: 0, 昨天: 1, 昨日: 1, 前天: 2 };
  for (const kw of Object.keys(relMap)) {
    if (s.includes(kw)) {
      const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() - relMap[kw]!);
      const tm = s.match(/(\d{1,2}):(\d{2})/);
      if (tm) d.setHours(Number(tm[1]), Number(tm[2]), 0, 0);
      else d.setHours(0, 0, 0, 0);
      if (!Number.isNaN(d.getTime())) return d;
    }
  }

  // 2) 完整日期 2026-09-24 / 2026/09/24（+ 可选时间）
  const full = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (full) {
    const d = new Date(
      Number(full[1]),
      Number(full[2]) - 1,
      Number(full[3]),
      full[4] ? Number(full[4]) : 0,
      full[5] ? Number(full[5]) : 0,
    );
    if (!Number.isNaN(d.getTime())) return d;
  }

  // 3) 缺年份：09-24 18:00 / 9/24 18:00（按当前年补齐）
  const part = s.match(/(\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (part) {
    const d = new Date(
      now.getFullYear(),
      Number(part[1]) - 1,
      Number(part[2]),
      part[3] ? Number(part[3]) : 0,
      part[4] ? Number(part[4]) : 0,
    );
    if (!Number.isNaN(d.getTime())) return d;
  }

  // 4) 兜底：浏览器原生解析
  const t = Date.parse(s);
  if (!Number.isNaN(t)) return new Date(t);
  return null;
}

/**
 * 从「时间」单元格提取【发货】时间文本。
 * 该单元格实际包含 6 个时间：下单 / 审核 / 申请 / 提交 / 发货 / 送达，
 * 例如：下单： 2026-09-24 06:09 审核： ... 发货： 2026-09-24 18:00 送达： 2026-10-21 14:59
 * 必须专门取「发货：」后面的那个，否则会误用「下单」日期（两者可能差一天）。
 */
function extractShipDate(text: string): string {
  if (!text) return '';
  const ship = text.match(/发货\s*[：:]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}:\d{2})?)/);
  if (ship) return ship[1] ?? '';
  // 兜底：没有「发货：」时取第一个日期
  const any = text.match(/\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:\s+\d{1,2}:\d{2})?/);
  return any ? any[0] : '';
}

/** 本地日期（当天 00:00:00） */
function startOfDay(d = new Date()): Date {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

/** 根据范围返回起始日期（含） */
function rangeStart(range: 'today' | '2d' | '3d'): Date {
  const s = startOfDay();
  if (range === '2d') s.setDate(s.getDate() - 1);
  else if (range === '3d') s.setDate(s.getDate() - 2);
  return s;
}

function rangeLabel(range: 'today' | '2d' | '3d'): string {
  return range === 'today' ? '当天' : range === '2d' ? '近两天' : '近三天';
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
 * 一个订单格可能含【多个商品】，真实文本形如：
 *   ! F8-YKAY-QP5Z x 1 USD 59.99 color ：... size ：4' x 5' ! M9-MGO2-R30H x 2 USD 24.99 color ：... size ：2' x 3'
 * 每个商品各有独立图片/尺寸/件数，因此按「!」拆成多段，一段=一个商品=导出一行。
 *
 * 图片为懒加载：未加载时 src 是 1x1 占位 base64（data:image/gif...），
 * 这类必须过滤掉，否则导出的是一堆占位图。
 */
function parseProducts(cell: Element | undefined): Product[] {
  const text = (cell?.textContent ?? '').replace(/\s+/g, ' ').trim();
  // 收集真实图片（过滤懒加载占位图）
  const imgs = Array.from(cell?.querySelectorAll('img') ?? [])
    .map((im) => im.getAttribute('src') || im.getAttribute('data-src') || '')
    .filter((s) => s && !s.startsWith('data:'));

  // 按商品拆分（每段以 SKU 开头，段间是渲染出的 "!"）
  const segs = text
    .split(/\s*!\s*/)
    .map((s) => s.trim())
    .filter(Boolean);
  const list = segs.length > 0 ? segs : [text];

  return list.map((seg, i) => ({
    图片: imgs[i] ? absUrl(imgs[i]!) : (imgs[0] ? absUrl(imgs[0]!) : ''),
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
 * 一个订单可能含多个商品 → 每个商品导出一行（共享同一运单号与发货时间）。
 */
function toRecords(
  rowMap: Map<string, Element>,
): Array<Record<string, string> & { __date: string; __dateRaw: string }> {
  const goodsCell = pickCell(rowMap, CONFIG.sourceColumns['商品信息']);
  const logisticsCell = pickCell(rowMap, CONFIG.sourceColumns['物流方式']);
  const timeCell = pickCell(rowMap, CONFIG.sourceColumns['时间']);

  const products = parseProducts(goodsCell);
  const 运单号 = parseTracking((logisticsCell?.textContent ?? '').trim());

  // 解析发货日期：优先取「时间」列里的【发货】时间，取不到则退化到整行文本找日期
  const timeText = (timeCell?.textContent ?? '').trim();
  let dateObj = parseDate(extractShipDate(timeText));
  let dateRaw = timeText;
  if (!dateObj) {
    const whole = Array.from(rowMap.values())
      .map((e) => (e.textContent ?? '').trim())
      .join(' ');
    dateObj = parseDate(extractShipDate(whole));
    if (dateObj) dateRaw = whole;
  }
  const __date = dateObj ? dateObj.toISOString() : '';

  return products.map((p) => ({
    __date,
    __dateRaw: dateRaw,
    图片: p.图片,
    尺寸: p.尺寸,
    件数: p.件数,
    材质: '', // 店小秘发货成功列表页无「材质」字段，留空
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

/**
 * 经 background 下载图片，返回 dataURL；失败返回 null。
 * （content script 直接 fetch 会被 CORS 拦，必须走 background 的主机权限）
 *
 * 会一并把「高清候选链」带过去，由 background 依次尝试，取第一个真正抓到的。
 */
async function fetchImageViaBackground(url: string): Promise<string | null> {
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
      return `data:${resp.mime || 'image/jpeg'};base64,${resp.base64}`;
    }
    console.warn('[店小秘导出] 图片下载失败:', url, resp?.error);
    return null;
  } catch (e) {
    console.warn('[店小秘导出] 图片下载异常:', url, e);
    return null;
  }
}

/**
 * 把图片等比缩到最长边 IMG_MAX_PX，并统一转成【白底 JPEG】。
 *
 * 为什么是 JPEG 而不是 PNG：
 *   Excel 只认得 png/jpeg/gif。转 PNG 是无损的，但一张 900px 商品图 PNG 动辄
 *   500KB~1MB，几十行就把 xlsx 撑到几十 MB；同样清晰度的 JPEG 只要 100KB 左右。
 *   商品图基本是白底照片，先铺白再画，避免透明区域被压成黑块。
 */
function toJpegDataUrl(dataUrl: string): Promise<string> {
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
        if (!ctx) return resolve(dataUrl);
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, c.width, c.height);
        ctx.drawImage(img, 0, 0, c.width, c.height);
        // dataURL 是同源的，canvas 不会被污染
        return resolve(c.toDataURL('image/jpeg', CONFIG.imgQuality));
      } catch {
        return resolve(dataUrl);
      }
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

/**
 * 导出 Excel：表头 + 数据行，并把商品图【嵌入】到「图片」列。
 *
 * filename 由调用方算好（用户名 + 日期范围），这里不再关心命名规则。
 */
async function exportToExcel(
  rows: Array<Record<string, string> & { __date: string; __dateRaw: string }>,
  filename: string,
  onProgress?: (done: number, total: number) => void,
): Promise<void> {
  // ---- 1) 先把用到的图片全部下载下来（同一 URL 只下一次） ----
  const urls = Array.from(new Set(rows.map((r) => r.图片).filter((u) => !!u)));
  const imgCache = new Map<string, string | null>();
  for (let i = 0; i < urls.length; i++) {
    onProgress?.(i + 1, urls.length);
    const raw = await fetchImageViaBackground(urls[i]!);
    imgCache.set(urls[i]!, raw ? await toJpegDataUrl(raw) : null);
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
  rows.forEach((r, i) => {
    const rowNo = i + 2; // 1-based 行号（表头占第 1 行）
    ws.addRow(['', r.尺寸, r.件数, r.材质, r.运单号]);
    ws.getRow(rowNo).height = CONFIG.imgRowHeight;

    const dataUrl = r.图片 ? imgCache.get(r.图片) : null;
    if (dataUrl) {
      const b64 = dataUrl.slice(dataUrl.indexOf(',') + 1);
      try {
        // ExcelJS 的类型声明写的是 Node 的 Buffer，浏览器环境实际传 Uint8Array，故整体断言
        const imageId = wb.addImage({
          buffer: base64ToBytes(b64),
          extension: 'jpeg',
        } as unknown as Parameters<typeof wb.addImage>[0]);
        ws.addImage(imageId, {
          tl: { col: 0, row: rowNo - 1 }, // tl 是 0-based
          ext: { width: CONFIG.imgPx, height: CONFIG.imgPx },
          editAs: 'oneCell',
        });
      } catch (e) {
        console.warn('[店小秘导出] 插图失败:', r.图片, e);
      }
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
}

/** 数据范围文本：当天 `20260925`，跨天 `20260923-20260925`（仅用于面板展示） */
function dateRangeText(range: 'today' | '2d' | '3d'): string {
  const end = startOfDay();
  const start = rangeStart(range);
  return ymd(start) === ymd(end) ? ymd(end) : `${ymd(start)}-${ymd(end)}`;
}

/**
 * 生成导出文件名：`用户名_导出当天年月日.xlsx`，例如 `jia_yangdong_20260925.xlsx`。
 *
 * 日期固定取【执行导出的当天】，不随筛选范围变化 ——
 * 文件名标记的是「这批数据是什么时候导的」，数据范围在面板上单独显示。
 * 非法文件名字符（\ / : * ? " < > |）替换成下划线。
 */
function buildFileName(username: string): string {
  const base = (username || '店小秘').replace(/[\\/:*?"<>|]/g, '_').trim() || '店小秘';
  return `${base}_${ymd(startOfDay())}.xlsx`;
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
    .opts { display: flex; gap: 6px; }
    .opt { flex: 1; text-align: center; padding: 6px 0; border: 1px solid #d8d8d8; border-radius: 6px;
           cursor: pointer; user-select: none; }
    .opt.on { border-color: #2b6cff; background: #eef3ff; color: #2b6cff; font-weight: 600; }
    input.un { width: 100%; padding: 6px 8px; border: 1px solid #d8d8d8; border-radius: 6px; font-size: 12px; }
    .unbox { display: flex; align-items: center; gap: 6px; }
    .uname { flex: 1; min-width: 0; padding: 5px 8px; border: 1px dashed #dfe3ea; border-radius: 6px;
             background: #f7f9fc; color: #2b3a55; font-weight: 600;
             overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .uname.none { color: #e54545; font-weight: 400; border-style: dashed; }
    .uedit { flex: none; font-size: 11px; color: #2b6cff; cursor: pointer; text-decoration: underline; }
    .fname { margin-top: 6px; font-size: 11px; color: #5b6b85; word-break: break-all; }
    .fname b { color: #2b3a55; }
    .fname i { font-style: normal; color: #8b98ad; }
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
        <span class="lbl">日期范围（按发货时间）</span>
        <div class="opts" id="opts">
          <div class="opt on" data-r="today">当天</div>
          <div class="opt" data-r="2d">近两天</div>
          <div class="opt" data-r="3d">近三天</div>
        </div>
      </div>
      <div class="row">
        <span class="lbl">登录账号（自动识别）</span>
        <div class="unbox">
          <span class="uname" id="uname">识别中…</span>
          <span class="uedit" id="uedit" title="识别不到时可手动改">修改</span>
        </div>
        <input class="un" id="un" placeholder="请输入店小秘账号名" style="display:none" />
        <div class="fname" id="fname"></div>
      </div>
      <button class="btn" id="exp">导出 Excel</button>
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
    const opts = shadow.getElementById('opts') as HTMLElement;
    const rescanBtn = shadow.getElementById('rescan') as HTMLElement;
    const boxEl = shadow.querySelector('.box') as HTMLElement;

    /**
     * 面板显示/收起只切显示，不删节点 ——
     * 选的日期范围、识别到的账号、手填的名字全部保留，重新展开就是原样。
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

    let currentRange: 'today' | '2d' | '3d' = 'today';
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
      // 日期固定是「导出当天」，所以数据范围不会再体现在文件名里，这里单独标出来
      const r = document.createElement('i');
      r.textContent = `（数据范围：${rangeLabel(currentRange)}）`;
      fnameEl.appendChild(r);
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

    opts.querySelectorAll('.opt').forEach((el) => {
      el.addEventListener('click', () => {
        opts.querySelectorAll('.opt').forEach((o) => o.classList.remove('on'));
        el.classList.add('on');
        currentRange = (el.getAttribute('data-r') as 'today' | '2d' | '3d') ?? 'today';
        updateFilePreview();
      });
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
          tip.classList.remove('err');
          tip.textContent = `已就绪，读到 ${rows.length} 个订单。导出前请上下滚动列表让商品图加载完；「材质」列本页无数据，留空。`;
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
        const rowMaps = readRows(bodyTableRef, headerMap);
        // 多商品订单会拆成多行（每个商品一行）
        const rows = rowMaps.flatMap(toRecords);
        if (rows.length === 0) {
          restore();
          tip.classList.add('err');
          tip.textContent = '没有可导出的数据行。';
          return;
        }
        // 按发货时间筛选
        const start = rangeStart(currentRange);
        const filtered = rows.filter((r) => {
          const d = parseDate(r.__date);
          return d != null && d >= start;
        });
        const dropped = rows.length - filtered.length;
        if (filtered.length === 0) {
          restore();
          tip.classList.add('err');
          // 显示「时间」列原始文本（诊断真实格式），而不是只显示解析结果
          const rawSamples = rows
            .slice(0, 3)
            .map((r) => `"${(r.__dateRaw || '').slice(0, 40) || '(空)'}"`)
            .join('、');
          console.log('[店小秘导出] 前3行 时间列原文 =', rows.slice(0, 3).map((r) => r.__dateRaw));
          console.log(
            '[店小秘导出] 第1行整行文本 =',
            Array.from(rowMaps[0]?.values() ?? []).map((e) => (e.textContent ?? '').trim().slice(0, 30)),
          );
          tip.textContent = `读到 ${rows.length} 行，但按「${rangeLabel(currentRange)}」过滤后为 0 行。时间列原文：${rawSamples}`;
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
            : '未能自动识别账号名。已在上方展开输入框，填一次即可（会记住）；填完再点「导出 Excel」。';
          console.log('[店小秘导出] 账号名待确认，诊断报告 =', buildUsernameReport());
          return;
        }
        const filename = buildFileName(currentUsername());
        const imgCount = new Set(filtered.map((r) => r.图片).filter(Boolean)).size;
        tip.textContent = `命中「${rangeLabel(currentRange)}」${filtered.length} 行${dropped ? `（已过滤 ${dropped} 行）` : ''}，正在下载图片 0/${imgCount}…`;

        await exportToExcel(filtered, filename, (done, total) => {
          tip.textContent = `命中 ${filtered.length} 行，正在下载高清图 ${done}/${total}…`;
        });

        console.log('[店小秘导出] 导出完成，文件名 =', filename, '；账号 =', currentUsername());
        tip.textContent = `导出成功：${filename}（${filtered.length} 行，含 ${imgCount} 张高清图）`;
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
