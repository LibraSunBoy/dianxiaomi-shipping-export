/**
 * Background Service Worker
 *
 * 职责：替 content script 下载商品图片。
 *
 * 为什么必须在 background 下载：
 *   商品图在 Amazon CDN 等跨域地址上，content script 直接 fetch 会被 CORS 拦截；
 *   而 background 拥有 host_permissions（见 wxt.config.ts），可以跨域抓取。
 *   页面里已加载的 <img> 也没有 crossorigin 属性，canvas 会被污染，无法直接取像素。
 */

interface FetchImageResponse {
  ok: boolean;
  base64?: string;
  mime?: string;
  error?: string;
  /** 实际成功抓到的那条 URL（便于确认高清替换是否生效） */
  url?: string;
}

/** 小于这个字节数视为无效图（占位/空白图），继续试下一个候选 */
const MIN_IMAGE_BYTES = 512;

/** 把 ArrayBuffer 转 base64（service worker 里没有 FileReader，用 btoa 分块处理） */
function arrayBufferToBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  let bin = '';
  const CHUNK = 0x8000; // 每次处理 32KB，避免 apply 参数过多爆栈
  for (let i = 0; i < bytes.length; i += CHUNK) {
    let s = '';
    const end = Math.min(i + CHUNK, bytes.length);
    for (let j = i; j < end; j++) s += String.fromCharCode(bytes[j]!);
    bin += s;
  }
  return btoa(bin);
}

/** 抓一张图；非图片类型、或小到不正常的都算失败（CDN 常对缺图返回 200 + 占位） */
async function fetchImageOnce(url: string): Promise<FetchImageResponse> {
  try {
    const res = await fetch(url, { credentials: 'omit' });
    if (!res.ok) return { ok: false, error: `HTTP ${res.status}` };
    const mime = res.headers.get('content-type') || 'image/jpeg';
    if (!/^image\//i.test(mime)) return { ok: false, error: `not image (${mime})` };
    const buf = await res.arrayBuffer();
    if (buf.byteLength < MIN_IMAGE_BYTES) {
      return { ok: false, error: `too small (${buf.byteLength}B)` };
    }
    return { ok: true, base64: arrayBufferToBase64(buf), mime, url };
  } catch (e) {
    return { ok: false, error: e instanceof Error ? e.message : String(e) };
  }
}

/**
 * 依次尝试候选 URL，返回第一个真正抓到图的。
 *
 * 之所以要「依次试」而不是直接改 URL：把 Amazon 的 `._SL100_` 去掉换成 `._SL1000_`
 * 在绝大多数图上是有效的，但只要有一张图的原始文件名形态特殊，硬替换就会 404，
 * 导致这张图在 Excel 里整格空白。有了回退链，最差也只是退回到页面上的缩略图。
 */
async function fetchImageBest(url: string, candidates?: string[]): Promise<FetchImageResponse> {
  const list = [...new Set([...(Array.isArray(candidates) ? candidates : []), url].filter(Boolean))];
  const errors: string[] = [];
  for (const u of list) {
    const r = await fetchImageOnce(u);
    if (r.ok) return r;
    errors.push(`${u} -> ${r.error}`);
  }
  return { ok: false, error: errors.join(' | ') };
}

// ============================ 面单 PDF ============================

interface PdfResponse {
  ok: boolean;
  url?: string;
  base64?: string;
  mime?: string;
  error?: string;
}

/** 是不是一个会展示/下载 PDF 的地址（面单就是 `https://print.dianxiaomi.com/2026-09-30/<uuid>.pdf`） */
function isPdfUrl(u: string | null | undefined): boolean {
  return !!u && /^https?:/i.test(u) && /\.pdf(\?|#|$)/i.test(u);
}

async function fetchPdf(url: string): Promise<PdfResponse> {
  try {
    // 带 cookie：print.dianxiaomi.com 虽然是 uuid 直链，但万一有会话校验也不会挂
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) return { ok: false, url, error: `HTTP ${res.status}` };
    const mime = res.headers.get('content-type') || 'application/pdf';
    const buf = await res.arrayBuffer();
    if (buf.byteLength < 64) return { ok: false, url, error: `too small (${buf.byteLength}B)` };
    return { ok: true, url, mime, base64: arrayBufferToBase64(buf) };
  } catch (e) {
    return { ok: false, url, error: e instanceof Error ? e.message : String(e) };
  }
}

/**
 * 等「点开始打印后 window.open 出来的那个 PDF 地址」。
 *
 * 内容脚本拿不到 window.open 打开的地址（那是新标签页，而且页面跑在主世界），
 * 也不该去 hook 页面代码（会被 CSP 挡）。所以改用 tabs.onCreated / onUpdated：
 * 只要出现 .pdf 地址，而且我们正在等（pdfWaiter 非空），就抓下来、顺手把这个
 * 一闪而过的新标签页关掉，把结果回给内容脚本。
 *
 * pdfWaiter 为空时**什么都不做** —— 用户自己点开的 PDF 不能被这个扩展关掉。
 */
let pdfWaiter: { resolve: (r: PdfResponse) => void; timer: ReturnType<typeof setTimeout> } | null =
  null;

function armPdfWaiter(timeoutMs: number): Promise<PdfResponse> {
  return new Promise((resolve) => {
    if (pdfWaiter) {
      clearTimeout(pdfWaiter.timer);
      pdfWaiter.resolve({ ok: false, error: '被新的一次等待顶掉了' });
    }
    const w = {
      resolve,
      timer: setTimeout(() => {
        if (pdfWaiter !== w) return;
        pdfWaiter = null;
        resolve({
          ok: false,
          error: `等待 PDF 超时（${Math.round(timeoutMs / 1000)}s 内没看到 .pdf 地址，多半是没点到「开始打印」，或浏览器把新标签页拦了）`,
        });
      }, timeoutMs),
    };
    pdfWaiter = w;
  });
}

function onTabUrl(tabId: number, url: string | null | undefined): void {
  if (!isPdfUrl(url) || !pdfWaiter || !url) return;
  const w = pdfWaiter;
  pdfWaiter = null;
  clearTimeout(w.timer);
  console.log('[店小秘导出] 捕获到面单 PDF:', url);
  void fetchPdf(url).then(w.resolve);
  // 这是我们自己触发的一次性标签页，抓完就关，免得每次导出都开一堆 PDF 标签
  if (tabId >= 0) browser.tabs.remove(tabId).catch(() => undefined);
}

export default defineBackground(() => {
  console.log('[店小秘导出] background 已启动', { id: browser.runtime.id });

  // 面单：window.open 打开 .pdf 时会先建标签页（onCreated 可能没 url，等 onUpdated 补）
  browser.tabs.onCreated.addListener((tab) => {
    onTabUrl(tab.id ?? -1, tab.pendingUrl ?? tab.url);
  });
  browser.tabs.onUpdated.addListener((tabId, info, tab) => {
    onTabUrl(tabId, info.url ?? tab.url);
  });

  browser.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg && msg.type === 'DXM_FETCH_IMAGE' && typeof msg.url === 'string') {
      fetchImageBest(msg.url, msg.candidates)
        .then((r) => sendResponse(r))
        .catch((e) => sendResponse({ ok: false, error: String(e) }));
      return true; // 异步响应，保持消息通道
    }
    if (msg && msg.type === 'DXM_AWAIT_PDF') {
      const raw = typeof msg.timeout === 'number' ? msg.timeout : 45000;
      const timeout = Math.min(120000, Math.max(3000, raw));
      armPdfWaiter(timeout)
        .then(sendResponse)
        .catch((e) => sendResponse({ ok: false, error: String(e) }));
      return true;
    }
    if (msg && msg.type === 'DXM_CANCEL_PDF') {
      // 内容脚本没点到「开始打印」时主动撤销，免得接下来用户自己开的 PDF
      // 被当成面单抓下来、连标签页一起关掉
      if (pdfWaiter) {
        clearTimeout(pdfWaiter.timer);
        pdfWaiter.resolve({ ok: false, error: '已取消' });
        pdfWaiter = null;
      }
      sendResponse({ ok: true });
      return false;
    }
    return false;
  });
});
