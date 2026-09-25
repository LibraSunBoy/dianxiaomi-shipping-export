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

export default defineBackground(() => {
  console.log('[店小秘导出] background 已启动', { id: browser.runtime.id });

  browser.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg && msg.type === 'DXM_FETCH_IMAGE' && typeof msg.url === 'string') {
      fetchImageBest(msg.url, msg.candidates)
        .then((r) => sendResponse(r))
        .catch((e) => sendResponse({ ok: false, error: String(e) }));
      return true; // 异步响应，保持消息通道
    }
    return false;
  });
});
