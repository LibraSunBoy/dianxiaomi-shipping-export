<script lang="ts" setup>
// 店小秘发货导出助手 —— 弹窗：使用说明 + 页面悬浮面板的开关。
// 实际操作在页面内的悬浮面板上完成。
import { computed, onMounted, ref } from 'vue';
import { browser } from 'wxt/browser';

/** 页面上面板当前是否展开；null = 还没问到（页面里没有插件或没响应） */
const visible = ref<boolean | null>(null);
const busy = ref(false);
const hint = ref('');

const label = computed(() => {
  if (visible.value === null) return '页面未响应';
  return visible.value ? '收起页面面板' : '显示页面面板';
});

/** 找到店小秘标签页发消息；返回 null = 找不到页面 */
async function send(type: string) {
  const tabs = await browser.tabs.query({ url: ['*://*.dianxiaomi.com/*'] });
  const tab = tabs.find((t) => t.active) ?? tabs[0];
  if (!tab?.id) {
    hint.value = '没找到已打开的店小秘页面，请先打开「发货成功列表」。';
    return null;
  }
  return (await browser.tabs.sendMessage(tab.id, { type })) as
    | { ok?: boolean; visible?: boolean }
    | undefined;
}

/** 打开弹窗时先问一下面板现在是展开还是收起，按钮文案才对得上 */
onMounted(async () => {
  try {
    const resp = await send('DXM_PANEL_STATE');
    if (resp?.ok) visible.value = !!resp.visible;
  } catch {
    visible.value = null;
  }
});

async function toggle() {
  busy.value = true;
  hint.value = '';
  try {
    const resp = await send('DXM_TOGGLE_PANEL');
    if (!resp?.ok) {
      hint.value = '页面没有响应，刷新店小秘页面后重试。';
      return;
    }
    visible.value = !!resp.visible;
    hint.value = resp.visible
      ? '面板已展开，关掉本弹窗即可。'
      : '面板已收起，刷新页面即可重新展开。';
  } catch {
    hint.value = '页面里还没加载插件，刷新店小秘页面后重试。';
  } finally {
    busy.value = false;
  }
}
</script>

<template>
  <div class="card">
    <h1>店小秘发货导出助手</h1>
    <ol>
      <li>打开 <b>店小秘</b> 并进入「<b>发货成功列表</b>」页面。</li>
      <li>页面右下角会出现蓝色悬浮面板。</li>
      <li>确认「导出文件名」为当前登录账号用户名（可手动修改）。</li>
      <li>点击「<b>导出Excel 和面单</b>」：先下当页全部订单的面单（<b>用户名_日期_面单.pdf</b>），再自动滚动列表加载商品图，导出 <b>用户名_日期.xlsx</b>。</li>
    </ol>

    <button class="open" id="panel-toggle" :disabled="busy || visible === null" @click="toggle">
      {{ busy ? '切换中…' : label }}
    </button>
    <p v-if="hint" class="hint">{{ hint }}</p>

    <p class="note">导出字段：图片 / 尺寸(CM) / 件数 / 材质 / 运单号（材质默认「水洗底」）。</p>
  </div>
</template>

<style scoped>
.card {
  width: 280px;
  padding: 16px;
  font-family: -apple-system, 'PingFang SC', 'Microsoft YaHei', sans-serif;
  color: #333;
  font-size: 13px;
  line-height: 1.6;
}
h1 {
  font-size: 16px;
  margin: 0 0 12px;
  color: #2b6cff;
}
ol {
  margin: 0 0 12px;
  padding-left: 18px;
}
.open {
  width: 100%;
  padding: 8px 0;
  border: none;
  border-radius: 6px;
  background: #2b6cff;
  color: #fff;
  font-size: 13px;
  font-weight: 600;
  cursor: pointer;
}
.open:hover:enabled {
  background: #1f5cf0;
}
.open:disabled {
  background: #b9c8ee;
  cursor: not-allowed;
}
.hint {
  margin: 8px 0 0;
  font-size: 12px;
  color: #d9822b;
}
.note {
  font-size: 12px;
  color: #888;
  background: #f6f8fb;
  border-radius: 8px;
  padding: 10px;
  margin: 10px 0 0;
}
</style>
