/**
 * u4l_format_scope_probe.mjs — U4l：字元格式作用域＋拖曳選取 瀏覽器閉環探針
 *
 * 用戶投訴（11-23 實錄）兩案：
 * (1)「粗體、斜體直接影響一個段落，而非圈選的段落」——根因＝外殼 getEffectiveRange
 *    「未選取→整段」回退（引擎 FormatTextRange 粒度實測正確）。修復＝needsSelection
 *    誠實拒絕＋狀態列提示。本探針鎖：未選取按粗體**零位元組位移**＋提示出現。
 * (2)「無法使用滑鼠拖曳選取段落」——根因＝pointermove 綁容器級且 e.target 须在頁內，
 *    拖出頁緣/跨頁即斷流。修復＝window 級監聽＋最近頁鉗制。本探針鎖：CDP 真滑鼠
 *    拖曳（含頁緣灰區外一點、跨頁終點）選區成立。
 *
 * 判定口徑（誠實分層）：
 * - 燈A 未選取拒絕：真 DOM click 粗體鈕 → 段落 XML 片段零位移＋選取高亮零塊＋
 *   提示 pill 可見且文案點名「圈選」。（BUGGY 覆核：restore 整段回退→本燈必紅）
 * - 燈B 局部性：真 DOM click（先設圈選）→ 段落 XML 出現 <w:b/> 且**bold run 文字
 *   ≠ 整段文字**（局部），undo 後段落片段回基準。（BUGGY 覆核：舊行為是整段加粗
 *   →bold run 文字＝整段文字，本燈同樣紅＝對舊缺陷有鑑別力）
 * - 燈C 拖曳選取：CDP Input.dispatchMouseEvent 全程（pressed→moves→released），
 *   中途一點刻意落在頁緣灰區外（x=頁右緣+14px，驗鉗制續流），終點落在下一頁；
 *   選區必非 null，且跨頁時 end.para > start.para。
 * - 燈D 拖回起點：同一路徑拖回原點放開 → selection 歸 null（Word 收回語義）。
 *
 * 用法（worktree 根起靜態伺服；Chrome --remote-debugging-port=9222）：
 *   python3 -m http.server 8791 --bind 127.0.0.1 &
 *   node tests/u4l_format_scope_probe.mjs
 * 環境變數：U4L_URL（默認 http://127.0.0.1:8791/word.html）、CDP_URL（默認 9222）、
 *           U4L_OUT（證據 JSON 路徑）。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BASE_URL = process.env.U4L_URL || 'http://127.0.0.1:8791/word.html';
const OUT = process.env.U4L_OUT || path.join(ROOT, 'evidence/word-ui-u4l-format-scope-evidence.json');
const CDP_HTTP = process.env.CDP_URL || 'http://127.0.0.1:9222';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const list = await (await fetch(`${CDP_HTTP}/json/list`)).json();
  const created = await (await fetch(`${CDP_HTTP}/json/new?about:blank`, { method: 'PUT' })).json();
  const sock = new WebSocket(created.webSocketDebuggerUrl);
  await new Promise((res, rej) => { sock.onopen = res; sock.onerror = rej; });
  let seq = 0;
  const pending = new Map();
  sock.onmessage = (ev) => {
    const msg = JSON.parse(ev.data);
    if (msg.id && pending.has(msg.id)) { pending.get(msg.id)(msg); pending.delete(msg.id); }
  };
  const send = (method, params = {}) => new Promise((res) => {
    const id = ++seq;
    pending.set(id, (m) => res(m.result || {}));
    sock.send(JSON.stringify({ id, method, params }));
  });
  const evalIn = async (expr) => {
    const r = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true, timeout: 120000 });
    if (r.exceptionDetails) throw new Error('page eval: ' + JSON.stringify(r.exceptionDetails).slice(0, 400));
    return r.result ? r.result.value : undefined;
  };
  const mouse = (type, x, y, button = 'left', buttons = 0) =>
    send('Input.dispatchMouseEvent', { type, x, y, button, buttons, clickCount: type === 'mousePressed' ? 1 : 0 });

  await send('Page.enable');
  await send('Runtime.enable');
  await send('Page.navigate', { url: BASE_URL });
  for (let i = 0; i < 240; i += 1) {
    const r = await evalIn('document.readyState');
    if (r === 'complete') break;
    await sleep(250);
  }
  // 開樣本檔、等引擎真的就緒（同 u4f 口徑）
  await evalIn('window.QWordShell.openDefaultSample()');
  for (let i = 0; i < 300; i += 1) {
    const ok = await evalIn('!!window.QWordEngine.engine && window.QWordEngine.paragraphCount() > 0');
    if (ok) break;
    await sleep(200);
  }
  if (!(await evalIn('!!window.QWordEngine.engine'))) throw new Error('開檔失敗：引擎未建立');
  await sleep(500);

  const R = { notes: {} };

  // ---- 共用：段落 XML 片段（第 paraIdx 個 <w:p>，含標籤）＋段落文字 ----
  const helpers = `(async () => {
    const A = window.QWordEngine;
    const xml = await (async () => {
      const bytes = A.save();
      const buf = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
      for (let i = 0; i + 4 <= buf.length; i += 1) {
        if (dv.getUint32(i, true) !== 0x04034b50) continue;
        const nameLen = dv.getUint16(i + 26, true);
        const extraLen = dv.getUint16(i + 28, true);
        const cSize = dv.getUint32(i + 18, true);
        const name = new TextDecoder().decode(buf.subarray(i + 30, i + 30 + nameLen));
        if (name !== 'word/document.xml') continue;
        const start = i + 30 + nameLen + extraLen;
        const comp = buf.subarray(start, start + cSize);
        try {
          const ds = new DecompressionStream('deflate-raw');
          return await new Response(new Blob([comp]).stream().pipeThrough(ds)).text();
        } catch { return new TextDecoder().decode(comp); }
      }
      return '';
    })();
    const paras = xml.match(/<w:p[ >][\\s\\S]*?<\\/w:p>/g) || [];
    return { xml, paras };
  })()`;
  const paraFragment = async (idx) => (await evalIn(helpers)).paras[idx] || '';
  const paraText = async (idx) => evalIn(`(window.QWordEngine.paragraphText(${idx}) || '')`);

  // 找一個夠長且**基準無粗體**的正文段（否則燈B 的 boldDelta／局部性判據會被
  // 示範檔自帶的 <w:b/> 污染——既有 bold run 文字可能等於整段，誤判成整段加粗）。
  // 索引對齊紀律：document.xml 的 <w:p> 序與引擎 paragraphText(i) 序必須由**文字
  // 逐字比對**證明同一段（表格段交錯時兩者可能錯位），比不過就換下一段。
  const firstSnap = await evalIn(helpers);
  const fragPlain = (frag) => (frag.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
    .map((t) => t.replace(/<[^>]+>/g, '')).join('');
  let paraIdx = -1;
  for (let i = 0; i < firstSnap.paras.length; i += 1) {
    const frag = firstSnap.paras[i];
    if (/<w:b\b/.test(frag)) continue;
    if (Array.from(fragPlain(frag)).length < 12) continue;
    const engTxt = await paraText(i);
    if (fragPlain(frag) === engTxt) { paraIdx = i; break; }
  }
  if (paraIdx < 0) throw new Error('找不到「≥12 字元、基準無粗體、XML↔引擎同段」的正文段');
  R.notes.paraIdx = paraIdx;
  R.notes.paraChars = (await paraText(paraIdx)).length;

  const pillState = () => evalIn(`(() => {
    const pill = document.querySelector('.qword-status-hint');
    if (!pill) return { found: false };
    const r = pill.getBoundingClientRect();
    return { found: true, visible: !pill.hidden && r.width > 0 && r.height > 0,
             text: (pill.querySelector('.qword-status-hint-text') || {}).textContent || '' };
  })()`);
  const selRectCount = () => evalIn(`document.querySelectorAll('.qword-selection-rect').length`);
  const clickBold = () => evalIn(`(() => {
    const b = document.querySelector('button[data-command="format:bold"]');
    if (!b) return false; b.click(); return true;
  })()`);
  const setCaretOnly = (p, cp) => evalIn(`(() => {
    const A = window.QWordEngine;
    A.caret = { para: ${p}, cp: ${cp} };
    A.selection = null;
    (window.QWordShell.render && window.QWordShell.render());
    return true;
  })()`);
  const setSelection = (p, a, b) => evalIn(`(() => {
    const A = window.QWordEngine;
    A.caret = { para: ${p}, cp: ${b} };
    A.selection = { target: 'Body', start: { para: ${p}, cp: ${a} }, end: { para: ${p}, cp: ${b} } };
    (window.QWordShell.render && window.QWordShell.render());
    return true;
  })()`);
  const drainUndo = async (max = 6) => {
    for (let i = 0; i < max; i += 1) {
      await evalIn(`window.QWordShell.onCommand('edit:undo', {})`);
      await sleep(160);
    }
  };

  const hasRenderFacade = await evalIn('typeof window.QWordShell.render === "function"');
  R.notes.hasRenderFacade = hasRenderFacade;

  // ============ 燈A：未選取按粗體 → 零位移＋提示 ============
  await setCaretOnly(paraIdx, 5);
  await sleep(150);
  const aBefore = await paraFragment(paraIdx);
  await clickBold();
  await sleep(500);
  const aAfter = await paraFragment(paraIdx);
  const aPill = await pillState();
  const aRects = await selRectCount();
  // 燈A 判據必須是**增量**：示範檔某些段落自帶 <w:b/>，用「片段是否含 <w:b/>」
  // 會把基準態誤讀成「套上了格式」。比對 <w:b 出現次數才對得上「零位移」。
  const countBold = (frag) => (frag.match(/<w:b\b/g) || []).length;
  R.A = { xmlUnchanged: aBefore === aAfter, boldBefore: countBold(aBefore),
          boldAfter: countBold(aAfter),
          pill: aPill, selectionRects: aRects };
  await drainUndo();

  // ============ 燈B：圈選局部 (0..5) 按粗體 → 局部加粗 ============
  const bBefore = await paraFragment(paraIdx);
  await setSelection(paraIdx, 0, 5);
  await sleep(200);
  const rectsWhileSel = await selRectCount();
  await clickBold();
  await sleep(500);
  const bAfter = await paraFragment(paraIdx);
  // 局部性判據（檔契約面，正是用戶保存後看得見的粒度）：Node 端正則從段落
  // XML 片段抽「含 <w:b/> 之 run 的 <w:t> 串接」，與整段文字比長度。
  const boldTexts = [];
  for (const run of bAfter.match(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g) || []) {
    if (/<w:b\b/.test(run)) {
      for (const t of run.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []) {
        boldTexts.push(t.replace(/<[^>]+>/g, ''));
      }
    }
  }
  const paraFull = await paraText(paraIdx);
  const boldJoined = boldTexts.join('');
  const partial = boldTexts.length > 0 && Array.from(boldJoined).length < Array.from(paraFull).length;
  await evalIn(`window.QWordShell.onCommand('edit:undo', {})`);
  await sleep(300);
  const bUndone = await paraFragment(paraIdx);
  const countBoldB = (frag) => (frag.match(/<w:b\b/g) || []).length;
  R.B = { rectsWhileSel, rectsAsserted: hasRenderFacade, boldDelta: countBoldB(bAfter) - countBoldB(bBefore),
          boldText: boldJoined, paraLen: Array.from(paraFull).length,
          partial, boldEqualsWhole: boldJoined === paraFull && boldTexts.length > 0,
          undoRestored: bUndone === bBefore };
  await drainUndo();
  // 收尾複核：undo 後段落既無 <w:b/> 也回到基準片段
  R.B.finalClean = (await paraFragment(paraIdx)) === bBefore && countBoldB(await paraFragment(paraIdx)) === countBoldB(bBefore);

  // ============ 燈C：CDP 真滑鼠拖曳（含頁緣灰區一點＋跨頁終點） ============
  const geo = await evalIn(`(() => {
    const A = window.QWordEngine;
    const r = A.caretRect(${paraIdx}, 2);
    if (!r) return null;
    const wrapper = document.querySelectorAll('.qword-real-page')[r.page];
    const canvas = wrapper.querySelector('canvas');
    const rect = canvas.getBoundingClientRect();
    const dl = A.getPageDisplayList(r.page);
    const scale = canvas.width / (dl.width_pt || 595);
    return { x: rect.left + r.x_pt * scale, y: rect.top + (r.y_pt - 4) * scale,
             pageRight: rect.right, pageBottom: rect.bottom };
  })()`);
  if (!geo) throw new Error('caretRect 回 None，燈C 無法幾何');
  await mouse('mousePressed', geo.x, geo.y, 'left', 1);
  await sleep(60);
  // 逐點拖曳：先段內右移 → 頁右緣外灰區 14px（驗鉗制續流）→ 頁面下方 → 下一頁
  const steps = [
    [geo.x + 40, geo.y],
    [geo.pageRight + 14, geo.y],              // 灰區外：舊碼此處斷流
    [geo.pageRight - 20, geo.pageBottom - 30],
    [geo.x + 30, geo.pageBottom + 80],        // 頁縫／下一頁方向
    [geo.x + 10, geo.pageBottom + 150]        // 終點（頁數少時自動落最遠頁）
  ];
  for (const [x, y] of steps) {
    await mouse('mouseMoved', Math.max(1, x), Math.max(1, y), 'left', 1);
    await sleep(90);
  }
  const midSel = await evalIn('JSON.stringify(window.QWordEngine.selection)');
  await mouse('mouseReleased', Math.max(1, steps[steps.length - 1][0]), Math.max(1, steps[steps.length - 1][1]), 'left', 0);
  await sleep(250);
  const endSel = JSON.parse(await evalIn('JSON.stringify(window.QWordEngine.selection)') || 'null');
  R.C = { midSel: midSel ? JSON.parse(midSel) : null, endSel,
          rectsAfter: await selRectCount() };
  R.C.pageCount = await evalIn('window.QWordEngine.pageCount()');
  R.C.dragWorked = !!endSel && (endSel.end.para > endSel.start.para
    || (endSel.end.para === endSel.start.para && endSel.end.cp > endSel.start.cp));
  R.C.crossPage = !!endSel && endSel.end.para !== endSel.start.para;
  // 跨頁硬判只在多頁文件成立時要求（單頁檔無跨頁可驗，不該誤紅）
  R.C.crossPageExpected = R.C.pageCount >= 2;

  // ============ 燈D：拖回起點 → 選區歸 null ============
  await mouse('mousePressed', geo.x, geo.y, 'left', 1);
  await sleep(60);
  await mouse('mouseMoved', geo.x + 60, geo.y, 'left', 1);
  await sleep(90);
  await mouse('mouseMoved', geo.x, geo.y, 'left', 1);   // 拖回原點
  await sleep(90);
  await mouse('mouseReleased', geo.x, geo.y, 'left', 0);
  await sleep(250);
  R.D = { selection: JSON.parse(await evalIn('JSON.stringify(window.QWordEngine.selection)') || 'null') };
  R.D.backToOriginClears = R.D.selection === null;

  try {
    const shot = await send('Page.captureScreenshot', { format: 'png' });
    if (shot && shot.data) fs.writeFileSync(OUT.replace(/\.json$/, '.png'), Buffer.from(shot.data, 'base64'));
  } catch {}
  try { await fetch(`${CDP_HTTP}/json/close/${created.id}`); } catch {}
  sock.close();
  fs.writeFileSync(OUT, JSON.stringify(R, null, 2));
  console.log(`證據檔 → ${OUT}`);

  // ---- 紅綠板 ----
  let fail = 0;
  const line = (cond, msg) => { console.log(`  ${cond ? '✓' : '✗'} ${msg}`); if (!cond) fail += 1; };
  console.log('=== U4l 格式作用域＋拖曳選取 判定 ===');
  line(R.A.xmlUnchanged && R.A.boldBefore === R.A.boldAfter,
    `燈A 未選取按粗體→段落零位元組位移（changed=${!R.A.xmlUnchanged} bold ${R.A.boldBefore}→${R.A.boldAfter}）`);
  line(R.A.pill.visible && /圈選/.test(R.A.pill.text),
    `燈A 提示 pill 可見且點名圈選（visible=${R.A.pill.visible} text="${(R.A.pill.text || '').slice(0, 30)}"）`);
  line(R.A.selectionRects === 0, `燈A 拒絕態無選取高亮（rects=${R.A.selectionRects}）`);
  line(R.B.boldDelta >= 1 && R.B.partial && !R.B.boldEqualsWhole,
    `燈B 圈選局部→局部加粗（boldDelta=${R.B.boldDelta} boldText="${R.B.boldText}" 整段=${R.B.paraLen}字）`);
  line(!R.B.rectsAsserted || R.B.rectsWhileSel > 0,
    `燈B 圈選態有高亮塊（rects=${R.B.rectsWhileSel}${R.B.rectsAsserted ? '' : '，舊版無 render facade 免判'}）`);
  line(R.B.undoRestored && R.B.finalClean, `燈B undo 回基準片段（restored=${R.B.undoRestored} clean=${R.B.finalClean}）`);
  line(R.C.dragWorked && (!R.C.crossPageExpected || R.C.crossPage),
    `燈C 真滑鼠拖曳選取成立＋拖出頁緣/跨頁續流（crossPage=${R.C.crossPage} pages=${R.C.pageCount} end=${JSON.stringify(R.C.endSel && R.C.endSel.end)}）`);
  line(R.C.rectsAfter > 0, `燈C 拖曳後有高亮渲染（rects=${R.C.rectsAfter}）`);
  line(R.D.backToOriginClears, `燈D 拖回起點選區歸零（=${JSON.stringify(R.D.selection)}）`);
  console.log(`\n判定：${fail === 0 ? 'PASS' : `FAIL（${fail} 項）`}`);
  process.exit(fail === 0 ? 0 : 1);
}

main().catch((e) => { console.error('探針失敗:', e.message); process.exit(2); });
