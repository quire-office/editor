#!/usr/bin/env node
/**
 * live-drift-check.mjs — 產線（editor.quireoffice.dev）與 origin/main 漂移檢查
 *
 * 用戶要求（11-26）：「editor 站是否已經更新？列入流程，用程式保證或更新。」
 * 三層判據（全部機械可判，零目視）：
 *  L1 部署版本＝Pages 部署綁 main：拉 origin/main 的 runtime 檔 sha256 對比產線
 *     同檔 sha256。任何 DIFF ＝ 產線落後或被手動改動。
 *  L2 引擎資產＝產線 vendor/pkg/* 的 sha256 必等 engine.assets.json 宣告值
 *     （CI fetch:wasm 已校驗過，這裡防的是「assets.json 更新了但部署沒重跑」）。
 *  L3 版本前進＝origin/main 的 engine.assets.json version 若比產線實際 wasm 新
 *     （L2 FAIL 且 main 的 assets 已改），提示需部署；附建議動作。
 * 用法：
 *   node scripts/live-drift-check.mjs            # 檢查，exit 0=同步 / 1=漂移
 *   node scripts/live-drift-check.mjs --dispatch # 漂移時自動觸發 deploy-pages
 *     （workflow_dispatch 需 CI token 或 repo token；env GH_TOKEN 覆蓋）
 * 依賴：僅 stdlib＋git（archive 讀 origin/main，不需工作樹髒）。
 */
import { execFileSync } from 'node:child_process';

const BASE = process.env.LIVE_BASE || 'https://editor.quireoffice.dev';
const RUNTIME_FILES = [
  'index.html', 'word.html',
  'css/style.css',
  'js/qword-ui.js', 'js/qword-core.js', 'js/engine-adapter.js', 'js/recent-store.js',
  'vendor/renderer.js',
];
// vendor/pkg/* 不在 git（fetch:wasm 依 engine.assets.json 拉入，產物嚴禁進倉），
// L1 對其豁免，由 L2（產線 wasm sha == assets.json 宣告）把關。
const ASSET_FILES = ['vendor/pkg/quire_web.js', 'vendor/pkg/quire_web_bg.wasm'];
const git = (...a) => execFileSync('git', a, { cwd: process.cwd(), encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
const sha = (buf) => import('node:crypto').then(c => c.createHash('sha256').update(buf).digest('hex'));

async function main() {
  const dispatch = process.argv.includes('--dispatch');
  // origin/main 為準（工作樹可能有未 push 的本地 commit，那不構成產線義務）
  git('fetch', 'origin', 'main', '--quiet');
  const mainSha = git('rev-parse', 'origin/main').trim();

  // 從 origin/main 抽檔做 hash（git archive 管線，不碰工作樹）
  const localHashes = new Map();
  for (const f of RUNTIME_FILES) {
    let buf;
    try {
      buf = execFileSync('git', ['show', `origin/main:${f}`], { cwd: process.cwd(), maxBuffer: 256 * 1024 * 1024 });
    } catch { localHashes.set(f, 'ABSENT'); continue; }
    localHashes.set(f, await sha(buf));
  }
  const assetsMain = JSON.parse(git('show', 'origin/main:engine.assets.json'));

  const lines = [];
  let drift = 0;
  let liveWasmSha = null;
  let liveGlueSha = null;
  for (const f of RUNTIME_FILES) {
    let live;
    try {
      const r = await fetch(`${BASE}/${f}`, { cache: 'no-store' });
      if (!r.ok) { lines.push(`MISS  ${f}  http=${r.status}`); drift++; continue; }
      live = await sha(Buffer.from(await r.arrayBuffer()));
    } catch (e) { lines.push(`ERR   ${f}  ${String(e).slice(0, 60)}`); drift++; continue; }
    if (localHashes.get(f) === live) {
      lines.push(`SAME  ${f}`);
    } else if (localHashes.get(f) === 'ABSENT') {
      lines.push(`GONE  ${f}  （main 已刪但產線仍供）`); drift++;
    } else {
      lines.push(`DIFF  ${f}  main=${(localHashes.get(f) || '').slice(0, 12)} live=${live.slice(0, 12)}`);
      drift++;
    }
  }
  // 產線引擎資產 sha（L1 豁免項，供 L2 信任錨比對）
  for (const f of ASSET_FILES) {
    try {
      const r = await fetch(`${BASE}/${f}`, { cache: 'no-store' });
      if (!r.ok) { lines.push(`MISS  ${f}  http=${r.status}`); drift++; continue; }
      const h = await sha(Buffer.from(await r.arrayBuffer()));
      if (f.endsWith('.wasm')) liveWasmSha = h; else liveGlueSha = h;
    } catch (e) { lines.push(`ERR   ${f}  ${String(e).slice(0, 60)}`); drift++; }
  }
  // L2：產線 wasm 必等 origin/main assets.json 宣告（fetch:wasm 的信任錨）
  const declared = (assetsMain.assets.find(a => a.file === 'quire_web_bg.wasm') || {}).sha256 || null;
  const declaredGlue = (assetsMain.assets.find(a => a.file === 'quire_web.js') || {}).sha256 || null;
  const l2 = declared && liveWasmSha ? (declared === liveWasmSha ? 'OK' : 'MISMATCH') : 'N/A';
  const l2g = declaredGlue && liveGlueSha ? (declaredGlue === liveGlueSha ? 'OK' : 'MISMATCH') : 'N/A';
  if (l2 === 'MISMATCH' || l2g === 'MISMATCH') drift++;

  console.log(`live-drift-check  base=${BASE}  origin/main=${mainSha.slice(0, 12)}  assets(main)=${assetsMain.version}`);
  for (const l of lines) console.log('  ' + l);
  console.log(`  L2 wasm==assets.json: ${l2}  (declared=${(declared || '').slice(0, 12)} live=${(liveWasmSha || '').slice(0, 12)})`);
  console.log(`  L2 glue==assets.json: ${l2g}  (declared=${(declaredGlue || '').slice(0, 12)} live=${(liveGlueSha || '').slice(0, 12)})`);

  if (drift === 0) { console.log('VERDICT: SYNCED'); return; }
  console.log(`VERDICT: DRIFT（${drift} 項）`);
  if (!dispatch) process.exitCode = 1;
  if (dispatch) {
    const token = process.env.GH_TOKEN || process.env.GITHUB_TOKEN;
    if (!token) { console.log('DISPATCH: 跳過（無 GH_TOKEN env）'); process.exitCode = 1; return; }
    const owner = process.env.LIVE_OWNER || 'quire-office';
    const repo = process.env.LIVE_REPO || 'editor';
    const wf = process.env.LIVE_WF || 'deploy-pages.yml';
    const r = await fetch(`https://api.github.com/repos/${owner}/${repo}/actions/workflows/${wf}/dispatches`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Accept': 'application/vnd.github+json', 'User-Agent': 'live-drift-check', 'Content-Type': 'application/json' },
      body: JSON.stringify({ ref: 'main' }),
    });
    console.log(`DISPATCH: deploy-pages → HTTP ${r.status}${r.ok ? '（已觸發，數分鐘後複查）' : '（失敗，需人工）'}`);
    process.exitCode = r.ok ? 0 : 1;
  }
}
main().catch(e => { console.error('CRASH:', e); process.exit(2); });
