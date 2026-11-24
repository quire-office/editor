// 編輯器外殼的引擎 wasm 資產取得管線（自 quire-www 的 fetch-wasm.mjs 同形改寫）
// 目的：本倉嚴禁直接提交引擎二進位；wasm 產物在 build／啟動前由本腳本取得。
//
// 取得順序（降級鏈）：
//   1) GitHub release 下載（URL 模板見 engine.assets.json），sha256 必須與清單相符；
//      相符才寫入 vendor/pkg/。
//   2) release 下載失敗 → 降級：自環境變數 QUIRE_ENGINE_PKG 指向的本機引擎
//      shells/web/pkg/ 複製，並列印醒目警告「使用本機建置，非 release」；
//      sha256 仍需與清單相符，不符視同清單過期，直接失敗。未設該環境變數時
//      不啟用降級（CI 由此天然禁降級）。
//   3) 兩者皆失敗 → 佔位模式：產生 vendor/pkg/MISSING-ASSETS.md（說明跑
//      npm run fetch:wasm），本腳本以 0 退出；頁面 JS 偵測 wasm 404 時顯示補救指示。
//
// 用法：node scripts/fetch-wasm.mjs（npm run fetch:wasm；亦由 prestart 自動呼叫）

import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const MANIFEST_PATH = join(ROOT, 'engine.assets.json');

const note = (msg) => console.log(`[fetch-wasm] ${msg}`);
const warn = (msg) => console.warn(`[fetch-wasm] ⚠️  ${msg}`);
const die = (msg) => {
  console.error(`\n[fetch-wasm] ✋ 失敗：${msg}`);
  console.error('[fetch-wasm] 修正後重跑：npm run fetch:wasm');
  process.exit(1);
};

function writeMissingGuide(reason) {
  const lines = [
    '# 引擎產物缺檔（佔位說明，非引擎錯誤）',
    '',
    '本倉構建／啟動時未取得編輯器所需的 WASM 引擎產物',
    '（`quire_web_bg.wasm`、`quire_web.js`）。引擎二進位依規定不進本倉，',
    '需先執行：',
    '',
    '```sh',
    'npm run fetch:wasm',
    '```',
    '',
    `自動取得失敗原因：${reason}`,
    '',
    '腳本會依 `engine.assets.json` 的記載自 GitHub release 下載並以 sha256 驗證；',
    '開發環境亦可設 `QUIRE_ENGINE_PKG=/path/to/quire/shells/web/pkg` 走本機降級。',
  ];
  mkdirSync(targetDir, { recursive: true });
  writeFileSync(join(targetDir, 'MISSING-ASSETS.md'), lines.join('\n'));
  note('已產生 vendor/pkg/MISSING-ASSETS.md（佔位說明）');
}

const manifest = JSON.parse(readFileSync(MANIFEST_PATH, 'utf8'));
const targetDir = join(ROOT, manifest.target_dir ?? 'vendor/pkg/');
if (!manifest.assets?.length) die(`${MANIFEST_PATH} 內 assets 為空，無法取得任何資產`);

function sha256File(path) {
  const h = createHash('sha256');
  h.update(readFileSync(path));
  return h.digest('hex');
}

function verifyCopy(srcPath, entry, sourceLabel) {
  const actual = sha256File(srcPath);
  if (actual !== entry.sha256) {
    die(
      `sha256 驗證不相符（${sourceLabel}）。檔案：${entry.file}\n` +
        `  期望：${entry.sha256}\n  實得：${actual}\n` +
        `  若引擎已更新，請以新產物重算 sha256 並更新 engine.assets.json，勿直接跳過驗證。`
    );
  }
  mkdirSync(targetDir, { recursive: true });
  const dest = join(targetDir, entry.file);
  rmSync(dest, { force: true });
  copyFileSync(srcPath, dest);
  note(`✔ ${entry.file}（sha256 相符，來源：${sourceLabel}）→ ${resolve(dest)}（${statSync(dest).size} bytes）`);
}

function releaseUrlFor(entry) {
  return (manifest.release_url_template ?? '').replace('{version}', manifest.version).replace('{file}', entry.file);
}

async function tryFetchRelease(entry) {
  const url = releaseUrlFor(entry);
  if (!url) throw new Error('清單缺 release_url_template');
  note(`嘗試自 release 下載：${url}`);
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);
  const buf = Buffer.from(await res.arrayBuffer());
  const tmp = join(targetDir, `.${entry.file}.download`);
  mkdirSync(targetDir, { recursive: true });
  writeFileSync(tmp, buf);
  const actual = createHash('sha256').update(buf).digest('hex');
  if (actual !== entry.sha256) {
    rmSync(tmp, { force: true });
    die(
      `release 下載的 ${entry.file} sha256 驗證不相符。\n` +
        `  期望：${entry.sha256}\n  實得：${actual}\n` +
        `  可能是 release 被換檔或清單過期；請確認後更新 engine.assets.json。`
    );
  }
  const dest = join(targetDir, entry.file);
  rmSync(dest, { force: true });
  copyFileSync(tmp, dest);
  rmSync(tmp, { force: true });
  note(`✔ ${entry.file}（sha256 相符，來源：GitHub release v${manifest.version}）`);
}

function tryLocalFallback(entry) {
  const dir = process.env.QUIRE_ENGINE_PKG;
  if (!dir) throw new Error('未設 QUIRE_ENGINE_PKG，本機降級不啟用');
  const src = join(dir, entry.file);
  if (!existsSync(src)) throw new Error(`${src} 不存在`);
  printWarningBox([
    `release 下載失敗，改用本機引擎建置：${src}`,
    '這是「非 release 產物」——僅限開發環境；CI 嚴禁設定 QUIRE_ENGINE_PKG。',
  ]);
  verifyCopy(src, entry, `本機 ${dir}`);
}

function printWarningBox(lines) {
  console.warn('[fetch-wasm]', '='.repeat(64));
  for (const line of lines) console.warn(`[fetch-wasm] ${line}`);
  console.warn('[fetch-wasm]', '='.repeat(64));
}

async function main() {
  note(`版本：v${manifest.version}`);
  let failure = null;
  for (const entry of manifest.assets) {
    const dest = join(targetDir, entry.file);
    // 已在位且相符 → 免重抓（開發迴圈省流量）
    if (existsSync(dest) && sha256File(dest) === entry.sha256) {
      note(`✔ ${entry.file} 已在位且 sha256 相符，跳過下載`);
      continue;
    }
    try {
      await tryFetchRelease(entry);
      rmSync(join(targetDir, 'MISSING-ASSETS.md'), { force: true });
    } catch (err) {
      try {
        tryLocalFallback(entry);
        rmSync(join(targetDir, 'MISSING-ASSETS.md'), { force: true });
      } catch (err2) {
        failure = `${entry.file}: release（${err.message}）；本機降級（${err2.message}）`;
        break;
      }
    }
  }
  if (failure) {
    writeMissingGuide(failure);
    warn(`全部取得路徑失敗：${failure}`);
    warn('頁面將以「引擎資產缺失」指示呈現；其餘靜態頁不受影響。');
    process.exit(0);
  }
  note('全部資產就緒。');
}

main();
