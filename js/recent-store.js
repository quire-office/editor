/**
 * 最近檔案本地庫（起始頁／編輯器共用）
 *
 * 隱私紅線：本模組把用戶開過的檔案位元組存在**本機瀏覽器 IndexedDB**，
 * 零上傳、零網路請求——「最近」清單是本地快取，不是雲端服務。
 * 清單可逐檔移除（remove），一鍵清空（clearAll）；移除即真刪。
 *
 * 記錄時機：word.html 開檔成功後（含範例檔之外的一切使用者檔案）。
 * 起始頁選檔後把位元組放 pending 鍵，word.html 載入時取用並清除。
 */

const DB_NAME = 'quire-editor-recent';
const DB_VERSION = 1;
const STORE = 'files';
export const PENDING_KEY = 'pend';

function openDb() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') { reject(new Error('此瀏覽器不支援 IndexedDB')); return; }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        const os = db.createObjectStore(STORE, { keyPath: 'id' });
        os.createIndex('openedAt', 'openedAt');
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('IndexedDB 開啟失敗'));
  });
}

function tx(db, mode, fn) {
  return new Promise((resolve, reject) => {
    const t = db.transaction(STORE, mode);
    const req = fn(t.objectStore(STORE));
    t.oncomplete = () => resolve(req && req.result);
    t.onerror = () => reject(t.error || new Error('IndexedDB 交易失敗'));
    t.onabort = () => reject(t.error || new Error('IndexedDB 交易中止'));
  });
}

function hashName(name, bytes) {
  // 穩定 id：檔名＋大小＋首尾取樣雜湊（FNV-1a，够用且零依賴）。
  const view = new Uint8Array(bytes);
  let h = 0x811c9dc5;
  const s = name + ':' + view.length;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193); }
  const picks = [0, 1, Math.floor(view.length / 2), view.length - 1];
  for (const p of picks) { if (p >= 0 && p < view.length) { h ^= view[p]; h = Math.imul(h, 0x01000193); } }
  return (h >>> 0).toString(36) + '-' + view.length;
}

export async function recordOpen(name, bytes) {
  const db = await openDb();
  try {
    const id = hashName(name, bytes);
    await tx(db, 'readwrite', (os) => os.put({ id, name, bytes, openedAt: Date.now() }));
    return id;
  } finally { db.close(); }
}

export async function listRecent(limit = 20) {
  const db = await openDb();
  try {
    const all = ((await tx(db, 'readonly', (os) => os.getAll())) || [])
      // pending 交接鍵與最近清單共用 store（省一個 store 的代價）：它是一次性
      // 交接物、不是使用者開過的檔案，清單必須排除（u4j 燈2a 坐實）。
      .filter((r) => r && r.id !== PENDING_KEY);
    all.sort((a, b) => b.openedAt - a.openedAt);
    // 清單只回中繼資料＋位元組長度；bytes 由 getRecent 單獨取，避免清單載入整庫。
    return all.slice(0, limit).map(({ id, name, openedAt, bytes }) => ({
      id, name, openedAt, size: bytes ? bytes.length : 0,
    }));
  } finally { db.close(); }
}

export async function getRecent(id) {
  const db = await openDb();
  try { return (await tx(db, 'readonly', (os) => os.get(id))) || null; }
  finally { db.close(); }
}

export async function removeRecent(id) {
  const db = await openDb();
  try { await tx(db, 'readwrite', (os) => os.delete(id)); }
  finally { db.close(); }
}

export async function clearAll() {
  const db = await openDb();
  try { await tx(db, 'readwrite', (os) => os.clear()); }
  finally { db.close(); }
}

export async function putPending(bytes, name) {
  const db = await openDb();
  try { await tx(db, 'readwrite', (os) => os.put({ id: PENDING_KEY, name, bytes, openedAt: Date.now() })); }
  finally { db.close(); }
}

export async function takePending() {
  const db = await openDb();
  try {
    const rec = (await tx(db, 'readonly', (os) => os.get(PENDING_KEY))) || null;
    if (rec) await tx(db, 'readwrite', (os) => os.delete(PENDING_KEY));
    return rec;
  } finally { db.close(); }
}
