# Quire 文書編輯器外殼

Quire 文書編輯器的瀏覽器外殼：視覺上對齊傳統桌上型文書處理器 Ribbon 佈局的零構建（zero-build）靜態應用，零第三方依賴、零 CDN，透過本倉的引擎適配層直連 Quire 排版引擎的 WebAssembly 產物，完成開檔、打字、格式化、表格操作與存檔的真實編輯閉環。

本倉於 roadmap 的「外殼分倉」條款自引擎倉（`quire-office/quire` 的 `shells/web/editor/`）獨立而來，歷史經 `git subtree split` 完整保留。線上部署：`https://editor.quireoffice.dev`。

---

## 1. 快速開始

引擎 wasm 產物**嚴禁進本倉**（含 git 歷史）。 clone 後第一步先取得產物：

```bash
npm run fetch:wasm   # 依 engine.assets.json 自引擎倉 GitHub release 下載＋sha256 校驗
npm start            # 本機靜態伺服器 http://127.0.0.1:8791/
```

開發環境若引擎倉有本地建置，可設 `QUIRE_ENGINE_PKG=/path/to/quire/shells/web/pkg`
讓 fetch 腳本走本機降級（醒目警告「非 release 產物」；CI 嚴禁設定此變數）。

直接以瀏覽器開啟 `index.html` 亦可（ES Module 需经 HTTP，`npm start` 即最簡路徑）。

## 2. 引擎產物取得契約（分倉的核心邊界）

- `engine.assets.json` 是唯一信任來源：釘死引擎版本與每個產物的 sha256。
- `scripts/fetch-wasm.mjs` 的降級鏈：release 下載 →（僅開發）本機引擎 pkg → 佔位說明檔。sha256 不符即死，絕不混入壞資產。
- 升版流程：引擎倉發 release → 更新 `engine.assets.json` 的 version 與 sha256 → `npm run fetch:wasm` 驗證 → 提交清單變更。外殼倉不追引擎 main 的浮動產物。

## 3. 架構與事件匯流排

1. **純邏輯層 (`js/qword-core.js`)**：零 DOM 依賴，可在 Node.js 直接 `import` 做無頭驗證；`QWordCore` 管理標籤、檢視模式、縮放、標尺與指令日誌，入口為 `onCommand(cmdId, payload, requirement)`。
2. **DOM 綁定層 (`js/qword-ui.js`)**：事件委派、Ribbon 鍵盤導航、標尺滑桿、縮放、彈出層；掛載 `window.QWordShell` 與 `window.__qwordLog`。
3. **引擎適配層 (`js/engine-adapter.js`)**：wasm 初始化（`vendor/pkg/`）、開檔／多頁 Canvas 繪製、Mutation 呼叫（打字、格式、表格、圖片、節）、Undo/Redo、游標反查，以及 wasm 邊界的誠實底線（引擎回傳字串兩級正規化，壞字串炸成「這次操作失敗」而非整個外殼無回應）。
4. **繪製器 (`vendor/renderer.js`)**：與官網 playground 共用同一顯示列表繪製器（來源登記見 `vendor/SYNC.md`），避免雙實作漂移。

探針可經 `window.QWordShell.onCommand(...)` 派發命令、`getState()` 讀狀態、`window.__qwordLog` 讀日誌。

## 4. 產品誠實紅線

每顆按鈕帶 `data-requirement`：
- `live`：外殼已真接引擎 Mutation 或外殼自身真功能。
- `pending:<卡>`：**可見但停用**，tooltip 如實說明缺口在引擎側哪張切片。引擎沒實作的（校閱面、參考資料面、主題面等）絕不假裝可用、絕不藏按鈕。

即時統計（勿信文件快照，跑這個）：

```bash
grep -o 'data-requirement="[^"]*"' index.html | sort | uniq -c | sort -rn
```

## 5. 考卷與探針

- 免瀏覽器節點自檢：`npm run selfcheck`（u1–u4e 七支；讀引擎源碼的斷言需 `QUIRE_ENGINE_REPO=/path/to/quire`，未設則如實點名跳過）。
- 真瀏覽器探針（`tests/`，需 Chrome `--remote-debugging-port=9222` 與 `npm start` 的 8791 伺服）：Ribbon 逐鈕位移（u4c）、格內 caret 與跑版（u4f）、底線形態畫布幾何（u4g）等；證據 JSON 落 `evidence/`。
- 圖示精靈：`icons/` 為單一真相源，改圖後 `npm run sprite` 重烘 `index.html` 的 sprite（禁止手改）。

## 6. 部署

`.github/workflows/deploy-pages.yml`：fetch:wasm → 七支節點自檢（引擎源碼斷言用引擎倉 shallow checkout）→ 組裝 `dist/`（只含執行時靜態檔＋CNAME）→ GitHub Pages。缺任一執行時檔即紅燈，不部署半成品。
