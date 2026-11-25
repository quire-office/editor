#!/usr/bin/env bash
# U4k BUGGY 態覆核：暫時摘除三處修復（Enter 退回軟換行／段頭退格退回誠實拒絕／
# 打字不吃選區），探針必紅才算有鑑別力；跑完回滾＋git diff 驗證乾淨。
set -euo pipefail
cd ~/projects/editor
ADAPTER=js/engine-adapter.js
BAK=/home/jerry/.hermes/cache/scratch/u4k_buggy_adapter.bak
cp "$ADAPTER" "$BAK"
trap 'cp "$BAK" "$ADAPTER"; echo "[replay] 已回滾"; git diff --stat -- "$ADAPTER" | tail -1; [ -z "$(git diff -- "$ADAPTER")" ] || { echo "[replay] 警告：回滾後仍有 diff！"; exit 3; }' EXIT

python3 - <<'EOF'
s = open('js/engine-adapter.js').read()
# 1) Enter 退回舊 InsertBreak（不分流）
s = s.replace("      : buildSplitParagraphMutation(targetPara, targetCp);",
              "      : buildInsertBreakMutation(targetPara, targetCp, 'LineWrapping'); // BUGGY-replay")
# 2) 段頭退格退回誠實拒絕
s = s.replace("      if (targetPara <= 0) {\n        return { ok: false, error: '已位於文件開頭，無可合併的前段' };",
              "      return { ok: false, error: 'BUGGY-replay 單段刪除暫不支援跨段合併' };\n      if (targetPara <= 0) {\n        return { ok: false, error: '已位於文件開頭，無可合併的前段' };")
# 3) 打字不吃選區
s = s.replace("    const pre = this._deleteSelectionIfExists();\n    if (pre.error) return { ok: false, error: pre.error };\n    const targetPara = para !== undefined ? Number(para) : (this.caret ? this.caret.para : 0);\n    const targetCp = cp !== undefined ? Number(cp) : (this.caret ? this.caret.cp : 0);\n    const mutation = buildInsertTextMutation(targetPara, targetCp, text);",
              "    void this.normalizedSelection(); // BUGGY-replay 打字不吃選區\n    const targetPara = para !== undefined ? Number(para) : (this.caret ? this.caret.para : 0);\n    const targetCp = cp !== undefined ? Number(cp) : (this.caret ? this.caret.cp : 0);\n    const mutation = buildInsertTextMutation(targetPara, targetCp, text);")
open('js/engine-adapter.js','w').write(s)
print('[replay] BUGGY 態注入完成')
EOF

timeout 300 node tests/u4k_paragraph_semantics_probe.mjs 2>&1 | grep -E "✗|判定" || true
RC=$(timeout 300 node tests/u4k_paragraph_semantics_probe.mjs >/dev/null 2>&1; echo $?)
echo "BUGGY 態退出碼：$RC（非 0＝探針有鑑別力）"
[ "$RC" != "0" ]
