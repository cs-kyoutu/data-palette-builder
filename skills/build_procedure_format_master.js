#!/usr/bin/env node
/*
 * procedure_format_master_raw.txt (列区切り=2+スペース, 5列: タスク名 / データ型 / 分岐条件 / 記載フォーマット / 選択)
 * を 3階層構造の JSON マスタ (skills/procedure_format_master.json) へ変換する。
 *
 * 出力構造:
 *   { "<タスク名>": [ { "data_types": [...], "branch": "<分岐条件|null>",
 *                       "skeleton": "<スロットを{col1}/{val1}/{choice1}/{num1}/{file1}に置換した文>",
 *                       "slots": [ {id, kind, options?, range?} ... ] }, ... ] }
 *
 * スロット記法 (ユーザー提供マスタ):
 *   「○○」「〇〇」「▢▢」「△△」「××」 = カラム名スロット -> {colN}    (kind: column)
 *   "○○" "○○,○○" "〇〇/〇〇/〇〇"      = 値スロット       -> {valN}    (kind: value)
 *   [○○]                              = 設定値フィールド -> {valN}    (kind: value)
 *   《A/B/..》(スラッシュ有)            = 選択肢スロット   -> {choiceN} (kind: choice, options:[...]) ※AIが選択して確定
 *   《N～M》(範囲)                      = 数値スロット     -> {numN}    (kind: number, range:[N,M])
 *   《〇〇》《〇》(記号列)               = 値スロット       -> {valN}    (kind: value)
 *   《単一トークン》                    = 分岐マーカー/固定ラベル -> リテラル保持(置換しない)
 *   【○○】                            = ファイル名スロット-> {fileN}   (kind: file)
 *   【固定名】[適用]等                  = リテラル保持
 */
const fs = require('fs');
const path = require('path');

const RAW = path.join(__dirname, '..', 'docs', 'procedure_format_master_raw.txt');
const OUT = path.join(__dirname, 'procedure_format_master.json');

// 特殊ダブルクオート(U+201C/U+201D/U+2033/U+FF02)を ASCII " に正規化（手入力ゆらぎ対策）
const raw = fs.readFileSync(RAW, 'utf8').replace(/[“”″＂]/g, '"');
const lines = raw.split(/\r?\n/);

const CIRCLE  = '○〇◯▢△×';                            // スロットを表す記号類
const COL_RE  = new RegExp(`「[${CIRCLE}]+」`, 'g');     // 「○○」= カラム名
const FILE_RE = new RegExp(`【[${CIRCLE}]+】`, 'g');     // 【○○】= ファイル名(固定名は除外)
const VAL_RE  = new RegExp(`"[^"]*[${CIRCLE}][^"]*"`, 'g'); // "..."内に記号を含む = 値
const BRK_RE  = new RegExp(`[\\[［][${CIRCLE}]+[\\]］]`, 'g'); // [○○]= 設定値フィールド

function convert(fmt) {
  let s = fmt;
  // sentinel: 書式テキストに出現しない @@kind|meta@@ トークン。idは後で左→右に付与。
  const T = (kind, meta) => `@@${kind}|${meta || ''}@@`;

  // --- パス1: 各スロットを sentinel トークンへ置換 ---
  // 《...》を最優先で処理（内部に "値" やスラッシュを含む複合選択肢を壊さないため）。
  // 選択肢内に残る "○○" 等はオプションラベルのリテラルとして保持する。
  s = s.replace(/《([^》]*)》/g, (m, inner) => {
    const rangeM = inner.match(/^(\d+)\s*[～~]\s*(\d+)$/);
    if (rangeM) return T('number', `${rangeM[1]}~${rangeM[2]}`);
    if (inner.includes('/')) {
      // 選択肢ラベル内にネストした記号(例 特定の日付:"○○")は後続VAL_REを汚染するため … に中性化
      const clean = inner.replace(new RegExp(`[${CIRCLE}]+`, 'g'), '…');
      return T('choice', clean);
    }
    if (new RegExp(`^[${CIRCLE}]+$`).test(inner)) return T('value');
    return m; // 単一トークン=分岐マーカー/固定ラベル → リテラル保持
  });
  s = s.replace(FILE_RE, () => T('file'));
  s = s.replace(COL_RE,  () => T('column'));
  s = s.replace(VAL_RE,  () => T('value'));
  s = s.replace(BRK_RE,  () => T('value'));

  // --- パス2: 左→右に走査し kind ごとに連番 id を付与 ---
  const slots = [];
  const counters = { file: 0, column: 0, value: 0, number: 0, choice: 0 };
  const PREFIX = { file: 'file', column: 'col', value: 'val', number: 'num', choice: 'choice' };
  const skeleton = s.replace(/@@([a-z]+)\|([^@]*)@@/g, (m, kind, meta) => {
    const id = `${PREFIX[kind]}${++counters[kind]}`;
    const slot = { id, kind };
    if (kind === 'choice') slot.options = meta.split('/').map(x => x.trim()).filter(Boolean);
    if (kind === 'number') { const [a, b] = meta.split('~'); slot.range = [Number(a), Number(b)]; }
    slots.push(slot);
    return `{${id}}`;
  });

  return { skeleton: skeleton.replace(/\s+/g, ' ').trim(), slots };
}

// 列区切りは「2個以上の半角スペース」。書式列内は単一スペースなので衝突しない。
const DELIM = / {2,}/;
const result = {};
const seen = new Set();
let parsed = 0, skipped = 0, dupes = 0;
const oddRows = [];

for (let li = 0; li < lines.length; li++) {
  const line = lines[li];
  if (!line.trim()) continue;
  const parts = line.split(DELIM).map(p => p.trim()).filter(p => p !== '');
  if (parts.length < 4) { skipped++; oddRows.push(`L${li + 1}: <4cols`); continue; }
  const task   = parts[0];
  const dtype  = parts[1];
  const branch = parts[2];
  // 4列目=書式。「選択」列(およびそれ以降の混入テキスト)で打ち切る。
  let fmtParts = parts.slice(3);
  const selIdx = fmtParts.findIndex(p => /^選択/.test(p));
  if (selIdx !== -1) fmtParts = fmtParts.slice(0, selIdx);
  const fmt = fmtParts.join(' ').trim();
  if (selIdx === -1) oddRows.push(`L${li + 1}: 「選択」列が見つからず (cols=${parts.length}) task=${task.slice(0, 16)}`);
  if (!task || !fmt) { skipped++; continue; }

  const key = [task, dtype, branch, fmt].join('');
  if (seen.has(key)) { dupes++; continue; }
  seen.add(key);

  const data_types = (dtype === 'ー' || dtype === '全て' || dtype === '')
    ? [] // 空配列 = 全データ型に適用
    : dtype.split('/').map(x => x.trim()).filter(Boolean);
  const branchVal = (branch === 'ー' || branch === '') ? null : branch;

  const { skeleton, slots } = convert(fmt);

  if (!result[task]) result[task] = [];
  result[task].push({ data_types, branch: branchVal, skeleton, slots });
  parsed++;
}

fs.writeFileSync(OUT, JSON.stringify(result, null, 2), 'utf8');

// ---- レポート ----
const ops = Object.keys(result);
console.log(`parsed entries : ${parsed}`);
console.log(`skipped lines  : ${skipped}`);
console.log(`dropped dupes  : ${dupes}`);
console.log(`operations     : ${ops.length}`);
console.log(`output         : ${OUT}`);
console.log('--- entries per operation ---');
for (const op of ops) {
  const slotKinds = {};
  for (const e of result[op]) for (const sl of e.slots) slotKinds[sl.kind] = (slotKinds[sl.kind] || 0) + 1;
  const kindStr = Object.entries(slotKinds).map(([k, v]) => `${k}:${v}`).join(' ');
  console.log(`  ${op}  →  ${result[op].length} 変形   [${kindStr}]`);
}
if (oddRows.length) {
  console.log('--- 非定型行(要確認) ---');
  for (const r of oddRows) console.log('  ' + r);
}
