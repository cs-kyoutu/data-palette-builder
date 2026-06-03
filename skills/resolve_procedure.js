#!/usr/bin/env node
/*
 * 要素2 参照実装 (node): procedure_format_master.json を消費して
 * processing_step → 正式手順書テキスト を生成するアルゴリズムの検証用。
 * Python(template_engine.py)移植前に、型解決・エントリ選択(ハイブリッド分岐)・
 * 位置スロット充填 のロジックを実データで検証する。
 *
 * 実行: node skills/resolve_procedure.js   (末尾のセルフテストが走る)
 */
const fs = require('fs');
const path = require('path');
const MASTER = JSON.parse(fs.readFileSync(path.join(__dirname, 'procedure_format_master.json'), 'utf8'));

// --- SQL型 → b→dash型 ---
const TYPE_MAP = {
  STRING: 'テキスト型', TEXT: 'テキスト型', VARCHAR: 'テキスト型',
  INTEGER: '整数型', INT: '整数型', BIGINT: '整数型',
  FLOAT: '小数型', DOUBLE: '小数型', DECIMAL: '小数型', NUMERIC: '小数型',
  DATE: '日付型', DATETIME: '日時型', TIMESTAMP: '日時型',
  BOOLEAN: '真偽値型', BOOL: '真偽値型',
};
function toBdashType(sqlType) {
  return TYPE_MAP[String(sqlType || '').toUpperCase()] || 'テキスト型';
}

// --- AI/エンジンの操作名 → マスタのタスク名 ---
const OP_ALIAS = {
  '絞込み': '絞り込み',
  'カラム名変更': 'カラム名の変更',
};
function taskName(op) {
  if (MASTER[op]) return op;
  if (OP_ALIAS[op] && MASTER[OP_ALIAS[op]]) return OP_ALIAS[op];
  // テンプレート系: 部分一致
  const t = Object.keys(MASTER).find(k => k === op || k.includes(op) || op.includes(k));
  return t || op;
}

// --- 入力テーブルからカラムのb→dash型を逆引き ---
function resolveColumnType(colName, inputTables, typeOverrides = {}) {
  if (typeOverrides[colName]) return typeOverrides[colName]; // 中間生成カラムの型トラッキング
  for (const t of inputTables || []) {
    for (const c of t.columns || []) {
      if (c.name === colName) return toBdashType(c.type);
    }
  }
  return null; // 不明
}

// --- エントリ選択（ハイブリッド: IF文/絞り込みは分岐厳格, それ以外はbest-effort）---
const STRICT_OPS = new Set(['IF文', '絞り込み']);
function selectEntry(op, settings, colType) {
  const task = taskName(op);
  const entries = MASTER[task];
  if (!entries) return { task, entry: null, reason: 'no-master' };

  // 1) データ型で絞る（空=ワイルドカード）
  let cands = entries.filter(e => e.data_types.length === 0 || (colType && e.data_types.includes(colType)));
  if (cands.length === 0) cands = entries; // 型不明時は全候補

  // 2) 分岐条件マッチ（複合ヒント対応: 期間種別+演算子 等を全て突き合わせ、最多一致を選ぶ）
  const hints = [
    settings.branch, settings.分岐, settings.期間, settings.期間種別,
    settings.演算子, settings.operator, settings.条件, settings.集計方法, settings.抽出方法,
  ].map(x => (x == null ? '' : String(x))).filter(Boolean);
  if (hints.length) {
    const strict = STRICT_OPS.has(task);
    const score = e => {
      if (!e.branch) return 0;
      const hay = e.branch + ' ' + e.slots.filter(s => s.kind === 'choice').flatMap(s => s.options).join(' ');
      return hints.reduce((n, h) => n + (hay.includes(h) ? 1 : 0), 0);
    };
    const ranked = cands.map(e => ({ e, sc: score(e) })).filter(x => x.sc > 0).sort((a, b) => b.sc - a.sc);
    if (ranked.length) return { task, entry: ranked[0].e, reason: `branch-match(${ranked[0].sc}/${hints.length})` };
    if (strict) return { task, entry: cands.find(e => e.branch) || cands[0], reason: 'strict-fallback' };
  }
  // 3) 分岐ヒント無し → branch=null(汎用) を優先、無ければ先頭
  return { task, entry: cands.find(e => e.branch === null) || cands[0], reason: 'default' };
}

// --- 位置スロット充填 ---
// values = {columns:[], values:[], choices:[], files:[], numbers:[]}
// 各 kind の出現順に対応する配列要素を消費。choice は値が options に含まれればそれを、無ければ先頭。
function fillSlots(entry, values) {
  const idx = { column: 0, value: 0, file: 0, number: 0, choice: 0 };
  const pick = { column: values.columns || [], value: values.values || [],
                 file: values.files || [], number: values.numbers || [], choice: values.choices || [] };
  let s = entry.skeleton;
  for (const slot of entry.slots) {
    const arr = pick[slot.kind];
    const raw = arr[idx[slot.kind]++];
    let v;
    if (slot.kind === 'column') v = `「${raw ?? '　'}」`;
    else if (slot.kind === 'file') v = `【${raw ?? '　'}】`;
    else if (slot.kind === 'value') v = `"${raw ?? ''}"`;
    else if (slot.kind === 'number') v = String(raw ?? slot.range[0]);
    else if (slot.kind === 'choice') v = (raw && slot.options.includes(raw)) ? raw : (raw || slot.options[0]);
    s = s.replace(`{${slot.id}}`, v);
  }
  return s.trim();
}

module.exports = { toBdashType, taskName, resolveColumnType, selectEntry, fillSlots, MASTER };

// ============ セルフテスト ============
if (require.main === module) {
  const payload = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'docs', 'test_crosssell_payload.json'), 'utf8'));
  const tables = payload.input_tables;

  const cases = [
    { name: '連結(姓+名)', op: '連結', col: '姓',
      values: { columns: ['姓', '名'], values: ['顧客氏名'], choices: ['残さない'] } },
    { name: '絞り込み(購入金額>=)', op: '絞り込み', col: '購入金額',
      settings: { 演算子: '以上' }, values: { columns: ['購入金額'], values: ['5000'], choices: ['以上'] } },
    { name: '絞り込み(登録日 相対期間)', op: '絞り込み', col: '登録日',
      settings: { 演算子: '次の期間にある', 期間: '相対期間' },
      values: { columns: ['登録日'], numbers: [0, 30], choices: ['相対期間', '年', '前', '次の期間にある'] } },
    { name: 'IF文(性別→フラグ)', op: 'IF文', col: '性別',
      settings: { 条件: '次に完全一致' },
      values: { columns: ['性別'], values: ['女性', '1', '0'], } },
    { name: '集約(顧客ID×購入金額合計)', op: '集約', col: '購入金額',
      settings: { 集計方法: '合計' },
      values: { columns: ['顧客ID', '購入金額'], choices: ['合計'] } },
  ];

  let ok = 0;
  for (const c of cases) {
    const colType = resolveColumnType(c.col, tables);
    const { task, entry, reason } = selectEntry(c.op, c.settings || {}, colType);
    console.log(`### ${c.name}`);
    console.log(`  col=${c.col} → 型=${colType} | task=${task} | select=${reason} | branch=${entry && entry.branch}`);
    if (entry) {
      const text = fillSlots(entry, c.values);
      console.log(`  → ${text}`);
      if (!/[{}]|[○〇◯▢△×]/.test(text)) ok++;
      else console.log('  ✗ 残骸あり');
    } else {
      console.log('  ✗ エントリ選択失敗');
    }
    console.log();
  }
  console.log(`=== ${ok}/${cases.length} 充填OK ===`);
  process.exit(ok === cases.length ? 0 : 1);
}
