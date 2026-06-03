#!/usr/bin/env node
/*
 * procedure_format_master.json の全エントリを、スロットにデモ値を流し込んで
 * レンダリングし、(1) 全スロットが埋まるか (2) 不正な残骸が無いか を検証する。
 * 実行: node skills/preview_procedure_master.js [--full]
 */
const m = require('./procedure_format_master.json');
const FULL = process.argv.includes('--full');

const DEMO = {
  column: ['顧客ID', '受注日', '購入金額'],
  value:  ['東京', '1000', '2020/01/01'],
  file:   ['顧客テーブル', '受注テーブル'],
};

function fill(entry) {
  let s = entry.skeleton;
  const used = new Set();
  const counters = { column: 0, value: 0, file: 0 };
  for (const slot of entry.slots) {
    let v;
    if (slot.kind === 'choice') v = slot.options[0];
    else if (slot.kind === 'number') v = String(slot.range[0]);
    else if (slot.kind === 'column') v = `「${DEMO.column[counters.column++ % DEMO.column.length]}」`;
    else if (slot.kind === 'file')   v = `【${DEMO.file[counters.file++ % DEMO.file.length]}】`;
    else /* value */                 v = `"${DEMO.value[counters.value++ % DEMO.value.length]}"`;
    s = s.replace(`{${slot.id}}`, v);
    used.add(slot.id);
  }
  return s;
}

let total = 0, bad = 0;
const problems = [];
for (const op of Object.keys(m)) {
  for (let i = 0; i < m[op].length; i++) {
    const e = m[op][i];
    total++;
    const rendered = fill(e);
    // 検証: 未置換の {slot} が残っていないか / 記号残骸が無いか
    const leftoverBraces = rendered.match(/\{[a-z]+\d+\}/g);
    const leftoverGlyph  = rendered.match(/[○〇◯▢△×]/g);
    if (leftoverBraces || leftoverGlyph) {
      bad++;
      problems.push(`${op}[${i}] branch=${e.branch}\n    ${rendered}\n    braces=${leftoverBraces} glyph=${leftoverGlyph}`);
    }
    if (FULL) {
      console.log(`### ${op} [${i}]  dt=${JSON.stringify(e.data_types)} branch=${e.branch}`);
      console.log('  ' + rendered + '\n');
    }
  }
}
console.log(`\n=== 検証結果 ===`);
console.log(`total entries : ${total}`);
console.log(`renderable OK : ${total - bad}`);
console.log(`problems      : ${bad}`);
for (const p of problems) console.log('  ✗ ' + p);
process.exit(bad ? 1 : 0);
