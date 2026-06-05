#!/usr/bin/env node
/*
 * 要素2 本体の参照実装(node): processing_step(AIのsettings) → 正式手順書テキスト の
 * エンドツーエンド・パイプライン。Python(template_engine.py)移植の仕様兼検証。
 *
 *   renderStep(step, ctx)
 *     = resolveColumnType(対象カラム)             … 型逆引き(SQL→b→dash)
 *     → EXTRACTORS[op](settings, colType, ctx)    … settings→{hints, columns, values, choices, files, numbers}
 *     → selectEntry(op, {分岐ヒント}, colType)     … 正式エントリ選択(ハイブリッド分岐)
 *     → fillSlots(entry, 配列群)                   … 位置スロット充填
 *
 * 別称マップ(AI縮約enum→正式分岐文字列)を抽出器側に置くことで selectEntry は汎用のまま。
 * 実行: node skills/procedure_pipeline.js
 */
const R = require('./resolve_procedure.js');

// ============ 別称マップ: AI演算子 → 正式分岐(データ型別) ============
const FILTER_OP = {
  テキスト型: { 含む: '次を含む', 含まない: '次を含まない', 等しい: '次に完全一致', 等しくない: '次に完全一致しない',
    空白: '空文字', 空白でない: '空文字ではない', リストに含まれる: '次を含む(カンマ区切りで複数指定)',
    リストに含まれない: '次を含まない(カンマ区切りで複数指定)', で始まる: '次で始まる', で終わる: '次で終わる' },
  整数型: { 等しい: '次と等しい', 等しくない: '次と等しくない', 以上: '以上', 以下: '以下',
    より大きい: '次より大きい', より小さい: '次より小さい', 空白: 'NULL', 空白でない: 'NULLではない' },
  日付型: { 期間: '次の期間にある', 過去N日以内: '次の期間にある', 空白: 'NULL', 空白でない: 'NULLではない' },
};
FILTER_OP.小数型 = FILTER_OP.整数型;
FILTER_OP.日時型 = FILTER_OP.日付型;

const IF_OP = { // IF文: AIの条件文字列 → 正式分岐(代表)
  含む: '次を含む', 含まない: '次を含まない', 等しい: '次に完全一致', 完全一致: '次に完全一致',
  以上: '以上', 以下: '以下', より大きい: '次より大きい', より小さい: '次より小さい',
  NULL: 'NULL', 空: '空文字', true: 'true', false: 'false',
};

const AGG_FUNC = { 合計: '合計', 平均: '平均', 最大: '最大', 最小: '最小', 最大値: '最大', 最小値: '最小',
  カウント: 'カウント', ユニークカウント: 'ユニークカウント', 最新: '最新', 最古: '最古', 行結合: '行結合',
  SUM: '合計', AVG: '平均', MAX: '最大', MIN: '最小', COUNT: 'カウント' };

const asList = v => Array.isArray(v) ? v : (v == null || v === '' ? [] : [v]);
const firstColOf = s => s.対象カラム || s.絞込み項目 || s.抽出対象項目 || s.置換対象項目 || s.変換対象項目 ||
  s.対象項目 || s.複製対象項目 || s.削除対象項目 || (asList(s.連結対象)[0]) || (asList(s.まとめる単位)[0]) ||
  (asList(s.名寄せキー)[0]) || s.変更前 || '';

// ============ 操作別 抽出器 ============
const EXTRACTORS = {
  連結(s) {
    const cols = asList(s.連結対象).length ? asList(s.連結対象) : [s.連結対象1, s.連結対象2, s.連結対象3].filter(Boolean);
    const save = s.保存名 || s.new_column || '';
    if (s.区切り文字) {
      return { hints: ['[テキスト挿入]'], columns: cols, values: [s.区切り文字, save],
               choices: [s.表示方法 || '残さない'] };
    }
    return { hints: [], columns: cols, values: [save].filter(Boolean),
             choices: [s.表示方法 || '残さない'] };
  },
  削除: s => ({ hints: [], columns: asList(s.columns).length ? asList(s.columns) : [s.削除対象項目] }),
  複製: s => ({ hints: [], columns: [s.複製対象項目 || s.対象カラム] }),
  'カラム名の変更': s => ({ hints: [], columns: [s.変更前 || s.from], values: [s.変更後 || s.to] }),
  絞り込み(s, colType) {
    const cond = s.絞込み条件 || {};
    const ope = cond.演算子 || s.演算子 || '';
    const formal = (FILTER_OP[colType] && FILTER_OP[colType][ope]) || ope;
    const val = cond.値 != null ? cond.値 : (s.値 != null ? s.値 : '');
    const isPeriod = ope === '期間' || ope === '過去N日以内';
    const hints = [formal];
    if (isPeriod) hints.push(ope === '過去N日以内' ? '相対期間' : '絶対期間');
    return { hints, columns: [s.絞込み項目 || cond.カラム], values: val !== '' ? [String(val)] : [],
             choices: [formal] };
  },
  IF文(s) {
    const cond = String(s.条件 || s.match_type || '');
    const formal = IF_OP[cond] || (Object.keys(IF_OP).find(k => cond.includes(k)) ? IF_OP[Object.keys(IF_OP).find(k => cond.includes(k))] : '次に完全一致');
    const col = s.対象項目 || s.column;
    const trueVal = s.格納値 != null ? s.格納値 : (s.then || '');
    const elseVal = s.その他の格納値 != null ? s.その他の格納値 : (s.else || '');
    // skeleton: col1が val1《分岐》の場合、val2/col2/《空白》... val3/col3/《空白》
    return { hints: [formal], columns: [col, '', ''], values: [s.比較値 || s.value || '', trueVal, elseVal] };
  },
  集約(s) {
    const keys = asList(s.まとめる単位 || s.集約キーカラム || s.group_by);
    const defs = asList(s.集約定義 || s.aggregations);
    const tgt = defs[0] ? (defs[0].対象項目 || defs[0].column) : s.集約対象項目;
    const fnRaw = defs[0] ? (defs[0].集計方法 || defs[0].function) : s.集計方法;
    const fn = AGG_FUNC[fnRaw] || fnRaw || 'カウント';
    return { hints: [fn], columns: [keys[0], tgt], choices: [fn] };
  },
  ランキング(s) {
    const tie = { 行番号: '同率なし', 同率あり: '同率あり(順位飛ばしあり)', 同率なし: '同率なし' }[s.行番号同率] || s.行番号同率 || '同率なし';
    const ord = s.昇順降順 === '降順' ? '大きい順' : '小さい順';
    const grp = s.まとめる単位 || s.グループカラム;
    if (grp) return { hints: ['グループ'], columns: [grp, s.ランキング順 || s.ソートカラム], choices: [ord, tie] };
    return { hints: [], columns: [s.ランキング順 || s.ソートカラム], choices: [ord, tie] };
  },
  名寄せ(s, colType) {
    const order = { 最新: '最も新しい日付', 最古: '最も古い日付', 最大: '最も大きい値', 最小: '最も小さい値' }[s.判定順] || s.判定順;
    return { hints: [order], columns: [asList(s.名寄せキー)[0], s.判定項目], choices: [order] };
  },
  抽出(s) {
    const method = { 先頭から: '先頭', 末尾から: '末尾' }[s.抽出方法] || s.抽出方法 || '先頭';
    return { hints: [method, 'テキスト型'], columns: [s.抽出対象項目], values: [String(s.文字数 || '')],
             choices: [method, 'テキスト型'] };
  },
  横統合(s) {
    const method = { 左: '先に選択したデータに対して統合する', 右: '後に選択したデータに対して統合する',
      内部: '共通のデータのみ統合する', 完全外部: '全てのデータを統合する' }[s.統合方法] || s.統合方法;
    const key = typeof s.統合キー === 'object' ? Object.values(s.統合キー) : asList(s.統合キー);
    return { hints: [], files: [s.左ファイル, s.右ファイル, s.左ファイル, s.右ファイル],
             columns: [key[0] || '', key[1] || key[0] || '', ...asList(s.残すカラム)],
             values: [s.保存名 || ''], choices: [method, '統合処理をエラーにせず実行する', '更新する'] };
  },
  縦統合(s) {
    const dup = s.重複設定 === '重複を除く' ? '重複データを除外する' : '重複データを除外しない';
    return { hints: [], files: asList(s.統合ファイル), values: [s.保存名 || ''], choices: ['更新する'] };
  },
  型変換(s) {
    const t = s.変換後の型 || s.convert_to;
    return { hints: [t], columns: [s.変換対象項目], choices: [t, '上書き保存'], values: [s.保存名 || ''] };
  },
  置換: s => ({ hints: [s.検索種別 || '次の値'], columns: [s.置換対象項目], values: [s.置換前, s.置換後, s.保存名 || ''], choices: [s.検索種別 || '次の値', '上書き保存'] }),
  時刻演算(s) {
    const u = s.算出単位 || s.unit || '日';
    const col = x => (x && typeof x === 'object') ? x.カラム名 : x;
    if (s.演算種別 === '加算' || s.演算種別 === '減算') {
      const sign = s.演算種別 === '加算' ? '+' : '-';
      return { hints: ['カスタム加減算'], columns: [col(s.基準日時)],
               choices: [sign, u, '残さない'], values: [String(s.加減算量 || ''), s.保存名 || ''] };
    }
    return { hints: ['2カラム'], columns: [col(s.引かれる値), col(s.引く値)], choices: [u, '残さない'], values: [s.保存名 || ''] };
  },
  並び替え: s => ({ hints: [], columns: asList(s.並び順) }),

  テキスト挿入: s => ({ hints: [], columns: [s.対象カラム], values: [s.挿入テキスト, s.保存名 || ''],
    choices: [s.位置 || '文末', '上書き保存'] }),
  分割: s => ({ hints: [], columns: [s.分割対象項目], values: [s.分割条件 || s.区切り文字 || ''],
    choices: [s.開始位置 || '左'], numbers: [s.何項目に分けるか || 2] }),
  四則演算(s, colType, ctx) {
    const right = s.演算対象_右側;
    const twoCol = isCol(right, ctx);
    const op = { '*': '×', x: '×', '/': '÷' }[s.四則演算種別] || s.四則演算種別 || '+';
    const round = s.端数処理 || '四捨五入';
    const dec = String(s.小数点 || '1');
    if (twoCol) return { hints: ['2カラム'], columns: [s.演算対象_左側, right],
      choices: [op, dec, round, '残さない'], values: [s.保存名 || ''] };
    return { hints: ['1カラム'], columns: [s.演算対象_左側],
      values: [String(right ?? ''), s.保存名 || ''], choices: [op, dec, round, '残さない'] };
  },
  追加(s) {
    const dt = s.データ型 || s.data_type || 'テキスト型';
    const pos = s.位置 || s.position || '右に追加';
    if (['日付型', '日時型'].includes(dt)) {
      if (s.加工処理実行日 || s.processing_date) return { hints: ['日付型', 'チェックを入れる'], columns: [s.カラム名 || ''], choices: [pos, dt] };
      return { hints: ['日付型', 'チェックを入れない'], columns: [s.カラム名 || ''], choices: [pos, dt], values: [s.格納値 ?? ''] };
    }
    if (dt === '真偽値型') return { hints: ['真偽値型'], columns: [s.カラム名 || ''], choices: [pos, s.格納値 === false || s.格納値 === 'False' ? 'False' : 'True'] };
    return { hints: ['テキスト型'], columns: [s.カラム名 || ''], choices: [pos, dt], values: [s.格納値 ?? ''] };
  },
  除外(s) {
    const m = s.除外方法 || '';
    const isCount = /文字/.test(m) || typeof s.除外文字数 === 'number';
    if (isCount) return { hints: ['文字数指定'], columns: [s.除外対象項目], values: [String(s.除外文字数 || s.除外文字列 || ''), s.保存名 || ''], choices: [/右/.test(m) ? '右から' : '左から', '上書き保存'] };
    return { hints: ['テキスト指定'], columns: [s.除外対象項目], values: [s.除外文字列 || '', s.保存名 || ''], choices: [m || '全ての', '上書き保存'] };
  },
  書式変換(s) {
    const opt = s.変換後 && !['半角', '全角'].includes(s.変換後);
    if (opt) return { hints: ['オプション'], columns: [s.変換対象項目], choices: [s.変換後, '上書き保存'], values: [s.保存名 || ''] };
    return { hints: [], columns: [s.変換対象項目], choices: [s.変換後 || '半角', '上書き保存'], values: [s.保存名 || ''] };
  },
  '0埋め': s => ({ hints: [], columns: [s.対象カラム], numbers: [s.桁数 || 1], choices: ['上書き保存'], values: [s.保存名 || ''] }),
  参照(s) {
    const kind = s.参照する順 || s.参照種別 || '最初の値';
    if (s.順番 || s.入力する値 && /\d/.test(String(s.入力する値))) // 特定の順番
      return { hints: ['特定の順番'], columns: [s.まとめる単位, s.参照する項目, s.参照する項目], choices: [s.昇順降順 || '昇順'], values: [String(s.順番 || s.入力する値 || ''), s.保存名 || ''] };
    if (/累計/.test(kind)) return { hints: ['累計'], columns: [s.まとめる単位, s.参照する項目, s.参照する項目], choices: [s.昇順降順 || '昇順'], values: [s.保存名 || ''] };
    return { hints: [kind], columns: [s.まとめる単位, s.参照する項目, s.参照する項目], choices: [kind, s.昇順降順 || '昇順', '無視する'], values: [s.保存名 || ''] };
  },

  // --- テンプレート系（カラム選択中心。順序通りに列を流す） ---
  'テンプレート 『顧客ごとに縦持ちのデータを横に並べて変換』': s => ({ hints: [],
    columns: [asList(s.集約キー || s.集約キーカラム)[0], ...asList(s.横並びカラム), asList(s.並び順カラム)[0]],
    choices: [s.並び順 || '昇順'], numbers: [s.件数 || 5] }),
  'テンプレート 『顧客ごとに横持ちのデータを縦に並べて変換』': s => ({ hints: [],
    columns: [...asList(s.縦持ちカラム), ...asList(s.残すカラム)] }),
  'テンプレート 『「生年月日」から「年齢」を算出』': s => ({ hints: [], columns: [s.生年月日カラム || s.対象カラム], choices: ['上書き保存'], values: [s.保存名 || ''] }),
  'テンプレート 『「都道府県」を「地域」に変換』': s => ({ hints: [], columns: [s.都道府県カラム || s.対象カラム], choices: ['上書き保存'], values: [s.保存名 || ''] }),
  'ID紐づけ(web×biz)': s => ({ hints: [],
    files: [s.webログファイル, s.受注ファイル || '', s.メールファイル || '', s.SMSファイル || '', s.LINEファイル || ''],
    columns: [s.webアクセスログID, s.ビジターID, s.PV_Click日時 || s['PV/Click日時'], s.ページURL,
      s.受注ID || '', s.顧客ID || '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    values: [s.保存名 || ''] }),
  編集方法: s => ({ hints: [], files: [s.対象ファイル || ''], choices: [s.編集方法 || '本加工データファイルを編集し上書き保存'] }),
  統合テンプレート: s => ({ hints: [], values: [s.設定項目 || '', s.設定先 || ''] }),
};

// テンプレート汎用抽出器: settings中のカラム的な値(文字列/配列)を出現順にcolumnsへ流し、
// 保存系choice/保存名valを補う。専用抽出器の無いテンプレートのフォールバック。
function genericTemplate(s) {
  const columns = [];
  let save = '';
  for (const [k, v] of Object.entries(s)) {
    if (/保存名|new_column|出力/.test(k)) { save = v; continue; }
    for (const item of asList(v)) if (typeof item === 'string' && item) columns.push(item);
  }
  return { hints: [], columns, choices: ['上書き保存'], values: [save] };
}

// 入力テーブルに存在するカラム名か
function isCol(name, ctx) {
  if (!name || !ctx) return false;
  return (ctx.input_tables || []).some(t => (t.columns || []).some(c => c.name === name));
}

// AIのテンプレート操作名 → マスタのタスク名
const TEMPLATE_ALIAS = {
  'テンプレート_縦横変換': 'テンプレート 『顧客ごとに縦持ちのデータを横に並べて変換』',
  'テンプレート_横縦変換': 'テンプレート 『顧客ごとに横持ちのデータを縦に並べて変換』',
  'テンプレート_年齢算出': 'テンプレート 『「生年月日」から「年齢」を算出』',
  'テンプレート_都道府県地域変換': 'テンプレート 『「都道府県」を「地域」に変換』',
  'テンプレート_IDマッピング': 'ID紐づけ(web×biz)',
};

// 型解決に使うカラムが firstColOf と異なる操作（集約/名寄せ等は集計対象の型で分岐する）
const TYPE_COL_OF = {
  集約: s => (asList(s.集約定義 || s.aggregations)[0] || {}).対象項目 || (asList(s.集約定義)[0] || {}).column || s.集約対象項目,
  名寄せ: s => s.判定項目 || s.優先カラム,
};

// ============ エンドツーエンド ============
function renderStep(step, ctx = {}) {
  const rawOp = step.operation || '';
  const op = TEMPLATE_ALIAS[rawOp] || rawOp; // テンプレート操作名をマスタのタスク名へ解決
  const s = step.settings || {};
  const typeCol = (TYPE_COL_OF[rawOp] && TYPE_COL_OF[rawOp](s)) || (TYPE_COL_OF[op] && TYPE_COL_OF[op](s)) || firstColOf(s);
  const colType = R.resolveColumnType(typeCol, ctx.input_tables || [], ctx.typeOverrides || {});
  const fallback = /テンプレート/.test(op) ? genericTemplate : (() => ({ hints: [] }));
  const ex = (EXTRACTORS[op] || EXTRACTORS[rawOp] || EXTRACTORS[R.taskName(op)] || fallback)(s, colType, ctx);
  const sel = R.selectEntry(op, { branch: ex.hints[0], 期間: ex.hints[1], 演算子: ex.hints[0] }, colType);
  if (!sel.entry) return { op, colType, reason: sel.reason, text: `『${op}』(マスタ未定義)` };
  const text = R.fillSlots(sel.entry, ex);
  return { op, colType, task: sel.task, reason: sel.reason, branch: sel.entry.branch, text };
}

module.exports = { renderStep, EXTRACTORS, FILTER_OP, IF_OP };

// ============ セルフテスト ============
if (require.main === module) {
  const fs = require('fs'), path = require('path');
  const payload = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'docs', 'test_crosssell_payload.json'), 'utf8'));
  const ctx = { input_tables: payload.input_tables };

  const steps = [
    { operation: '連結', settings: { 連結対象: ['姓', '名'], 保存名: '顧客氏名' } },
    { operation: '絞り込み', settings: { 絞込み項目: '購入金額', 絞込み条件: { 演算子: '以上', 値: 5000 } } },
    { operation: '絞り込み', settings: { 絞込み項目: '都道府県', 絞込み条件: { 演算子: '含む', 値: '東京' } } },
    { operation: 'IF文', settings: { 対象項目: '性別', 条件: '完全一致', 比較値: '女性', 格納値: '1', その他の格納値: '0' } },
    { operation: '集約', settings: { まとめる単位: ['顧客ID'], 集約定義: [{ 対象項目: '購入金額', 集計方法: '合計' }] } },
    { operation: 'ランキング', settings: { まとめる単位: '顧客ID', ランキング順: '購入金額', 昇順降順: '降順', 行番号同率: '同率なし' } },
    { operation: '横統合', settings: { 左ファイル: '購買ログ_2026年1月', 右ファイル: '顧客マスタ', 統合方法: '左', 統合キー: '顧客ID', 残すカラム: ['性別', '都道府県'] } },
    { operation: '型変換', settings: { 変換対象項目: '購入金額', 変換後の型: 'テキスト型' } },
    { operation: 'カラム名の変更', settings: { 変更前: '購入金額', 変更後: '税込金額' } },
  ];

  let ok = 0;
  for (const st of steps) {
    const r = renderStep(st, ctx);
    // ×は乗算演算子として正当に出力されるため残骸チェックから除外（未充填は{}で検知）
    const clean = !/[{}]|[○〇◯▢△]/.test(r.text);
    if (clean) ok++;
    console.log(`### ${r.op}  [型=${r.colType}/${r.reason}/branch=${r.branch}]`);
    console.log(`  ${r.text}`);
    if (!clean) console.log('  ✗ 残骸あり');
    console.log();
  }
  console.log(`=== ${ok}/${steps.length} 構造OK（意味的妥当性は目視確認） ===`);
}
