/**
 * セッション履歴 整形スクリプト
 *
 * 使い方:
 * 1. アプリのヘッダーから「📥 セッション履歴CSV」をダウンロード
 * 2. Google Sheets で「ファイル」→「インポート」→「アップロード」
 *    インポート場所は「現在のシートを置換」、シート名は「raw」
 *    （シート名がデフォルトのままなら名前を「raw」に変更）
 * 3. メニュー「📊 セッション履歴」→「クリーンアップ実行」
 * 4. 「全セッション」「最終アウトプット」シートに整形結果が出力される
 *
 * 出力シート:
 *   - 全セッション : 全 row を日本語カラム名・人間可読な transcript で整理
 *   - 最終アウトプット : Phase 3（設計書・手順書）のみを抜粋＋精度評価欄
 */

const RAW_SHEET = 'raw';
const CLEAN_SHEET = '全セッション';
const FINAL_SHEET = '最終アウトプット';

// アプリのベースURL。Excelダウンロードリンクの相対パスに前置する。
const APP_BASE_URL = 'http://dpb-alb-1673181131.ap-northeast-1.elb.amazonaws.com';

const MODE_LABELS = {
  'consultation': 'Phase 1: 施策相談',
  'organization': 'Phase 2: テーブル整理',
  '': 'Phase 3: 設計書・手順書生成',
};

const CLEAN_HEADERS = [
  'セッションID',
  '段階',
  '最初の質問',
  '施策名',
  '施策概要',
  'アウトプット項目',
  'インプットテーブル',
  '設計書概要',
  '処理ステップ数',
  'メッセージ数',
  '手順書本文',
  'Excelダウンロード',
  '会話全文',
];

const FINAL_HEADERS = [
  'セッションID',
  '設計書概要',
  '処理ステップ数',
  '手順書本文',
  'Excelダウンロード',
  '精度評価 (1-5)',
  'コメント',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 セッション履歴')
    .addItem('クリーンアップ実行', 'runCleanup')
    .addToUi();
}

function runCleanup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_SHEET);
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert(
      '"raw" シートが見つかりません。CSVをこの名前のシートにインポートしてください。'
    );
    return;
  }

  const data = rawSheet.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('rawシートにデータがありません。');
    return;
  }

  const idx = {};
  data[0].forEach((h, i) => { idx[String(h).trim()] = i; });

  const cleanRows = data.slice(1).map(row => buildCleanRow(row, idx));

  writeSheet(ss, CLEAN_SHEET, CLEAN_HEADERS, cleanRows);

  const finalRows = filterFinalOutputs(cleanRows);
  writeSheet(ss, FINAL_SHEET, FINAL_HEADERS, finalRows);

  SpreadsheetApp.getUi().alert(
    `完了\n` +
    `  全セッション: ${cleanRows.length} 件\n` +
    `  最終アウトプット: ${finalRows.length} 件`
  );
}

function buildCleanRow(row, idx) {
  const get = (k) => idx[k] !== undefined ? (row[idx[k]] || '') : '';
  const rawMode = String(get('mode'));
  const stage = MODE_LABELS[rawMode] !== undefined ? MODE_LABELS[rawMode] : rawMode;
  return [
    get('session_id'),
    stage,
    get('first_user_message'),
    get('strategy_name'),
    get('strategy_summary'),
    get('output_columns'),
    get('input_table_names'),
    get('design_summary'),
    get('processing_steps_count'),
    get('message_count'),
    get('procedure_text'),
    absolutizeUrl(String(get('excel_download_url'))),
    formatTranscript(String(get('full_transcript'))),
  ];
}

/**
 * "/api/download/..." の相対URLを APP_BASE_URL を前置した絶対URLにする。
 * 既に http(s) で始まっていればそのまま返す。空文字も空文字のまま。
 */
function absolutizeUrl(url) {
  if (!url) return '';
  if (/^https?:\/\//i.test(url)) return url;
  if (url.charAt(0) !== '/') return APP_BASE_URL + '/' + url;
  return APP_BASE_URL + url;
}

function filterFinalOutputs(cleanRows) {
  const stageIdx = CLEAN_HEADERS.indexOf('段階');
  const procIdx = CLEAN_HEADERS.indexOf('手順書本文');
  const sidIdx = CLEAN_HEADERS.indexOf('セッションID');
  const sumIdx = CLEAN_HEADERS.indexOf('設計書概要');
  const stepsIdx = CLEAN_HEADERS.indexOf('処理ステップ数');
  const linkIdx = CLEAN_HEADERS.indexOf('Excelダウンロード');

  return cleanRows
    .filter(r => r[stageIdx] === 'Phase 3: 設計書・手順書生成' && r[procIdx])
    .map(r => [
      r[sidIdx],
      r[sumIdx],
      r[stepsIdx],
      r[procIdx],
      r[linkIdx],
      '',  // 精度評価
      '',  // コメント
    ]);
}

/**
 * full_transcript を人が読みやすい形に変換。
 *  - 役割ごとに区切り線
 *  - ```json``` ブロック内の JSON を整形(インデント付き)
 */
function formatTranscript(transcript) {
  if (!transcript) return '';
  const sep = '━'.repeat(50);
  const blocks = transcript.split('\n\n');
  const out = [];

  for (const raw of blocks) {
    const m = raw.match(/^\[(\w+)\]\s*([\s\S]*)$/);
    if (!m) {
      if (raw.trim()) out.push(raw);
      continue;
    }
    const role = m[1];
    const roleLabel = role === 'user' ? '👤 ユーザー'
                    : role === 'assistant' ? '🤖 AI'
                    : role;
    const content = prettyPrintJsonBlocks(m[2].trim());
    out.push(`${sep}\n${roleLabel}\n${sep}\n${content}`);
  }
  return out.join('\n\n');
}

/**
 * メッセージ内の ```json {...} ``` を JSON.stringify(indent=2) で再フォーマット。
 */
function prettyPrintJsonBlocks(content) {
  return content.replace(/```json\s*\n?([\s\S]*?)\n?```/g, (_, jsonStr) => {
    try {
      const parsed = JSON.parse(jsonStr.trim());
      return '```json\n' + JSON.stringify(parsed, null, 2) + '\n```';
    } catch (e) {
      return '```json\n' + jsonStr.trim() + '\n```';
    }
  });
}

function writeSheet(ss, name, headers, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  // ヘッダー装飾
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('white');
  sh.setFrozenRows(1);
  // 手順書本文・会話全文 列は広く + ワードラップ
  const wideTargets = ['手順書本文', '会話全文'];
  wideTargets.forEach(label => {
    const col = headers.indexOf(label);
    if (col >= 0) {
      sh.setColumnWidth(col + 1, 600);
      sh.getRange(2, col + 1, Math.max(rows.length, 1), 1)
        .setWrap(true)
        .setVerticalAlignment('top');
    }
  });
}
