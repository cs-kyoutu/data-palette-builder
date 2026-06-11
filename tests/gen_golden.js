#!/usr/bin/env node
/*
 * node参照実装(procedure_pipeline.js)で正式手順書テキストを生成し、
 *   tests/procedure_cases.json  … 入力(input_tables + ステップ列)。node/python共通入力。
 *   tests/procedure_golden.json … node基準の期待出力 {name: text}
 * を書き出す。python側テストはこの golden と一致するかを検証する。
 */
const fs = require('fs');
const path = require('path');
const { renderStep } = require('../skills/procedure_pipeline.js');

const payload = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'docs', 'test_crosssell_payload.json'), 'utf8'));
const input_tables = payload.input_tables;

const cases = [
  { name: '連結', operation: '連結', settings: { 連結対象: ['姓', '名'], 保存名: '顧客氏名' } },
  { name: '連結_区切り文字', operation: '連結', settings: { 連結対象: ['姓', '名'], 区切り文字: ' ', 保存名: '顧客氏名' } },
  { name: '絞り込み_整数以上', operation: '絞り込み', settings: { 絞込み項目: '購入金額', 絞込み条件: { 演算子: '以上', 値: 5000 } } },
  { name: '絞り込み_テキスト含む', operation: '絞り込み', settings: { 絞込み項目: '都道府県', 絞込み条件: { 演算子: '含む', 値: '東京' } } },
  { name: 'IF文', operation: 'IF文', settings: { 対象項目: '性別', 条件: '完全一致', 比較値: '女性', 格納値: '1', その他の格納値: '0' } },
  { name: '集約', operation: '集約', settings: { まとめる単位: ['顧客ID'], 集約定義: [{ 対象項目: '購入金額', 集計方法: '合計' }] } },
  { name: 'ランキング', operation: 'ランキング', settings: { まとめる単位: '顧客ID', ランキング順: '購入金額', 昇順降順: '降順', 行番号同率: '同率なし' } },
  { name: '横統合', operation: '横統合', settings: { 左ファイル: '購買ログ_2026年1月', 右ファイル: '顧客マスタ', 統合方法: '左', 統合キー: '顧客ID', 残すカラム: ['性別', '都道府県'] } },
  { name: '型変換', operation: '型変換', settings: { 変換対象項目: '購入金額', 変換後の型: 'テキスト型' } },
  { name: 'カラム名の変更', operation: 'カラム名の変更', settings: { 変更前: '購入金額', 変更後: '税込金額' } },
  { name: 'テキスト挿入', operation: 'テキスト挿入', settings: { 対象カラム: '顧客ID', 位置: '文頭', 挿入テキスト: 'C-', 保存名: '顧客ID整形' } },
  { name: '分割', operation: '分割', settings: { 分割対象項目: 'メールアドレス', 分割条件: '@', 何項目に分けるか: 2 } },
  { name: '四則演算_2カラム', operation: '四則演算', settings: { 演算対象_左側: '購入金額', 演算対象_右側: '数量', 四則演算種別: '/', 小数点: '1', 端数処理: '四捨五入' } },
  { name: '四則演算_1カラム', operation: '四則演算', settings: { 演算対象_左側: '購入金額', 演算対象_右側: '1.1', 四則演算種別: 'x' } },
  { name: '追加_真偽', operation: '追加', settings: { カラム名: 'フラグ', データ型: '真偽値型', 格納値: 'True' } },
  { name: '除外', operation: '除外', settings: { 除外対象項目: '都道府県', 除外方法: '右から', 除外文字列: '県' } },
  { name: '書式変換', operation: '書式変換', settings: { 変換対象項目: 'メールアドレス', 変換後: '半角' } },
  { name: '0埋め', operation: '0埋め', settings: { 対象カラム: '顧客ID', 桁数: 8 } },
  { name: '参照', operation: '参照', settings: { まとめる単位: '顧客ID', 参照する項目: '購入金額', 参照する順: '最後の値', 昇順降順: '昇順' } },
  { name: '時刻演算', operation: '時刻演算', settings: { 引かれる値: { 種別: 'カラム', カラム名: '受注日時' }, 引く値: { 種別: 'カラム', カラム名: '登録日' }, 算出単位: '日' } },
  { name: '時刻演算_加算', operation: '時刻演算', settings: { 演算種別: '加算', 基準日時: { 種別: 'カラム', カラム名: '登録日' }, 加減算量: 7, 算出単位: '日', 保存名: '登録日_7日後' } },
  { name: '時刻演算_減算', operation: '時刻演算', settings: { 演算種別: '減算', 基準日時: { 種別: 'カラム', カラム名: '登録日' }, 加減算量: 30, 算出単位: '日', 保存名: '登録日_30日前' } },
  { name: '時刻演算_本日差分', operation: '時刻演算', settings: { 引かれる値: { 種別: 'カラム', カラム名: '登録日' }, 引く値: { 種別: '本日' }, 算出単位: '日', 保存名: '経過日数' } },
  { name: '時刻演算_特定日差分', operation: '時刻演算', settings: { 引かれる値: { 種別: 'カラム', カラム名: '登録日' }, 引く値: { 種別: '固定値', 値: '2026/01/01' }, 算出単位: '日', 保存名: '基準日からの日数' } },
  { name: '縦統合', operation: '縦統合', settings: { 統合ファイル: ['購買ログ_2026年1月', '購買ログ_2026年2月'], 重複設定: '重複を残す' } },
  { name: '置換', operation: '置換', settings: { 置換対象項目: '都道府県', 検索種別: '次の値', 置換前: '東京都', 置換後: '東京' } },
  { name: '名寄せ', operation: '名寄せ', settings: { 名寄せキー: ['顧客ID'], 判定項目: '受注日時', 判定順: '最新' } },
  { name: '抽出', operation: '抽出', settings: { 抽出対象項目: '顧客ID', 抽出方法: '先頭', 文字数: 4 } },
  { name: '削除', operation: '削除', settings: { 削除対象項目: '店舗コード' } },
  { name: 'テンプレート_縦横変換', operation: 'テンプレート_縦横変換', settings: { 集約キー: '顧客ID', 横並びカラム: ['商品名'], 並び順カラム: '受注日時', 並び順: '降順', 件数: 3 } },
  { name: 'テンプレート_年齢算出', operation: 'テンプレート_年齢算出', settings: { 生年月日カラム: '登録日' } },
];

const ctx = { input_tables };
const golden = {};
for (const c of cases) golden[c.name] = renderStep({ operation: c.operation, settings: c.settings }, ctx).text;

fs.writeFileSync(path.join(__dirname, 'procedure_cases.json'),
  JSON.stringify({ input_tables, cases: cases.map(({ name, operation, settings }) => ({ name, operation, settings })) }, null, 2), 'utf8');
fs.writeFileSync(path.join(__dirname, 'procedure_golden.json'), JSON.stringify(golden, null, 2), 'utf8');

console.log(`wrote ${cases.length} cases + golden`);
for (const [k, v] of Object.entries(golden)) console.log(`  ${k}: ${v.slice(0, 70)}`);
