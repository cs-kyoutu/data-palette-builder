"""エイリアス変換の網羅的テスト
Phase1が出しそうな全パターンの変数名ブレをテスト"""
import json, sys, os
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
exec(open(os.path.join(os.path.dirname(__file__), '..', 'backend', 'template_engine_v2.py'), encoding='utf-8').read())

total = 0
passed = 0
failed = 0
errors = []

def test(name, step_json):
    global total, passed, failed
    total += 1
    step = json.loads(step_json)
    try:
        result = render_step(step)
        if '変数不足' in result:
            failed += 1
            errors.append(f"{name}: {[l for l in result.split(chr(10)) if '変数不足' in l]}")
            print(f"  [FAIL] {name}")
        else:
            passed += 1
            print(f"  [OK] {name}: {result[:80]}...")
    except Exception as e:
        failed += 1
        errors.append(f"{name}: {e}")
        print(f"  [ERR] {name}: {e}")

print("=== 名寄せ エイリアステスト ===")
# Phase1が出しそうなキー名パターン
test("名寄せ: keys", '{"operation":"名寄せ","settings":{"keys":["顧客ID","商品ID"],"priority_column":"PV/Click日時","priority_order":"最も新しい日時"}}')
test("名寄せ: key_columns", '{"operation":"名寄せ","settings":{"key_columns":["顧客ID","商品ID"],"priority_column":"PV/Click日時","priority_order":"降順"}}')
test("名寄せ: key_column(単数)", '{"operation":"名寄せ","settings":{"key_column":"顧客ID","priority_column":"受注日時","priority_order":"最も古い日時"}}')

print("\n=== ランキング エイリアステスト ===")
test("ランキング: 正規", '{"operation":"ランキング","settings":{"group_key":"顧客ID","sort_column":"売上金額","sort_order":"大きい順","tie_handling":"同率なし"}}')
test("ランキング: group_by+降順", '{"operation":"ランキング","settings":{"group_by":["顧客ID"],"sort_column":"売上金額","sort_order":"降順","rank_method":"同率なし"}}')
test("ランキング: group_byなし", '{"operation":"ランキング","settings":{"sort_column":"閲覧数","sort_order":"昇順","tie_handling":"同率あり(順位飛ばしあり)"}}')

print("\n=== テンプレート縦→横 エイリアステスト ===")
test("テンプレート: 正規", '{"operation":"テンプレート 縦持ちを横持ちに変換","settings":{"aggregation_keys":["顧客ID"],"horizontal_columns":["商品名","価格"],"sort_column":"ランキング","sort_order":"昇順","top_n":5}}')
test("テンプレート: Phase1出力", '{"operation":"テンプレート","template_name":"縦持ちを横持ちに変換","settings":{"aggregate_keys":["顧客ID"],"pivot_columns":["商品名","価格"],"sort_column":"ランキング","sort_order":"昇順","max_columns":5}}')

print("\n=== カラム名変更 エイリアステスト ===")
test("カラム名変更: renames(単数)", '{"operation":"カラム名の変更","settings":{"renames":[{"from":"IF文","to":"受注日時"}]}}')
test("カラム名変更: renames(複数)", '{"operation":"カラム名の変更","settings":{"renames":[{"from":"商品名","to":"閲覧商品名1"},{"from":"商品名(1)","to":"閲覧商品名2"}]}}')
test("カラム名変更: rename_rules", '{"operation":"カラム名変更","settings":{"rename_rules":[{"from":"IF文","to":"フラグ"},{"from":"商品名","to":"閲覧商品名1"}]}}')

print("\n=== 連結 エイリアステスト ===")
test("連結: 正規", '{"operation":"連結","settings":{"column_1":"姓","column_2":"名","separator":" ","task_name":"名前","keep_original":"残さない"}}')
test("連結: left/right_column", '{"operation":"連結","settings":{"left_column":"姓","right_column":"名","separator":" ","new_column_name":"名前","keep_original":false}}')

print("\n=== 横統合 エイリアステスト ===")
test("横統合: 正規", '{"operation":"横統合","settings":{"method":"先に選択したデータに対して統合する","left_file":"顧客","right_file":"受注","key_left":"顧客ID","key_right":"顧客ID","left_columns":["顧客ID","メール"],"right_columns":["受注日時"],"save_as":"統合1","duplicate_handling":"統合処理をエラーにする","update_setting":"更新する"}}')
test("横統合: keep_columns", '{"operation":"横統合","settings":{"method":"共通のデータのみを統合する","left_file":"A","right_file":"B","key_left":"ID","key_right":"ID","keep_columns":["カラム1","カラム2"],"save_as":"結果"}}')

print("\n=== IF文 エイリアステスト ===")
test("IF文: 正規", '{"operation":"IF文","settings":{"conditions":[{"column":"ステータス","condition":"次に完全一致","value":"公開","result":"1"}],"else_value":"《空白》"}}')
test("IF文: カラム参照", '{"operation":"IF文","settings":{"conditions":[{"column":"商品ID","condition":"次のカラムの値に完全一致","value":"商品ID(1)","result":"1"}],"else_value":"《空白》"}}')
test("IF文: 期間判定", '{"operation":"IF文","settings":{"type":"period","date_column":"本日日付","start_column":"開始日","end_column":"終了日","true_value":"公開","false_value":"非公開"}}')

print("\n=== 時刻演算 エイリアステスト ===")
test("時刻演算: カラム間", '{"operation":"時刻演算","settings":{"left_column":"催行日","operator":"-","right_column":"申込日","unit":"日","task_name":"差分","keep_original":"残す"}}')
test("時刻演算: 本日の日付", '{"operation":"時刻演算","settings":{"use_today":true,"operator":"-","right_column":"最終購入日","unit":"日","task_name":"経過日数","keep_original":"残す"}}')
test("時刻演算: カスタム", '{"operation":"時刻演算","settings":{"left_column":"本日年月","operator":"-","custom_value":"3","unit":"ヶ月","task_name":"3ヶ月前","keep_original":"残す"}}')

print("\n=== 集約 エイリアステスト ===")
test("集約: 正規", '{"operation":"集約","settings":{"group_keys":["顧客ID"],"aggregations":[{"column":"受注ID","function":"ユニークカウント"},{"column":"売上","function":"合計"}]}}')

print("\n=== 参照 エイリアステスト ===")
test("参照: 最初の値", '{"operation":"参照","settings":{"pattern":"最初の値","group_key":"ビジターID","sort_column":"PV/Click日時","sort_order":"降順","value_column":"セッションID","null_handling":"無視する","task_name":"最新セッション"}}')
test("参照: 後ろの行", '{"operation":"参照","settings":{"pattern":"後ろの行の値","group_key":"顧客ID","sort_column":"配信月","sort_order":"昇順","value_column":"配信月","null_handling":"無視する","task_name":"次回配信月"}}')

print("\n=== 四則演算 エイリアステスト ===")
test("四則演算: カラム間", '{"operation":"四則演算","settings":{"column_left":"売上","operator":"÷","column_right":"件数","decimal_place":"1","rounding":"四捨五入","task_name":"単価","keep_original":"残す"}}')
test("四則演算: 固定値", '{"operation":"四則演算","settings":{"column_left":"割合","operator":"×","fixed_value":"100","decimal_place":"1","rounding":"四捨五入","task_name":"パーセント","keep_original":"残さない"}}')

print("\n=== テキスト挿入 エイリアステスト ===")
test("テキスト挿入: 単一", '{"operation":"テキスト挿入","settings":{"column":"ROI","position":"文末","text":"％","save_method":"上書き保存","task_name":"ROI"}}')
test("テキスト挿入: 両方", '{"operation":"テキスト挿入","settings":{"column":"ITEM_CD","text_head":"https://example.com/","text_tail":".html","save_method":"名前を付けて保存","task_name":"URL"}}')

print("\n=== 置換 エイリアステスト ===")
test("置換: 通常", '{"operation":"置換","settings":{"column":"金額","from_value":"NULL","to_value":"0","save_method":"上書き保存","task_name":"金額"}}')
test("置換: 正規表現", '{"operation":"置換","settings":{"column":"コード","match_type":"《正規表現》","from_value":"^1[0-9]{4}$","to_value":"1****","save_method":"上書き保存","task_name":"コード"}}')

print("\n=== 追加 エイリアステスト ===")
test("追加: 通常", '{"operation":"追加","settings":{"base_column":"ID","position":"右に追加","data_type":"テキスト型","default_value":""}}')
test("追加: 実行日", '{"operation":"追加","settings":{"base_column":"ID","position":"右に追加","data_type":"日付型","is_execution_date":true}}')

print("\n=== 削除 エイリアステスト ===")
test("削除", '{"operation":"削除","settings":{"columns":["中間カラム1","中間カラム2"]}}')

print("\n=== 型変換 エイリアステスト ===")
test("型変換", '{"operation":"型変換","settings":{"column":"日付","target_type":"日付型","error_handling":"Nullに変換する","save_method":"上書き保存","task_name":"日付"}}')

print("\n=== 抽出 エイリアステスト ===")
test("抽出: 先頭", '{"operation":"抽出","settings":{"column":"日付","method":"先頭","length":7,"data_type":"テキスト型","save_method":"名前を付けて保存","task_name":"年月"}}')
test("抽出: 末尾", '{"operation":"抽出","settings":{"column":"日付","method":"末尾","length":5,"data_type":"テキスト型","save_method":"名前を付けて保存","task_name":"月日"}}')
test("抽出: 中間", '{"operation":"抽出","settings":{"column":"日時","method":"中間","start":12,"end":13,"data_type":"テキスト型","save_method":"名前を付けて保存","task_name":"時間"}}')

print("\n=== 書式変換 エイリアステスト ===")
test("書式変換", '{"operation":"書式変換","settings":{"column":"リレーション","conversion":"半角英小文字","save_method":"上書き保存","task_name":"リレーション"}}')

print("\n=== 複製 エイリアステスト ===")
test("複製", '{"operation":"複製","settings":{"column":"元カラム","new_column":"コピー"}}')

print("\n=== 除外 エイリアステスト ===")
test("除外", '{"operation":"除外","settings":{"column":"電話番号","scope":"全ての","text":"-","save_method":"上書き保存","task_name":"電話番号"}}')

print("\n=== 分割 エイリアステスト ===")
test("分割", '{"operation":"分割","settings":{"column":"リレーション","delimiter":",","direction":"左","count":3,"new_columns":["A","B","C"]}}')

print("\n=== テンプレート金額カンマ エイリアステスト ===")
test("金額カンマ", '{"operation":"テンプレート 金額をカンマ区切り","settings":{"column":"売上"}}')

print("\n=== テンプレート曜日 エイリアステスト ===")
test("曜日算出", '{"operation":"テンプレート 曜日算出","settings":{"column":"受注日","new_column":"曜日"}}')

print("\n=== テンプレート横→縦 エイリアステスト ===")
test("横→縦変換", '{"operation":"テンプレート 横持ちを縦持ちに変換","settings":{"horizontal_columns":["1月","2月","3月"],"keep_columns":["顧客ID"]}}')

print("\n=== テンプレート3カラム連結 エイリアステスト ===")
test("3カラム連結", '{"operation":"テンプレート 3カラム以上連結","settings":{"columns":["受注ID","モール名","商品ID"],"separator":",","task_name":"複合ID","keep_original":"残す"}}')

print(f"\n{'='*60}")
print(f"  RESULTS: {passed}/{total} passed, {failed} failed")
print(f"{'='*60}")
if errors:
    print("\nErrors:")
    for e in errors:
        print(f"  - {e}")
