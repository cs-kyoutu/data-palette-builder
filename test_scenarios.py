import json, os, sys, shutil
sys.stdout.reconfigure(encoding='utf-8')
os.environ['ANTHROPIC_API_KEY'] = 'REDACTED_API_KEY'
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from backend.app import get_system_prompt
from backend.app_phase2 import format_design_document, SYSTEM_PROMPT_PHASE2, build_spreadsheet
import anthropic

client = anthropic.Anthropic()

SCENARIOS = {
    "初回購入フォローメール": {
        "input_tables": [
            {"table_name": "顧客", "columns": [
                {"name": "顧客ID", "type": "テキスト", "description": "一意の顧客識別子"},
                {"name": "メールアドレス(メイン)", "type": "テキスト", "description": "メインのメールアドレス"},
                {"name": "メルマガ配信許可フラグ", "type": "テキスト", "description": "メルマガ配信許可フラグ"},
                {"name": "姓", "type": "テキスト", "description": "顧客の姓"},
                {"name": "名", "type": "テキスト", "description": "顧客の名"},
                {"name": "会員登録日時", "type": "日時", "description": "会員登録日時"},
            ]},
            {"table_name": "受注", "columns": [
                {"name": "受注ID", "type": "テキスト", "description": "受注の一意識別子"},
                {"name": "顧客ID", "type": "テキスト", "description": "顧客ID"},
                {"name": "受注日時", "type": "日時", "description": "受注日時"},
                {"name": "受注金額", "type": "数値", "description": "受注金額（税込）"},
            ]},
            {"table_name": "受注明細", "columns": [
                {"name": "受注ID", "type": "テキスト", "description": "受注ID"},
                {"name": "商品ID", "type": "テキスト", "description": "商品ID"},
                {"name": "商品名", "type": "テキスト", "description": "商品名"},
                {"name": "数量", "type": "数値", "description": "数量"},
                {"name": "単価", "type": "数値", "description": "単価"},
            ]},
            {"table_name": "商品", "columns": [
                {"name": "商品ID", "type": "テキスト", "description": "商品ID"},
                {"name": "商品名", "type": "テキスト", "description": "商品名"},
                {"name": "カテゴリ", "type": "テキスト", "description": "商品カテゴリ"},
                {"name": "商品画像URL_品番単位(差込用)", "type": "テキスト", "description": "商品画像URL"},
                {"name": "商品詳細ページURL(差込用)", "type": "テキスト", "description": "商品詳細ページURL"},
                {"name": "現在価格(差込用)", "type": "テキスト", "description": "現在の販売価格"},
            ]},
        ],
        "output_mapping": {"columns": [
            {"name": "顧客ID", "definition": "-", "source_column": "顧客ID", "source_table": "顧客"},
            {"name": "メールアドレス", "definition": "-", "source_column": "メールアドレス(メイン)", "source_table": "顧客"},
            {"name": "メルマガ配信許可フラグ", "definition": "-", "source_column": "メルマガ配信許可フラグ", "source_table": "顧客"},
            {"name": "名前", "definition": "-", "source_column": None, "source_table": "顧客", "derivation": "姓と名をスペースで結合"},
            {"name": "初回購入日", "definition": "初めて購入した日付", "source_column": None, "source_table": None, "derivation": "顧客IDごとの受注日時の最小値"},
            {"name": "初回購入金額", "definition": "初めて購入した金額", "source_column": None, "source_table": None, "derivation": "初回受注の受注金額"},
            {"name": "初回購入商品名1〜3", "definition": "初回購入した商品名", "source_column": "商品名", "source_table": "受注明細", "derivation": "初回受注の商品を上位3件横展開"},
            {"name": "初回購入商品画像URL1〜3", "definition": "初回購入した商品画像", "source_column": "商品画像URL_品番単位(差込用)", "source_table": "商品", "derivation": "初回購入商品に対応する画像URL"},
        ]},
    },
    "休眠顧客掘り起こし": {
        "input_tables": [
            {"table_name": "顧客", "columns": [
                {"name": "顧客ID", "type": "テキスト", "description": "一意の顧客識別子"},
                {"name": "メールアドレス(メイン)", "type": "テキスト", "description": "メインのメールアドレス"},
                {"name": "メルマガ配信許可フラグ", "type": "テキスト", "description": "メルマガ配信許可フラグ"},
                {"name": "姓", "type": "テキスト", "description": "顧客の姓"},
                {"name": "名", "type": "テキスト", "description": "顧客の名"},
                {"name": "会員ランク", "type": "テキスト", "description": "会員ランク名"},
            ]},
            {"table_name": "受注", "columns": [
                {"name": "受注ID", "type": "テキスト", "description": "受注の一意識別子"},
                {"name": "顧客ID", "type": "テキスト", "description": "顧客ID"},
                {"name": "受注日時", "type": "日時", "description": "受注日時"},
                {"name": "受注金額", "type": "数値", "description": "受注金額（税込）"},
            ]},
            {"table_name": "受注明細", "columns": [
                {"name": "受注ID", "type": "テキスト", "description": "受注ID"},
                {"name": "商品ID", "type": "テキスト", "description": "商品ID"},
                {"name": "商品名", "type": "テキスト", "description": "商品名"},
            ]},
            {"table_name": "商品", "columns": [
                {"name": "商品ID", "type": "テキスト", "description": "商品ID"},
                {"name": "商品名", "type": "テキスト", "description": "商品名"},
                {"name": "商品画像URL_品番単位(差込用)", "type": "テキスト", "description": "商品画像URL"},
                {"name": "現在価格(差込用)", "type": "テキスト", "description": "現在の販売価格"},
            ]},
        ],
        "output_mapping": {"columns": [
            {"name": "顧客ID", "definition": "-", "source_column": "顧客ID", "source_table": "顧客"},
            {"name": "メールアドレス", "definition": "-", "source_column": "メールアドレス(メイン)", "source_table": "顧客"},
            {"name": "メルマガ配信許可フラグ", "definition": "-", "source_column": "メルマガ配信許可フラグ", "source_table": "顧客"},
            {"name": "名前", "definition": "-", "source_column": None, "source_table": "顧客", "derivation": "姓と名をスペースで結合"},
            {"name": "会員ランク", "definition": "-", "source_column": "会員ランク", "source_table": "顧客"},
            {"name": "最終購入日", "definition": "最後に購入した日付", "source_column": None, "source_table": None, "derivation": "顧客IDごとの受注日時の最大値"},
            {"name": "最終購入からの経過日数", "definition": "最後の購入からの日数", "source_column": None, "source_table": None, "derivation": "現在日時 - 最終購入日（日単位）"},
            {"name": "累計購入回数", "definition": "これまでの購入回数", "source_column": None, "source_table": None, "derivation": "顧客IDごとの受注IDのユニークカウント"},
            {"name": "累計購入金額", "definition": "これまでの購入金額合計", "source_column": None, "source_table": None, "derivation": "顧客IDごとの受注金額の合計"},
            {"name": "最終購入商品名1〜3", "definition": "最後に買った商品名", "source_column": "商品名", "source_table": "受注明細", "derivation": "最終受注の商品を上位3件横展開"},
            {"name": "最終購入商品画像URL1〜3", "definition": "最後に買った商品の画像", "source_column": "商品画像URL_品番単位(差込用)", "source_table": "商品", "derivation": "最終購入商品に対応する画像URL"},
            {"name": "最終購入商品価格1〜3", "definition": "最後に買った商品の価格", "source_column": "現在価格(差込用)", "source_table": "商品", "derivation": "最終購入商品に対応する現在価格"},
        ]},
    },
    "誕生日クーポンメール": {
        "input_tables": [
            {"table_name": "顧客", "columns": [
                {"name": "顧客ID", "type": "テキスト", "description": "一意の顧客識別子"},
                {"name": "メールアドレス(メイン)", "type": "テキスト", "description": "メインのメールアドレス"},
                {"name": "メルマガ配信許可フラグ", "type": "テキスト", "description": "メルマガ配信許可フラグ"},
                {"name": "姓", "type": "テキスト", "description": "顧客の姓"},
                {"name": "名", "type": "テキスト", "description": "顧客の名"},
                {"name": "生年月日", "type": "日付", "description": "生年月日"},
                {"name": "会員ランク", "type": "テキスト", "description": "会員ランク名"},
                {"name": "保持ポイント", "type": "数値", "description": "現在保持しているポイント数"},
            ]},
            {"table_name": "受注", "columns": [
                {"name": "受注ID", "type": "テキスト", "description": "受注の一意識別子"},
                {"name": "顧客ID", "type": "テキスト", "description": "顧客ID"},
                {"name": "受注日時", "type": "日時", "description": "受注日時"},
                {"name": "受注金額", "type": "数値", "description": "受注金額（税込）"},
            ]},
        ],
        "output_mapping": {"columns": [
            {"name": "顧客ID", "definition": "-", "source_column": "顧客ID", "source_table": "顧客"},
            {"name": "メールアドレス", "definition": "-", "source_column": "メールアドレス(メイン)", "source_table": "顧客"},
            {"name": "メルマガ配信許可フラグ", "definition": "-", "source_column": "メルマガ配信許可フラグ", "source_table": "顧客"},
            {"name": "名前", "definition": "-", "source_column": None, "source_table": "顧客", "derivation": "姓と名をスペースで結合"},
            {"name": "誕生日", "definition": "顧客の誕生日", "source_column": "生年月日", "source_table": "顧客"},
            {"name": "誕生月", "definition": "誕生月（MM形式）", "source_column": None, "source_table": None, "derivation": "生年月日から月のみ抽出"},
            {"name": "年齢", "definition": "現在の年齢", "source_column": None, "source_table": None, "derivation": "生年月日から年齢算出（テンプレート使用）"},
            {"name": "会員ランク", "definition": "-", "source_column": "会員ランク", "source_table": "顧客"},
            {"name": "保持ポイント", "definition": "-", "source_column": "保持ポイント", "source_table": "顧客"},
            {"name": "累計購入回数", "definition": "これまでの購入回数", "source_column": None, "source_table": None, "derivation": "顧客IDごとの受注IDのユニークカウント"},
            {"name": "累計購入金額", "definition": "これまでの購入金額合計", "source_column": None, "source_table": None, "derivation": "顧客IDごとの受注金額の合計"},
            {"name": "クーポン種別", "definition": "会員ランク別のクーポン種別", "source_column": None, "source_table": None, "derivation": "会員ランクがゴールド以上→20%OFF、シルバー→15%OFF、その他→10%OFF"},
        ]},
    },
}

def run_scenario(name, scenario):
    print(f"\n{'='*60}")
    print(f"  {name}")
    print(f"{'='*60}")
    system_prompt = get_system_prompt(scenario["input_tables"], scenario["output_mapping"])
    print(f"Phase1: {len(system_prompt)}c...")
    resp1 = client.messages.create(
        model='claude-sonnet-4-20250514', max_tokens=8192, system=system_prompt,
        messages=[{'role': 'user', 'content': f'{name}用のデータパレット設計書を生成してください。processing_stepsをb-dash操作レベルで。質問不要。'}],
    )
    text1 = resp1.content[0].text
    print(f"Phase1 resp: {len(text1)}c, stop: {resp1.stop_reason}")
    if '```json' not in text1:
        print(f"FAIL: No JSON. Text: {text1[:300]}")
        return
    design_doc = json.loads(text1.split('```json')[1].split('```')[0].strip())
    steps = design_doc.get('processing_steps', [])
    print(f"Steps: {len(steps)}")
    for s in steps:
        sn = s.get('step', '')
        print(f"  {'Step '+str(sn) if sn else '  (cont)'}: {s.get('operation','')}")

    design_text = format_design_document(design_doc)
    system_prompt2 = SYSTEM_PROMPT_PHASE2.format(design_document=design_text)
    print(f"Phase2: {len(system_prompt2)}c...")
    resp2 = client.messages.create(
        model='claude-sonnet-4-20250514', max_tokens=16384, system=system_prompt2,
        messages=[{'role': 'user', 'content': 'Generate procedure doc. Same structure as sample. JSON only.'}],
    )
    text2 = resp2.content[0].text
    print(f"Phase2 resp: {len(text2)}c, stop: {resp2.stop_reason}")
    if '```json' not in text2:
        print(f"FAIL: No JSON. Text: {text2[:300]}")
        return
    proc_data = json.loads(text2.split('```json')[1].split('```')[0].strip())
    if proc_data.get('action') == 'generate':
        fp, fn = build_spreadsheet(proc_data)
        safe_name = name.replace('/', '_')
        dest = f'C:\\Users\\yugam\\Downloads\\手順書_{safe_name}.xlsx'
        shutil.copy(fp, dest)
        print(f"SUCCESS -> {dest}")
        for s in proc_data['sections']:
            print(f"  {s['sheet_name']}: {len(s.get('rows',[]))} rows")

if __name__ == '__main__':
    target = sys.argv[1] if len(sys.argv) > 1 else 'all'
    if target == 'all':
        for name, sc in SCENARIOS.items():
            run_scenario(name, sc)
    elif target in SCENARIOS:
        run_scenario(target, SCENARIOS[target])
    else:
        print("Available:", list(SCENARIOS.keys()))
