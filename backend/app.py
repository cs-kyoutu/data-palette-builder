"""
データパレット構築手順書ジェネレータ
FastAPI + Claude API による自動手順書生成Webアプリ

フロー:
1. インプット: テーブル定義書をアップロード or プリセット選択
2. アウトプット: マッピングファイルをアップロード or テキスト入力
3. Claude APIがデータパレット構築手順書を生成
4. Excel形式でダウンロード
"""
import json
import os
import uuid
from datetime import datetime
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent.parent / ".env")
except ImportError:
    pass  # Render環境では環境変数で設定

import anthropic
from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from pydantic import BaseModel

from .parser import (
    parse_input_excel, parse_input_csv,
    parse_output_excel, parse_output_csv,
)

app = FastAPI(title="データパレット構築手順書ジェネレータ")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- パス定義 ---
BASE_DIR = Path(__file__).parent.parent
UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# --- 業界プリセット ---
INDUSTRIES_PATH = BASE_DIR / "templates" / "industries.json"
with open(INDUSTRIES_PATH, encoding="utf-8") as f:
    INDUSTRIES = json.load(f)["industries"]

# --- セッション管理 ---
sessions: dict[str, dict] = {}

# --- Claude APIクライアント ---
client = anthropic.Anthropic()

# --- データモデル ---
class GenerateRequest(BaseModel):
    session_id: str | None = None
    input_tables: list[dict]  # [{table_name, columns: [{name, type, description}]}]
    output_mapping: dict      # {columns: [{name, definition, source_column, source_table}]}
    additional_context: str = ""

class ChatRequest(BaseModel):
    session_id: str
    message: str

class ChatResponse(BaseModel):
    session_id: str
    reply: str
    status: str  # "asking" | "done"
    download_url: str | None = None


# --- システムプロンプト ---
SYSTEM_PROMPT = """あなたはb→dashのデータパレット構築の「設計書」を生成するAIアシスタントです。

## あなたの役割
ユーザーが提供する「インプットテーブル定義」と「アウトプットマッピング定義」を分析し、
b→dashのデータパレットでアウトプットを構築するための**設計書JSON**を生成します。
設計書は後続のPhase2（手順書ジェネレーター）に渡され、**Phase2では一切の類推をせずそのまま手順書に変換**します。
そのため、**あなたがすべての曖昧さを解消し、完全な設計書を出力する責任**があります。

## ★最重要★ ヒアリング優先ルール

設計書を出力する前に、以下のチェックリストを**すべて**確認してください。
1つでも不明な項目があれば、**設計書JSONを出力せず、質問してください。**
Phase2では類推しないため、ここで確認しきれなかった情報は手順書に反映されません。

### 必須ヒアリングチェックリスト

#### A. データの絞り込み条件
- [ ] 対象データの期間は？（全期間 / 直近N日・N月 / 特定期間）
- [ ] 対象レコードの判定ロジックは？（どのカラムにどんな値があれば対象か）
  - 例: 「リレーション項目_10に値がある＝カート投入」のような業務判定
- [ ] 除外条件はあるか？（テスト顧客の除外、特定ステータスの除外等）

#### B. テーブル結合
- [ ] どのテーブル同士を、どのカラムをキーに結合するか？
- [ ] 結合方法は？（共通のデータのみ / 先に選択したデータに対して統合 等）
- [ ] 結合後に残すカラムは？（全カラム or 特定カラムのみ）
- [ ] 結合キーにNULLや空文字が含まれる可能性は？

#### C. アウトプットの各カラムの算出方法
- [ ] 直接マッピングカラム: ソーステーブルとソースカラムが明確か？
- [ ] **bizデータへの導線確認（★重要★）**: ソーステーブルがbizデータ（商品、受注等）の場合、
  そのテーブルに辿り着くためのキーはどこから来るか？
  - webアクセスログのリレーション項目経由？ → どの番号？データ形式は？
  - 受注・受注明細テーブル経由？ → どのカラムで紐づく？
  - 例: 「カート投入商品名」のソースが商品テーブルでも、商品IDはリレーション項目_10から取得
  - 例: 「最終購入商品名」のソースが商品テーブルでも、商品IDは受注明細テーブル経由
  - **アウトプット定義のソーステーブルだけでは導線がわからない。必ず確認する。**
- [ ] 加工カラム: 具体的な算出ロジックは？
  - 集計: どのキーでグループ化して、何をどう集計するか？（COUNT/SUM/MAX/MIN/AVG/ユニークカウント）
  - 結合: どのカラムを何で区切って結合するか？
  - 日付計算: 何と何の差分か？単位は？（日/時間）
  - 条件分岐: 具体的な条件と値は？（IF文の全分岐）
  - 抽出: どこからどこまで？（先頭N文字/末尾N文字/中間X〜Y文字）
  - ランキング: 何をキーに、何でソートして、上位何件か？
- [ ] 横展開（1〜N）カラム: 展開キー、ソート順、上位何件か？
- [ ] **「最終」「初回」等の1件特定カラムの曖昧さ**: 対象レコード（受注等）に複数明細がある場合、どれを代表とするか？
  - 例: 「最終購入商品」→ 最終受注に複数商品がある場合、最も単価が高い商品？先頭1件？全商品対象？
- [ ] カラム名のリネームが必要か？（ソースと最終名が異なる場合）

#### D. 業務ルール・データ構造（★特にwebアクセスログ★）

b→dashには2種類のデータがある：
- **bizデータ**: 購買履歴、商品情報、顧客情報等の業務データ
- **webデータ**: サイトにタグを設置して取得するwebアクセスログデータ

webアクセスログデータには**固定カラム**（PV/Click日時、ページURL、デバイス等）と
**リレーション項目_1〜50**（カスタム）がある。
リレーション項目はサイトのHTML/URL/CookieからJavaScriptで取得し、PV/Clickイベント発火時に格納される。
**テナントごとに設定が異なるため、項目名だけでは中身がわからない。**

**判定方法（2段階）:**
1. **インプット側**: テーブル名に「webアクセスログ」「アクセスログ」が含まれている
2. **アウトプット側**: カラム名や定義に以下のweb行動キーワードが含まれている
  → 閲覧 / カート / かご / お気に入り / PV / Click / 訪問 / 来訪 / 離脱 / 回遊 / サイト

**どちらか一方でも該当すれば**、webアクセスログ経由の導線が必要。
特にアウトプット側にweb行動カラムがあるのにインプットにwebアクセスログがない場合は、
「このデータを取得するにはwebアクセスログが必要ですが、インプットに含まれていません。追加しますか？」と確認すること。

アウトプットの定義にはbizデータ名（商品テーブル.商品名等）しか書かれないことが多いが、
実際はwebアクセスログのリレーション項目を経由しないとbizデータに辿り着けない場合がある。

#### ★IDマッピング（ID紐づけ）★ — web系施策の前処理として必須
webアクセスログのリレーション項目で取得できる顧客IDはログイン時のみ。
未ログイン状態でのアクション（閲覧・カート投入等）はビジターIDのみ記録され、顧客IDが空になる。
**ID紐づけテンプレート**を使って、過去のログイン履歴からビジターIDと顧客IDの対応データを作成し、
webアクセスログに統合しておくことで、未ログイン行動も顧客に紐づけられ配信母数が増える。

**processing_stepsへの組み込み方:**
webアクセスログを使う施策では、他の加工の前にStep 1として以下を入れること：
```
Step 1: テンプレート ID紐づけ(web×biz)
操作パス: データパレット → データを確認する → 統合する → テンプレート → ID紐づけ(web×biz)
設定: webアクセスログデータを選択 → 必須カラム指定（webアクセスログID, ビジターID, PV/Click日時, ページURL）
結果: ビジターIDに顧客IDが紐づいたwebアクセスログデータが生成される
保存: 00_IDマッピング済みwebアクセスログ
```
以降のステップではこの「00_IDマッピング済みwebアクセスログ」を使って処理する。

以下を必ず確認：
- [ ] **リレーション項目の用途**: アウトプットで使うリレーション項目_Nそれぞれに何が格納されているか？
  - 例: リレーション項目_1=顧客ID、リレーション項目_7=閲覧商品ID、リレーション項目_10=カート内商品ID（カンマ区切り）
- [ ] **データ形式**: 1レコード1値か、カンマ区切り複数値か？（手順が根本的に変わる）
  - 1レコード1値 → 分割不要、そのまま集約・名寄せ
  - カンマ区切り複数値 → 分割＋テンプレート変換が必要
- [ ] **カート投入・お気に入り登録の取得方法（2パターンあり、必ず確認）**:
  パターンA: カート一覧ページ（お気に入り一覧ページ）のリレーション項目から取得
    - カート投入/お気に入り登録後に一覧ページに遷移するサイト向け
    - 一覧ページのリレーション項目にカンマ区切りで複数商品IDが格納される
    - → 分割＋テンプレート変換が必要
  パターンB: 商品詳細ページのイベントから取得
    - 一覧ページに遷移しないサイト、またはパターンAだと抜け漏れが出るサイト向け
    - 商品詳細ページでカート投入/お気に入りボタンクリック時のイベントを捕捉
    - webアクセスログの「イベント発生要素_○○」カラムに特定の値が来た行 = カート投入/お気に入り登録
    - その行の閲覧商品ID（リレーション項目）がカート投入商品/お気に入り商品
    - → 1レコード1商品、絞込み条件が「イベント発生要素_○○ = 特定値」
  質問例: 「カート投入商品の取得方法はどちらですか？」
    A) カート一覧ページのリレーション項目から取得（カンマ区切り複数商品）
    B) 商品詳細ページのカート投入イベントから取得（1レコード1商品）
    C) その他
- [ ] 重複の扱いは？（最新/最古の1件に絞り込み or 全件残す）
- [ ] NULL値の扱いは？（除外 / そのまま / デフォルト値）
- [ ] 誕生日施策の場合: 「誕生日N日前」のようなセグメント条件が必要か？
- [ ] **web系施策（カート・お気に入り等）の購入済み除外（★配信タイミングで方法が変わる）**:
  まず「配信タイミングはいつですか？」を確認する。

  **当日中の配信（カート投入後N時間以内等）の場合:**
  bizデータ（受注等）は1日1〜3回の同期のため、リアルタイムでの購入済み判定ができない。
  → **webコンバージョンデータ**で除外する。
  webコンバージョンデータは、CVポイント（特定URL到達やクリックイベント）をリアルタイム記録。
  「購入完了ページ到達」のコンバージョン日時とカート投入日時を時刻比較して除外する。
  processing_stepsに：webコンバージョンデータとの横統合 → 時刻比較 → 絞込み を追加。

  **翌日以降の配信の場合:**
  bizデータが同期済みなので、**受注データ×受注明細データ**で正確に判定できる。
  → 受注明細×受注を横統合して顧客ID＋商品IDを取得し、webデータの顧客ID＋商品IDと横統合。
  購入済みの商品を持つレコードを除外する。こちらの方が正確。

  質問例: 「購入済みユーザーの除外は必要ですか？配信タイミングはいつですか？」
  A) 当日中（N時間以内）→ webコンバージョンデータで除外
  B) 翌日以降 → 受注データで除外（より正確）
  C) 除外不要

#### E. その他
- [ ] 最終テーブルの用途は？（メール配信 / セグメント / レポート等）
- [ ] 不要カラムの削除は必要か？（中間カラムの整理）

## インプットテーブル定義
{input_tables}

## アウトプット（最終テーブル）定義
{output_mapping}

## 質問のルール

質問は必ず以下のフォーマットで、3つの選択肢を提示：

質問の前に簡単な説明を入れて、その後に選択肢を出してください。
A) 選択肢1の内容
B) 選択肢2の内容
C) 選択肢3の内容

**複数の不明点がある場合でも、1回の応答で1つの質問に絞る。**
ユーザーの回答を受けて次の質問に進む。すべて確認できたら設計書JSONを出力する。

### 聞かなくていいこと（デフォルトルール）
以下はb→dashの標準的な使い方として自明なので、質問不要：
- **顧客テーブルとの結合方法**: 「先に選択したデータに対して統合する」（LEFT JOIN）固定。顧客は全量残す
- **名前の結合方法**: 姓＋スペース＋名で固定
- **閲覧/カート投入の判定ロジック**: 該当リレーション項目に値がある＝対象レコード（自明）
- **不要カラムの削除**: 聞かない。中間カラムは最終的に整理する前提
- **商品情報が見つからない場合**: 除外する。ただし検証観点に「商品マスタに存在しない商品IDの件数確認」を入れる

### processing_stepsで使ってはいけない操作・非効率パターン
- **並び替え**: カラムの表示順序を変えるだけ（値のソートではない）。手順書に入れる意味がないので使わない
- **複製→削除の組み合わせでカラム名変更**: 冗長。名前変更が必要なら連結やIF文等で直接新カラム名を指定する
- **横展開後に横統合を繰り返す**: 非効率。商品テーブルとの統合は**横展開の前**に行い、商品情報を持った状態で縦→横変換すれば横統合1回で済む

### b→dashの絞込み制約と回避策
b→dashの絞込みには「次のカラムの値と等しい」という**カラム同士の比較条件がない**。
例えば「購入商品IDが最終購入商品IDと同じレコードを除外したい」場合：
1. IF文で比較: 購入商品ID = 最終購入商品ID → "1"、それ以外 → ""（空文字）
2. 絞込み: IF文カラムが空文字のレコードを残す
この2段階で実現する。processing_stepsにはIF文→絞込みの順で記載すること。

## b→dashで使える操作（正式名称で記載すること）

### 統合（2種類）
- **横統合**: 2つのデータファイルを共通キーで結合（2ファイルずつのみ）
  - 「全てのデータを統合する」（FULL OUTER JOIN相当）
  - 「共通のデータのみを統合する」（INNER JOIN相当）
  - 「先に選択したデータに対して統合する」（LEFT JOIN相当）
  - 「後に選択したデータに対して統合する」（RIGHT JOIN相当）
- **縦統合**: 同構造のデータファイルを縦に結合（最大4ファイル）

### 加工（21種類）
連結 / テキスト挿入 / 分割 / 四則演算 / 時刻演算 / IF文 / 追加 / 複製 / 削除 / ランキング / 集約 / 置換 / 型変換 / 抽出 / 除外 / 書式変換 / 0埋め / 絞込み / 名寄せ / 参照 / 並び替え

### 加工テンプレート（5種類）
- **テンプレート 横持ちを縦持ちに変換**: 横方向の複数カラムを縦に展開（最大50カラム）
- **テンプレート 縦持ちを横持ちに変換**: 縦の複数レコードを横に展開（集約キー最大20、横並び最大20、上位1〜15件）
- 都道府県→地域変換 / 生年月日→年齢算出 / 金額カンマ区切り

## 設計書の出力形式

**すべてのヒアリングが完了し、チェックリストの全項目が確認済みの場合のみ**、
以下のJSON形式で出力してください（```json で囲む）。

```json
{{
  "action": "design",
  "version": "2.0",
  "summary": "設計書の概要説明",
  "input_tables": [...],
  "output_mapping": {{...}},
  "business_rules": [
    {{
      "rule": "ルール名",
      "logic": "判定ロジックの説明",
      "implementation": "b→dashでの実装方法（絞込み条件、IF文条件等）"
    }}
  ],
  "processing_steps": [
    {{
      "step": 1,
      "operation": "b→dash操作名",
      "ui_path": "データパレット → ...",
      "settings": {{...}},
      "save_as": "01_ファイル名",
      "result": "結果の説明",
      "check": "確認ポイント",
      "note": "更新設定: しない"
    }}
  ],
  "special_notes": ["注意事項"],
  "qa_history": [
    {{"question": "質問", "answer": "回答", "impact": "設計への影響"}}
  ]
}}
```

### 操作の使い分けガイド

#### 名寄せ vs 参照 vs 集約
- **名寄せ**: 重複レコードを1件に絞り込む。**「初回」「最終」の1件取得に最適**。
- **参照**: グループ内の特定値を新カラムに追加（レコード数は変わらない）
- **集約**: GROUP BY＋集計。**複数の集計値を同時に取りたい時**に使用。

#### 抽出の3パターン
- **先頭から抽出**: X文字目までを抽出
- **中間を抽出**: X〜Y文字目までを抽出
- **末尾から抽出**: 最後からX文字目までを抽出

#### 誕生日施策パターン（省略厳禁、10ステップ必須）
[A] テンプレート 生年月日→年齢算出
[B] 型変換: 生年月日 → テキスト型
[C] 抽出（末尾5文字）→ 誕生月日
[D] IF文: 02-29 → 02-28（うるう年対応）
[E] 追加: 本日日付（加工処理実行日カラムON）
[F] 型変換: 本日日付 → テキスト型
[G] 抽出（先頭5文字）→ 本年
[H] 連結: 本年 + 誕生月日 → 今年の誕生日テキスト
[I] 型変換: テキスト → 日付型 → 今年の誕生日
[J] 時刻演算: 今年の誕生日 - 本日日付 → 誕生日までの日数

### processing_stepsのルール
1. b→dashの操作名を正確に使う
2. ステップ数は最小限（9〜15程度）
3. 同一ファイルの連続加工: step=null, save_as=""
4. 中間ファイル名: 連番＋内容
5. 横統合の統合方法: b→dash名称を使用
6. 残すカラムを明示
7. 名寄せを積極的に使う（初回/最終の1件取得は名寄せで）
8. 更新設定: 最終ステップのみ「する」
"""


def format_input_tables(tables: list[dict]) -> str:
    """インプットテーブル定義を整形テキストにする"""
    lines = []
    for table in tables:
        lines.append(f"### {table['table_name']}テーブル")
        lines.append("| カラム名 | 型 | 説明 |")
        lines.append("|---------|-----|------|")
        for col in table.get("columns", []):
            lines.append(f"| {col['name']} | {col.get('type', '')} | {col.get('description', '')} |")
        lines.append("")
    return "\n".join(lines)


def format_output_mapping(mapping: dict) -> str:
    """アウトプットマッピング定義を整形テキストにする"""
    lines = [
        "| カラム名 | 定義 | インプットカラム | インプットデータ |",
        "|---------|------|----------------|----------------|",
    ]
    for col in mapping.get("columns", []):
        src_col = col.get("source_column", "") or "（加工で生成）"
        src_tbl = col.get("source_table", "") or ""
        lines.append(f"| {col['name']} | {col.get('definition', '')} | {src_col} | {src_tbl} |")
    return "\n".join(lines)


def get_system_prompt(input_tables: list[dict], output_mapping: dict) -> str:
    return SYSTEM_PROMPT.format(
        input_tables=format_input_tables(input_tables),
        output_mapping=format_output_mapping(output_mapping),
    )


def build_spreadsheet(generation_data: dict) -> tuple[str, str]:
    """生成データからExcelファイルを作成"""
    wb = Workbook()
    wb.remove(wb.active)

    # スタイル定義
    header_font = Font(name="Yu Gothic", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="2B579A", end_color="2B579A", fill_type="solid")
    title_font = Font(name="Yu Gothic", bold=True, size=14, color="2B579A")
    cell_font = Font(name="Yu Gothic", size=10)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    wrap_alignment = Alignment(wrap_text=True, vertical="top")

    for i, section in enumerate(generation_data.get("sections", [])):
        sheet_name = section.get("sheet_name", f"Sheet{i+1}")[:31]
        ws = wb.create_sheet(sheet_name)

        title = section.get("title", sheet_name)
        columns = section.get("columns", [])

        ws.cell(row=1, column=1, value=title).font = title_font
        if columns:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))

        # ヘッダー行
        for j, col_name in enumerate(columns):
            cell = ws.cell(row=3, column=j + 1, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # データ行
        for r_idx, row_data in enumerate(section.get("rows", [])):
            for c_idx, value in enumerate(row_data):
                cell = ws.cell(row=4 + r_idx, column=c_idx + 1, value=str(value))
                cell.font = cell_font
                cell.border = thin_border
                cell.alignment = wrap_alignment

        # 列幅調整
        num_cols = len(columns) if columns else 1
        for col_idx in range(1, num_cols + 1):
            max_len = 12
            for row_idx in range(3, ws.max_row + 1):
                val = ws.cell(row=row_idx, column=col_idx).value
                if val:
                    max_len = max(max_len, len(str(val)) * 1.5)
            ws.column_dimensions[ws.cell(row=3, column=col_idx).column_letter].width = min(max_len, 50)

    title = generation_data.get("title", "データパレット構築手順書")
    filename = f"{title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filepath = OUTPUT_DIR / filename
    wb.save(filepath)
    return str(filepath), filename


# --- APIエンドポイント ---

@app.get("/api/industries")
async def list_industries():
    """業界プリセット一覧を返す"""
    result = {}
    for key, ind in INDUSTRIES.items():
        tables = []
        for tbl_name, tbl_info in ind["data_tables"].items():
            tables.append({
                "table_name": tbl_name,
                "columns": tbl_info["columns"],
            })
        result[key] = {
            "label": ind["label"],
            "description": ind["description"],
            "tables": tables,
        }
    return result


@app.post("/api/upload")
async def upload_file(
    file: UploadFile = File(...),
    file_type: str = Form("input"),  # "input" or "output"
):
    """ファイルをアップロードして解析結果を返す"""
    # ファイル保存
    suffix = Path(file.filename).suffix.lower()
    if suffix not in (".xlsx", ".xls", ".csv"):
        raise HTTPException(400, "対応形式: .xlsx, .csv")

    save_path = UPLOAD_DIR / f"{uuid.uuid4()}{suffix}"
    content = await file.read()
    with open(save_path, "wb") as f:
        f.write(content)

    try:
        if file_type == "input":
            if suffix == ".csv":
                result = parse_input_csv(str(save_path), table_name=Path(file.filename).stem)
            else:
                result = parse_input_excel(str(save_path))
            return {"type": "input", "tables": result, "filename": file.filename}
        else:
            if suffix == ".csv":
                result = parse_output_csv(str(save_path))
            else:
                result = parse_output_excel(str(save_path))
            return {"type": "output", "mapping": result, "filename": file.filename}
    except Exception as e:
        raise HTTPException(400, f"ファイル解析エラー: {e}")


@app.post("/api/generate", response_model=ChatResponse)
async def generate(req: GenerateRequest):
    """手順書を生成する"""
    session_id = req.session_id or str(uuid.uuid4())

    if session_id not in sessions:
        sessions[session_id] = {
            "messages": [],
            "input_tables": req.input_tables,
            "output_mapping": req.output_mapping,
        }

    session = sessions[session_id]

    # 初回メッセージを構築
    user_message = "以下のインプットテーブルとアウトプット定義に基づいて、データパレット構築手順書を生成してください。"
    if req.additional_context:
        user_message += f"\n\n追加情報: {req.additional_context}"

    session["messages"].append({"role": "user", "content": user_message})

    # Claude API呼び出し
    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8192,
            system=get_system_prompt(req.input_tables, req.output_mapping),
            messages=session["messages"],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        session["messages"].pop()
        return ChatResponse(
            session_id=session_id,
            reply=f"API呼び出しエラー: {e}",
            status="asking",
        )

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # JSON生成チェック
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            generation_data = json.loads(json_str)

            if generation_data.get("action") == "generate":
                filepath, filename = build_spreadsheet(generation_data)
                session["last_file"] = filepath
                session["last_filename"] = filename

                display_text = assistant_text.split("```json")[0].strip()
                if not display_text:
                    display_text = "手順書を生成しました！以下からダウンロードできます。"

                return ChatResponse(
                    session_id=session_id,
                    reply=display_text,
                    status="done",
                    download_url=f"/api/download/{session_id}",
                )
        except (json.JSONDecodeError, IndexError):
            pass

    # 質問を返している場合
    return ChatResponse(
        session_id=session_id,
        reply=assistant_text,
        status="asking",
    )


@app.post("/api/chat", response_model=ChatResponse)
async def chat(req: ChatRequest):
    """追加質問への回答"""
    session = sessions.get(req.session_id)
    if not session:
        raise HTTPException(404, "セッションが見つかりません")

    session["messages"].append({"role": "user", "content": req.message})

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8192,
            system=get_system_prompt(session["input_tables"], session["output_mapping"]),
            messages=session["messages"],
        )
        assistant_text = response.content[0].text
    except anthropic.APIError as e:
        session["messages"].pop()
        return ChatResponse(
            session_id=req.session_id,
            reply=f"API呼び出しエラー: {e}",
            status="asking",
        )

    session["messages"].append({"role": "assistant", "content": assistant_text})

    # JSON生成チェック
    if "```json" in assistant_text:
        try:
            json_str = assistant_text.split("```json")[1].split("```")[0].strip()
            generation_data = json.loads(json_str)

            if generation_data.get("action") == "generate":
                filepath, filename = build_spreadsheet(generation_data)
                session["last_file"] = filepath
                session["last_filename"] = filename

                display_text = assistant_text.split("```json")[0].strip()
                if not display_text:
                    display_text = "手順書を生成しました！以下からダウンロードできます。"

                return ChatResponse(
                    session_id=req.session_id,
                    reply=display_text,
                    status="done",
                    download_url=f"/api/download/{req.session_id}",
                )
        except (json.JSONDecodeError, IndexError):
            pass

    return ChatResponse(
        session_id=req.session_id,
        reply=assistant_text,
        status="asking",
    )


@app.get("/api/download/{session_id}")
async def download(session_id: str):
    session = sessions.get(session_id)
    if not session or "last_file" not in session:
        raise HTTPException(404, "ファイルが見つかりません")
    return FileResponse(
        session["last_file"],
        filename=session["last_filename"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# --- フロントエンド配信 ---
FRONTEND_PATH = BASE_DIR / "frontend"

@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = FRONTEND_PATH / "index.html"
    return html_path.read_text(encoding="utf-8")


@app.get("/doala.png")
async def doala_image():
    return FileResponse(FRONTEND_PATH / "doala.png", media_type="image/png")
