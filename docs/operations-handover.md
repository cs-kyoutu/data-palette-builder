# データパレット構築 設計書・手順書ジェネレータ（β版）運用引継ぎ資料

## 1. サービス概要

b→dashのデータパレット構築手順書を、AIとの対話で自動生成するWebアプリケーション。

**URL**: https://data-palette-builder.onrender.com

**3フェーズパイプライン:**
```
Phase1: 施策相談（①誰に ②何を ③いつ ④誰には送らないか を確定）
  ↓
Phase2: テーブル定義の整理（実テーブルとのマッピング）
  ↓
Phase3: 手順書生成（b→dash操作手順をExcel出力）
```

---

## 2. 技術スタック

| レイヤー | 技術 | バージョン |
|---|---|---|
| **バックエンド** | Python + FastAPI | Python 3.11.9 / FastAPI 0.115.0 |
| **AI** | Anthropic Claude API (Sonnet 4) | anthropic 0.42.0 |
| **フロントエンド** | Pure HTML/CSS/JS（SPA、フレームワークなし） | - |
| **データベース** | Supabase (PostgreSQL) | supabase-py 2.15.2 |
| **ホスティング** | Render (Free tier) | - |
| **Excel生成** | openpyxl | 3.1.5 |
| **認証** | Bearer Token (HTTPBearer) | - |
| **レート制限** | slowapi | 0.1.9 |

---

## 3. 外部サービス・アカウント

### 3-1. Render（ホスティング）
- **ダッシュボード**: https://dashboard.render.com
- **サービス名**: data-palette-builder
- **プラン**: Free tier
- **デプロイ**: GitHubのmasterブランチへのpushで自動デプロイ
- **注意**: 無料枠は15分放置でスリープ。復帰に30秒〜1分かかる

### 3-2. Anthropic（Claude API）
- **コンソール**: https://console.anthropic.com
- **使用モデル**: claude-sonnet-4-20250514（Phase1/Phase3）、claude-sonnet-4-5-20250929（Phase2）
- **コスト**: 約60円/施策（3フェーズフル完遂時）
- **レート制限**: generate/start系 10req/min、chat系 20req/min

### 3-3. Supabase（データベース）
- **ダッシュボード**: https://supabase.com/dashboard
- **プロジェクト名**: data-palette-builder
- **リージョン**: Northeast Asia (Tokyo)
- **プラン**: Free tier（500MB）
- **テーブル**: 下記3つ

### 3-4. GitHub（ソースコード）
- **リポジトリ**: https://github.com/y-miyazaki-biginner/data-palette-builder
- **ブランチ**: master（本番デプロイ対象）

---

## 4. 環境変数

Renderダッシュボード → Environment で管理。

| 変数名 | 用途 | 設定場所 |
|---|---|---|
| `ANTHROPIC_API_KEY` | Claude API認証キー | Render（手動） |
| `APP_AUTH_TOKEN` | アプリのBearer Token認証 | Render（手動） |
| `CORS_ORIGIN` | CORS許可オリジン | Render（`https://data-palette-builder.onrender.com`） |
| `SUPABASE_URL` | Supabase接続URL | Render（手動） |
| `SUPABASE_KEY` | Supabase service_roleキー | Render（手動） |
| `PYTHON_VERSION` | Pythonバージョン | render.yaml（`3.11.9`固定） |

**ローカル開発時**: `.env` ファイルに `ANTHROPIC_API_KEY`, `SUPABASE_URL`, `SUPABASE_KEY` を記載。`APP_AUTH_TOKEN` を未設定にすると認証スキップ。

---

## 5. データベース構造（Supabase）

### 5-1. knowledge テーブル
ユーザーのFBや施策相談結果を蓄積。AIのナレッジとして次回以降の精度向上に活用。

| カラム | 型 | 用途 |
|---|---|---|
| id | TEXT PK | UUID短縮8桁 |
| type | TEXT | consultation / organization / success / correction |
| strategy_name | TEXT | 施策名 |
| strategy_summary | TEXT | 施策概要 |
| output_columns | TEXT[] | アウトプットカラム名一覧 |
| input_table_names | TEXT[] | 使用テーブル名一覧 |
| correction | TEXT | 修正FB内容 |
| summary | TEXT | 設計書サマリ |
| processing_steps | JSONB | 設計ステップ一覧 |
| processing_groups | JSONB | 処理グループ |
| raw_data | JSONB | 全データ（後方互換） |
| created_at | TIMESTAMPTZ | 作成日時 |

### 5-2. industries テーブル
業界プリセット（テーブル定義のセット）。

| カラム | 型 | 用途 |
|---|---|---|
| id | TEXT PK | 業界キー（例: ec_専業） |
| label | TEXT | 表示名 |
| description | TEXT | 説明 |
| tables | JSONB | テーブル定義の配列 |
| created_at | TIMESTAMPTZ | 作成日時 |
| updated_at | TIMESTAMPTZ | 更新日時 |

### 5-3. strategy_templates テーブル
施策テンプレート（カート落ち、誕生日等のパターン）。

| カラム | 型 | 用途 |
|---|---|---|
| id | TEXT PK | テンプレートID |
| name | TEXT | テンプレート名 |
| description | TEXT | 説明 |
| who / what / when | TEXT | 4観点のデフォルト値 |
| exclude | JSONB | 除外条件リスト |
| ask_user | JSONB | AIが必ずユーザーに聞く項目 |
| processing_notes | JSONB | 処理上の注意点 |
| output_columns | JSONB | デフォルトアウトプットカラム |
| input_tables | JSONB | 必要テーブルリスト |

---

## 6. ファイル構成

```
data-palette-builder/
├── backend/
│   ├── app.py              ← メインアプリ（2002行）全APIエンドポイント
│   ├── parser.py            ← Excel/CSVパーサー
│   ├── template_engine.py   ← Phase3設計→手順書変換
│   ├── excel_builder.py     ← 手順書Excelフォーマット
│   ├── uploads/             ← アップロード一時保存（24h自動削除）
│   └── output/              ← 生成Excel一時保存（24h自動削除）
├── frontend/
│   ├── index.html           ← SPA（3008行）全UI
│   ├── bdash_hakase.png     ← AIアバター画像
│   └── *.svg                ← ロゴ等
├── skills/
│   ├── design_patterns.md   ← 共通設計ルール（AIプロンプト）
│   ├── operation_masters.yaml ← b→dash操作テンプレート
│   ├── hearing_defaults.md  ← Phase3ヒアリングルール
│   ├── purchase_exclusion.md ← 購入除外ロジック
│   ├── web_data.md          ← webアクセスログルール
│   ├── birthday_pattern.md  ← 誕生日施策パターン
│   ├── strategy_templates.json ← フォールバック用（DB優先）
│   ├── knowledge_base.json  ← フォールバック用（DB優先）
│   └── patterns/            ← 施策別パターン（6ファイル）
│       ├── cart_and_purchase.md
│       ├── web_access.md
│       ├── calculation_and_format.md
│       ├── advanced_integration.md
│       ├── text_and_cleanup.md
│       └── operational.md
├── templates/
│   └── industries.json      ← フォールバック用（DB優先）
├── requirements.txt
├── render.yaml
├── start.py                 ← ローカル起動スクリプト
└── .env                     ← ローカル環境変数（gitignore済み）
```

---

## 7. APIエンドポイント一覧

全エンドポイントにBearer Token認証あり（`/`, `/bdash_hakase.png`, `/favicon.png` を除く）。

### Phase1: 施策相談
| メソッド | パス | 用途 | レート制限 |
|---|---|---|---|
| POST | /api/consultation/start | 相談セッション開始 | 10/min |
| POST | /api/chat | チャット続行 | 20/min |
| POST | /api/consultation/apply | （後方互換）直接手順書生成 | - |

### Phase2: テーブル定義整理
| メソッド | パス | 用途 | レート制限 |
|---|---|---|---|
| POST | /api/organization/start | 整理セッション開始 | 10/min |
| POST | /api/organization/chat | マッピングQ&A | 20/min |
| POST | /api/organization/update-tables | テーブル差し替え | - |
| POST | /api/organization/finalize | 確定→手順書生成へ | 10/min |

### Phase3: 手順書生成
| メソッド | パス | 用途 | レート制限 |
|---|---|---|---|
| POST | /api/generate | 手順書生成（2段階） | 10/min |
| GET | /api/download/{id} | Excel手順書DL | - |
| GET | /api/download-design/{id} | 設計JSON DL | - |
| POST | /api/feedback | FB送信 | - |

### データ管理
| メソッド | パス | 用途 |
|---|---|---|
| GET | /api/industries | 業界プリセット一覧 |
| POST/PUT/DELETE | /api/industries/{id} | 業界CRUD |
| GET | /api/strategy-templates | 施策テンプレート一覧 |
| POST/PUT/DELETE | /api/strategy-templates/{id} | テンプレートCRUD |
| GET | /api/knowledge | ナレッジ一覧 |
| GET/PUT/DELETE | /api/knowledge/{id} | ナレッジCRUD |
| POST | /api/upload | ファイルアップロード |

---

## 8. セキュリティ対策

| 対策 | 実装 |
|---|---|
| **認証** | Bearer Token（APP_AUTH_TOKEN環境変数） |
| **CORS** | 自ドメインのみ許可 |
| **レート制限** | slowapi（IPベース、10〜20req/min） |
| **XSS** | escapeHtml()で全動的コンテンツをサニタイズ |
| **APIキー保護** | サーバーサイドのみ。フロントに露出なし |
| **エラーメッセージ** | 内部例外を非公開化 |
| **ファイルクリーンアップ** | 24時間後に自動削除 |
| **セッション管理** | UUIDベース、24時間後に自動破棄 |
| **依存関係** | 全パッケージバージョン固定 |

---

## 9. ローカル開発

```bash
# リポジトリクローン
git clone https://github.com/y-miyazaki-biginner/data-palette-builder.git
cd data-palette-builder

# .envファイル作成
cp .env.example .env
# ANTHROPIC_API_KEY, SUPABASE_URL, SUPABASE_KEY を記入

# 依存インストール
pip install -r requirements.txt

# 起動
python start.py
# → http://localhost:8002 で起動
# APP_AUTH_TOKEN未設定のため認証スキップ
```

---

## 10. デプロイ手順

```bash
# コード修正後
git add -A
git commit -m "変更内容"
git push origin master
# → Renderが自動デプロイ（2〜5分）
```

**手動デプロイ**: Renderダッシュボード → Manual Deploy → Deploy latest commit

---

## 11. トラブルシューティング

| 症状 | 原因 | 対処 |
|---|---|---|
| サイトが表示されない | Renderスリープ中 | 30秒〜1分待つ |
| APIが401エラー | Bearer Token不一致 | Renderの`APP_AUTH_TOKEN`を確認 |
| 手順書DLできない | 認証ヘッダ不足（旧バグ） | fetchラッパーでトークン自動付与済み |
| ナレッジが消える | Supabase未接続 | `SUPABASE_URL`/`SUPABASE_KEY`を確認 |
| AIが応答しない | Anthropic APIエラー/レート制限 | リトライボタンで再試行 |
| プリセットが表示されない | industriesDataの取得失敗 | ブラウザF12でconsoleエラー確認 |

---

## 12. コスト

| サービス | 月額 |
|---|---|
| Render Free tier | ¥0 |
| Supabase Free tier | ¥0 |
| Claude API | 使用量に応じて（約60円/施策） |

---

## 13. 今後の改善候補

- Prompt Cachingによるコスト40%削減
- Supabase有料化 or Render Starter($7/月)でスリープ解消
- ユーザー認証（個人ごとのナレッジ分離）
- 施策テンプレートの追加（人材/公共インフラ向け）
- Phase2→Phase3の自動遷移改善
