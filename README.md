# AIEvaluateToDoList

Excelの課題管理表を行ごとにAIへ渡し、次の6項目を推論してExcelへ追記します。既存の内容は残したまま、日時とモデル名を明記して追記します（書式は極力保持）。

- AI_工数（例: 0.5人日）
- AI_Level（低/中/高）
- AI_NextAction（次にやるべきこと）
- AI_Advice（進め方の助言）
- AI_Evaluation（総合評価）
- AI_Caution（注意点の配列）

「ステータス」が「完了」の行はスキップします。

## クイックスタート（Windows PowerShell）

1) 依存関係のインストール

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2) OpenAI APIキーの設定（どちらか）

```powershell
$env:OPENAI_API_KEY = "sk-..."
# または .env を作成して OPENAI_API_KEY=sk-... を記入（自動読込）
```

3) 実行例

- 新規作成（既定: 同名があれば連番で新規に保存）

```powershell
python -m src.ai_enricher --input .\kadai_kanri_simple.xlsx --json .\ai_results.json --model gpt-5 --write-mode new
```

- 上書き（--output未指定なら入力ファイルに上書きします）

```powershell
python -m src.ai_enricher --input .\kadai_kanri_simple.xlsx --write-mode overwrite
```

- Dry-run（APIを呼ばずに挙動確認）

```powershell
python -m src.ai_enricher --input .\kadai_kanri_simple.xlsx --dry-run --write-mode new
```

- 最小実行（既定値のみ）

```powershell
python -m src.ai_enricher
```

- JSON出力なし（Excelのみ）

```powershell
python -m src.ai_enricher --input .\kadai_kanri_simple.xlsx --json '' --write-mode new
```

- モデルを指定（例: gpt-4o-mini）

```powershell
python -m src.ai_enricher --input .\kadai_kanri_simple.xlsx --model gpt-4o-mini --write-mode new
```

- 出力先を明示して新規作成

```powershell
python -m src.ai_enricher --input .\kadai_kanri_simple.xlsx --output .\out\my_tasks_with_ai.xlsx --write-mode new
```

- 別パスへ上書き保存

```powershell
python -m src.ai_enricher --input .\kadai_kanri_simple.xlsx --output .\out\my_tasks.xlsx --write-mode overwrite
```

## 入力と出力の仕様

- 入力列（見出し行は1行目）
	- `No.`, `課題`, `ステータス`, `人員`, `担当者`, `現状把握`, `備考・コメント`
- 追記される列（無ければ右端にヘッダーごと自動追加）
	- `AI_工数`, `AI_Level`, `AI_NextAction`, `AI_Advice`, `AI_Evaluation`, `AI_Caution`
- 追記の書式（セル内テキスト）
	- 1行目: `YYYY-MM-DD HH:MM:SS model:<モデル名>`
	- 2行目以降: 推論本文（`AI_Caution` は「・」で始まる箇条書き、改行区切り）
- 書式について
	- 出力はopenpyxlで元ブックを読み込み、AI列のみ追記します。既存の書式は可能な限り保持されます。
	- 新規に追加されるAI列のスタイルはブック既定です。
- JSON出力
	- `--json` で指定したパスに行ごとの結果を配列で保存します（例: `[{"No.": "1", "AI_工数": "0.5人日", ...}, ...]`）。

## オプション一覧

- `--input/-i`  入力Excel（既定: `kadai_kanri_simple.xlsx`）
- `--output/-o` 出力Excel（既定: 入力名に `_with_ai` を付与）
- `--json`      行ごとのAI結果JSONの出力パス（不要なら `--json ''`）
- `--model`     OpenAIモデル（既定: `gpt-5`）
- `--dry-run`   APIを呼ばずにダミー結果で検証
- `--write-mode` 出力方法（`new` または `overwrite`）
	- `new`（既定）: 既存ファイル名があれば `(...).xlsx` のように連番で新規作成
	- `overwrite`: 指定先に上書き（`--output`未指定時は入力ファイルに上書き）

モデル互換性について:
- 将来のモデル（例: `gpt-5` など）を指定可能です。
- まず Chat Completions API で JSON Schema による厳密JSONを要求し、失敗時は Responses API に自動フォールバック。さらに一部モデルで `temperature` 未対応の場合は自動で省略して再試行します。

## よくある質問 / トラブルシュート

- エラー: `OPENAI_API_KEY が環境変数に設定されていません`
	- PowerShellで `($env:OPENAI_API_KEY)` を確認し、未設定なら再設定してください。`.env` ファイルでも可。
- エラー: `ImportError: openai`
	- `pip install -r requirements.txt` を再実行してください。
- モデル名に関するエラー
	- 利用可能な正式モデル名を指定してください。提供状況はアカウント/リージョンに依存します。
- Excelの書式が変わる
	- 本ツールは既存ブックを読み込み、AI列のみ追記するため、基本的に書式は保持されます。個別の列配置や体裁のご希望があれば調整可能です。

## ライセンス

本プロジェクトは `LICENSE` に従います。