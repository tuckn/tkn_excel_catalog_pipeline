# tkn-excel-note — Excel を Markdown にする

Excel ブックを、RAG・生成 AI・検索で参照しやすい Frontmatter 付き Markdown にします。
まず `export` で1つのブック、またはフォルダ配下のブックを書き出せます。
継続的に更新を取り込み、Frontmatter の変更を Excel へ反映する場合は、`source` を設定して `pull` / `push` で同期します。
Excel が原本です。セル・図形の編集を Markdown から書き戻す機能はありません。

## インストールする

Python 3.11 以降と uv を用意し、リポジトリのフォルダで実行します。

```shell
cd "C:\path\to\tkn_excel_note"
uv tool install .
tkn-excel-note --help
```

入力は `.xlsx` / `.xlsm` です。通常の書き出しとメタデータ同期には、Excel の起動や AI 接続は不要です。
`--context` による説明生成とシート cover の描画には Windows とデスクトップ版 Microsoft Excel が必要です。
`--context` は GenAI Bridge の画像対応接続先も使用します。[AI 接続の準備](docs/guides/sheet-content.md#準備と実行)を参照してください。

## 1つの Excel を Markdown にする

Excel を保存してから実行します。`source` 登録は不要です。

```shell
tkn-excel-note export "C:\My Note\2026-10-01_note.xlsx" --context
```

同じフォルダに `2026-10-01_note.xlsx.md` を生成します。Excel の拡張子を含む元ファイル名に `.md` を付けます。

- Frontmatter: タイトル・作成者などのメタデータ、原本のパス、保存内容の識別値、生成日時。
- 本文: シート一覧・構造と、セル位置付きの保存値・数式・読み取れる図形の文字。
- `--context` 指定時: シートの画像・文字・配置を解析した説明と、それらを統合したブック要約。

AI が不要なら `--context` を省略します。出力先を変える場合は、単体ファイルに `--output` を指定します。

```shell
tkn-excel-note export "C:\path\to\book.xlsx"
tkn-excel-note export "C:\path\to\book.xlsx" --output "C:\path\to\context.md" --context
```

### フォルダ配下をまとめて書き出す

```shell
tkn-excel-note export "C:\My Note" --context
```

サブフォルダを含む全 `.xlsx` / `.xlsm` が対象です。各ブックの隣に `<元ファイル名>.md` を保存します。
Excel の一時ファイル `~$...` は除外します。フォルダ指定には `--output` を使えません。
設定ファイルの `sources`、探索パターン、同期ノートの保存先は、このコマンドの対象選択に使いません。

### 確認・再生成する

```shell
tkn-excel-note export "C:\My Note" --context --dry-run
tkn-excel-note export "C:\path\to\book.xlsx" --context --force
```

`--dry-run` は入力・設定・出力先を検証するだけで、Excel の画像化・AI 呼び出し・ファイル保存は行いません。
既存 Markdown は既定で上書きしません。置き換える場合だけ `--force` を指定します。
ブック単位で全生成が成功してから公開し、失敗した場合や生成中の編集を検出した場合は既存 Markdown を保護します。
フォルダ処理では成功したブックの出力を残し、失敗したブックを結果に報告します。

`export` は毎回、新しい書き出しを作ります。同期設定・対応関係・同期記録・再利用用キャッシュは保存しません。
`type: ExcelExport` の出力は `push` の対象になりません。Frontmatter の編集を Excel に反映する場合は、以下の source 同期を使ってください。
画像と抽出根拠は Markdown の隣の `img/` に保存し、再生成前の画像は残します。AI の使用量記録は `~/.tkn/excel_note/state/export/usage/` に保存します。

既定の解析対象は表示中の全ワークシートです。`--sheet` で名前を指定すると非表示シートも選べます。
未選択シートの文字も抽出しますが、AI 説明には未解析の範囲を表示します。
`--profile` は文章の構成と言語を選びます。既定は `default-ja`、英語は `default-en`、独自の構成はユーザー定義profileを指定します。

```shell
tkn-excel-note export "C:\path\to\book.xlsx" --context --sheet "Sheet1" --profile default-ja
```

生成した説明には解釈が含まれます。重要な箇所は元シートと見比べてください。未保存の編集は読み取りません。
[本文構成・profile・抽出上限](docs/guides/sheet-content.md)も参照してください。

## 補足文を添えて説明を生成する

背景や用語の補足を、直接の文章と UTF-8 ファイルで指定できます。両方を併用でき、繰り返し指定も可能です。

```powershell
tkn-excel-note export "C:\path\to\book.xlsx" --context `
  --reference "このブックは移行方式の比較検討メモです。" `
  --reference-file "C:\path\to\background.md"
```

Excel の記載を優先し、補足文だけにある事実や結論をブックの記載として扱わないよう指示します。
補足文はシート説明とブック要約の両方に渡します。`--no-reference` は設定を含むすべての補足を無効にします。
これらのオプションには `--context` が必要です。

接続先・文章のprofile・共通補足は `generation.generators.<id>` に保存し、`--generator <id>` で選べます。
source 固有の背景は `sources.<id>.generation.reference` に置けます。
[設定例と選択順](docs/reference/configuration.md#名前付きgeneratorと補足文)を参照してください。

## 設定したフォルダを継続的に同期する

`source` は、Excel の入力フォルダと代理ノートの保存先を組にした同期対象です。
1ブックだけのフォルダにも、複数ブックのフォルダにも使えます。

### セットアップ

```shell
tkn-excel-note config init
tkn-excel-note config show
```

`config init` が表示した設定ファイルを編集します。保存先は `~/.tkn/excel_note/config.yaml` です。
既存の設定は `--force` を付けない限り置き換えません。

```yaml
schema_version: "2.1.0"
sources:
  workbooks:
    workbooks_dir: 'C:\path\to\workbooks'
    recursive: true
    include: ['*.xlsx', '*.xlsm']
    notes:
      dir: 'C:\path\to\notes'
      frontmatter_term_format: plain
```

### 一括で更新・反映する

```shell
tkn-excel-note pull --source workbooks --context --dry-run
tkn-excel-note pull --source workbooks --context
tkn-excel-note push --source workbooks --dry-run
tkn-excel-note push --source workbooks
tkn-excel-note status --source workbooks
```

AI が不要な更新では `--context` を省略します。
設定した source に対する処理は `--source ID` が必須です。省略時は引数エラーとなり、全 source を暗黙に処理しません。
ブックやノートのパスを位置引数にする `pull` / `push` は提供しません。
反映するノートを絞る場合は `push --source workbooks --note "book.xlsx.md"` とします。

同期ノートは `type: Excel` と識別子を持ち、前回一致した値を使って変更方向を判定します。
`export` の Markdown を自動で同期へ登録しません。同期ノートは別の保存先から始めてください。
各コマンドは呼び出したときに一度処理します。通常の同期で Excel や代理ノートを削除することはありません。

## Excel の更新をノートへ取り込む

Excel を保存した後、同じコマンドを実行します。

```shell
tkn-excel-note pull --source workbooks --context
```

内容・生成条件・補足文が変わっていないシートの説明と画像は再利用します。
補足ファイルはパスだけでなく内容で変更を判定します。
ブック全体の説明も、材料となるシート説明や生成条件が同じなら再利用します。
初回の AI 呼び出しは選択シート数 + 1 回、全結果を再利用できる場合は 0 回です。

`--context` なしで更新すると、AI の説明は保持します。
Excel の内容が説明生成時から変わっていれば Frontmatter の `contextStatus` を `stale` にして、再生成が必要だと示します。
`current` は保存済み Excel と生成記録が対応している状態で、人が正確さを確認したという意味ではありません。

未知の Frontmatter 項目と、管理マーカー外の手書き本文は保持します。
生成部分を手直ししている場合、`--context` は上書きを止めます。
意図して生成し直す場合だけ `--context --force` を使います。[範囲・保護・状態の詳細](docs/guides/sheet-content.md)を参照してください。

## Frontmatter の変更を Excel に反映する

ノートの `title`、`subject`、`author`、`keywords`、`categories`、`comments` を編集し、反映します。

```shell
tkn-excel-note push --source workbooks --note "book.xlsx.md" --dry-run
tkn-excel-note push --source workbooks --note "book.xlsx.md"
```

Excel をバックアップしてから文書プロパティを更新します。
`description`、AI の説明、手書き本文は Excel へ書き戻しません。セルや図形の編集も対象外です。
`push` でノートの本文が短縮・再生成されることはありません。

Excel とノートで同じ項目を別々の値に変えた場合は、競合として報告します。
前回一致した値を使って変更方向を判定するため、同期記録も継続して保管します。
[項目の対応と競合の解決](docs/reference/synchronization.md)を参照してください。

## シート画像を cover にする

埋め込みサムネイルがない、または表示範囲が狭い場合は、保存済みシートの指定範囲から cover を作れます。AI 接続は不要です。

```shell
tkn-excel-note pull --source workbooks --cover sheet
tkn-excel-note pull --source workbooks --cover sheet --cover-sheet "概要" --cover-range "A1:Q50" --cover-width 2400
```

初回の既定値は先頭の表示ワークシート、`A1:Q50`、幅 2400 px の PNG です。範囲を1枚に収め、用紙の余白を取り除きます。
`--context` とも併用できます。cover 自体は AI に送信しません。
成功した生成方式・範囲・画像幅をブックごとに記録するため、次回の `pull` はオプションを省略しても同じ条件を使います。
変更がなく画像も正常なら Excel を起動せず再利用し、ブックや条件が変わったとき、画像が欠損・破損したときに再生成します。

設定ファイルのトップレベルに `cover: {mode: sheet}` を追加すると、選択したブックに常時適用できます。
事前確認は `--cover sheet --dry-run`、埋め込み方式への切り替えは `--cover embedded` です。
手動設定した cover は保持し、画像化に失敗した場合も既存の cover を残してメタデータ同期を続けます。
[詳しい仕様と注意点](docs/guides/catalog-operations.md#シートの指定範囲から-cover-を作る)を参照してください。

## シート構造を確認する

通常の `pull` は、保存済みの Excel からシート別の事実を取得し、
既存の `Workbook Map` を表に更新します。この構造取得には Excel 起動・画像化・AI 呼び出しは不要です。
シート cover の更新が必要な場合は、その画像生成のために Excel を起動します。
`--dry-run` は読み取りだけで、ノートや同期状態を保存しません。
既存のシート説明（`context-*`）や管理領域外の本文は保持します。

| 項目 | 意味 |
| --- | --- |
| Stored range | 保存された worksheet dimension。書式だけのセルを含む場合があります。未記録なら unknown。 |
| Content range | 値または数式のあるセルを囲む最小の矩形。空白だけの文字列も値として扱います。 |
| Populated cells | 値または数式のあるセル数。0・FALSE・結果未保存の数式も対象。書式だけのセルは除外。 |
| Tables | Excel で定義されたテーブル数。見た目の表は判定しません。 |
| Shapes | 通常の図形・コネクタ数。グループの容器は数えず、内部の要素を個別に数えます。 |
| Images / Charts | シート上の配置数。同じ画像の複数配置も別々に数えます。 |
| Notes | 取得失敗・未対応項目の理由。 |

集計は非表示シートも対象で、文字抽出の上限 `sync.max_extracted_text_chars` や
AI 用の `generation.max_cells` には左右されません。広いセル範囲の面積をセル数にはしません。
既知の空は `0`、空の存在範囲は `—`、不明・未対応は `unknown` として区別します。
VML、OLE、コントロール、拡張オブジェクトなどを検出した場合は、部分的な個数を総数と誤認しないよう
図形・画像・グラフ数を unknown とします。チャートシートなど worksheet 以外も未対応として記録します。
数式の再計算や、保存後の未保存編集の取得は行いません。

同じ結果を同期状態 `~/.tkn/excel_note/state/sync-state.json` の各 entry の
`sheetInventory` に保存します。`schemaVersion: 1` と `sheets` 配列を持ち、
各シートには `name`・`sheetId`・`state`・パッケージ内の `path`、
`stored_range`・`content_range`・`populated_cells`・`tables`・`shapes`・`images`・`charts`・
`warnings` を記録します。不明値は JSON の `null`、既知の空の存在範囲は空文字列です。
VLM の対象選択に利用するための基礎情報であり、今回の追加で AI の対象選択・呼び出し条件は変わりません。

## コマンドと結果

| 目的 | コマンド |
| --- | --- |
| 単体・フォルダの独立した Markdown 書き出し | `export <file-or-folder> [--context]` |
| 設定済み source の代理ノートを作成・更新 | `pull --source ID [--context]` |
| source 内のメタデータを Excel に反映 | `push --source ID [--note ...]` |
| source の追跡状況を確認 | `status --source ID` |
| 保存済みブックのシート一覧 | `workbook list-sheets --workbook book.xlsx` |
| 設定の作成・確認 | `config init` / `config show` |
| 固定 ID を Excel に付与 | `adopt --source ID` |
| 元ブックのない代理ノートをバックアップして削除 | `delete-notes --source ID --note ...` |

進捗は標準エラー出力、処理結果は標準出力の1行 JSONです。
`export` は同期レポートを保存しません。同期コマンドは通常実行でレポートを保存します。
オプションは各コマンドの `--help`、[状態と終了コード](docs/reference/synchronization.md#状態と次の操作)を参照してください。

## 旧版から更新する

```shell
cd "C:\path\to\tkn_excel_note"
uv tool install . --reinstall
tkn-excel-note --version
```

製品・配布名を `tkn-excel-note`、リポジトリ名を `tkn_excel_note` に変更しました。
3.1.0では名前付きgeneratorと補足文を追加しました。旧 `generation.bridge_profile` / `prompt_profile` / `overrides` は引き続き使え、設定ファイルを自動変更しません。更新後は `uv tool install . --reinstall` でインストール済みCLIを更新してください。

3.0.0では、単体・フォルダの書き出しを `export` に分離しました。旧単体 `pull <workbook>` は `export <workbook>` へ変更してください。
同期は設定済み source を対象にし、`pull` / `push` / `status` / `adopt` / `delete-notes` に `--source ID` が必須です。`--all-sources` は廃止しました。旧 `context build` の別名は提供しません。
新 CLI の動作確認後、旧ツール環境があれば削除できます。

```shell
uv tool uninstall tkn-excel-catalog-pipeline
```

保存領域は `~/.tkn/excel_note/` に統一しています。旧 `~/.tkn/excel_catalog_pipeline/` を使っていた場合は、CLI を停止し、移行先が存在しないことを確認してフォルダ全体を `excel_note` へ名前変更してください。設定・同期記録・生成履歴・バックアップをまとめて引き継ぎます。両方のフォルダがある場合は自動で統合せず、内容を確認してください。
ノートの識別子、管理マーカー、Excel の固定 ID も維持するため、既存ノートを作り直す必要はありません。
既存のシート説明はブック要約の材料として保持します。選択した生成済みシートは新しいprofileで再生成します。取り込み済みの説明は保持し、記録と本文が一致する旧形式のブロックだけ見出しを整えます。旧contextのシート要約は、明示的に再生成するまで追加しません。
旧 `export` の同期情報を持たない Markdown は、独立した書き出しとして残ります。新ノートは別の保存先で作成してください。

## 詳しい情報と開発

- [AI による説明の生成](docs/guides/sheet-content.md): 接続設定、解析範囲、画像、再利用、状態。
- [設定リファレンス](docs/reference/configuration.md): 入出力先、探索範囲、設定の優先順位。
- [同期と実行結果](docs/reference/synchronization.md): Frontmatter、競合、名前変更、保存領域。
- [代理ノートの管理](docs/guides/catalog-operations.md): サムネイル、明示的な削除と復旧、旧 `.xls` の変換。
- [変更履歴](CHANGELOG.md)、[利用許諾](LICENSE)。

```shell
uv sync --locked
uv run pytest
uv run ruff check .
uv run mypy src
uv build
```

シート cover の実機回帰テストは、Windows とデスクトップ版 Excel がある環境で明示的に実行します。架空データのブックだけを作成し、非表示の専用 Excel インスタンスを使います。

```powershell
$env:TKN_EXCEL_NOTE_NATIVE_TESTS = '1'
uv run pytest tests/test_sheet_cover_native.py
Remove-Item Env:TKN_EXCEL_NOTE_NATIVE_TESTS
```

内部 Python パッケージ名は既存の `excel_catalog_pipeline` を維持しています。
[CLI](src/excel_catalog_pipeline/cli.py)、[独立した書き出し](src/excel_catalog_pipeline/export.py)、[同期処理](src/excel_catalog_pipeline/pipeline.py)、[AI 統合](src/excel_catalog_pipeline/ai_pull.py)、[生成処理](src/excel_catalog_pipeline/context.py)が実装の入口です。
[ノートテンプレート](src/excel_catalog_pipeline/note_profiles/tkn-obsidian-v1/template.md)と[context profile](src/excel_catalog_pipeline/context_profiles/)のプロンプト・出力スキーマ・テンプレートをパッケージに同梱します。
コード編集をインストール環境へ直接反映する開発用途では `uv tool install -e . --reinstall` も使えます。
依存・メタデータ・リソース変更後は再インストールします。テストと設定例には架空データだけを使います。
