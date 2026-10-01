# tkn-excel-note — Excel を更新可能な代理 Markdown にする

Excel ブックを、生成 AI が参照しやすい Markdown ノートにします。
1 ブックにつき 1 ノートを作り、Excel の更新を取り込み、ノートの Frontmatter で編集したメタデータを Excel に反映できます。
複数ブックを扱う場合は、同じ処理をフォルダ単位で実行できます。

`pull --context` はシートの画像・文字・配置から説明を生成し、取得済みシートのcontextを統合したブック要約を同じノートにまとめます。
AI を使わない `pull` は、メタデータ、シート一覧、抽出テキストを更新します。
Excel が原本で、Markdown は参照・検索・編集可能なメタデータのための代理ノートです。

## インストールする

Python 3.11 以降と uv を用意し、リポジトリのフォルダで実行します。

```shell
cd "C:\path\to\tkn_excel_note"
uv tool install .
tkn-excel-note --help
```

入力は `.xlsx` / `.xlsm` です。メタデータ同期と埋め込みサムネイルの抽出には、Excel の起動や AI 接続は不要です。
シートから cover を生成・更新する場合は Windows とデスクトップ版 Microsoft Excel が必要です。
`pull --context` には Windows、デスクトップ版 Microsoft Excel、GenAI Bridge の画像対応接続先が必要です。
[AI 接続の準備](docs/guides/sheet-content.md#準備と実行)を参照してください。

## 1つの Excel を Markdown にする

Excel を保存してから実行します。フォルダの `source` 登録は不要です。

```shell
tkn-excel-note pull "C:\path\to\book.xlsx" --output "C:\path\to\book.xlsx.md" --context
```

次の内容を、Frontmatter 付きの 1 ノートに保存します。

- 元 Excel のパス、タイトル、作成者などのメタデータ。
- ブックのシート一覧と抽出テキスト。
- 取得済みcontext全体に基づくブック要約と、未取得・古い・未検証のシートの表示。
- ブック内の順番に並ぶ各シートの要約、必要な結論・要点、内容と出典画像。
- 末尾の Workbook Map。

`--sheet` は今回更新するシートを選びます。以前に取得した別シートのcontextもブック要約の材料として保持します。シート本文は「シート要約」を共通の入口とし、結論・要点は必要な場合だけ表示します。「内容」の中の見出しはシートに応じて生成します。

```shell
tkn-excel-note pull "C:\path\to\book.xlsx" --context --sheet "Sheet1" --profile default-ja
```

`--profile` は文章の構成と言語を選びます。既定は `default-ja`（日本語）です。英語は `default-en`、独自の構成や言語はユーザー定義profileを指定します。生成言語は各profileのテンプレートで管理します。[本文構成とprofileの作り方](docs/guides/sheet-content.md#contextのprofile)を参照してください。

2.0.0では `generation.language` と `prompt_profile: auto` を廃止しました。旧設定が残っている場合は、[言語設定の移行手順](docs/guides/sheet-content.md#旧言語設定からの移行)に沿って修正してください。

画像と抽出根拠はノートと同じ階層の `img/` に保存します。
生成した説明には解釈が含まれるため、重要な判断に使う箇所は元シートと見比べてください。
保存されていない Excel 上の編集は読み取りません。

`--output` を省略すると、初回は元ブックの隣に `book.xlsx.md` を作ります。設定済みの source に属するブックは、その source のノート保存先を使います。
以後は記録されたノートを再利用します。関係のない既存 Markdown は上書きしません。

AI を使わず、メタデータと抽出テキストだけで始めることもできます。

```shell
tkn-excel-note pull "C:\path\to\book.xlsx"
```

事前に対象・設定・保護条件を確認するには `--dry-run` を付けます。
ノート、画像、同期記録、バックアップ、レポートを保存せず、Excel の画像化や AI 呼び出しも行いません。

```shell
tkn-excel-note pull "C:\path\to\book.xlsx" --context --dry-run
```


## シート画像を cover にする

埋め込みサムネイルがない、または表示範囲が狭い場合は、保存済みシートの指定範囲から cover を作れます。AI 接続は不要です。

```shell
tkn-excel-note pull "C:\path\to\book.xlsx" --cover sheet
tkn-excel-note pull "C:\path\to\book.xlsx" --cover sheet --cover-sheet "概要" --cover-range "A1:Q50" --cover-width 2400
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

## Excel の更新をノートへ取り込む

Excel を保存した後、同じコマンドを実行します。

```shell
tkn-excel-note pull "C:\path\to\book.xlsx" --context
```

内容・生成条件が変わっていないシートの説明と画像は再利用します。
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
tkn-excel-note push "C:\path\to\book.xlsx.md" --dry-run
tkn-excel-note push "C:\path\to\book.xlsx.md"
```

Excel をバックアップしてから文書プロパティを更新します。
`description`、AI の説明、手書き本文は Excel へ書き戻しません。セルや図形の編集も対象外です。
`push` でノートの本文が短縮・再生成されることはありません。

Excel とノートで同じ項目を別々の値に変えた場合は、競合として報告します。
前回一致した値を使って変更方向を判定するため、同期記録も継続して保管します。
[項目の対応と競合の解決](docs/reference/synchronization.md)を参照してください。

## 複数の Excel をフォルダ単位で管理する

単体で使っている処理を、一括で実行するための設定です。

### セットアップ

```shell
tkn-excel-note config init
tkn-excel-note config show
```

`config init` が表示した設定ファイルを編集します。既存の設定は `--force` を付けない限り置き換えません。
保存先は `~/.tkn/excel_note/config.yaml` です。旧版からの移行方法は[旧版から更新する](#旧版から更新する)を参照してください。

```yaml
schema_version: "1.1.0"
sources:
  workbooks:
    workbooks_dir: 'C:\path\to\workbooks'
    recursive: true
    include: ['*.xlsx', '*.xlsm']
    notes:
      dir: 'C:\path\to\notes'
      frontmatter_term_format: plain
```

単体で作ったノートも、元ブックを `workbooks_dir`、既存ノートを `notes.dir` の探索範囲に含めると、同じノート ID と同期記録を使って管理できます。
カスタムのノート名も維持します。ノートが設定した探索範囲の外にある場合は、二重作成せず競合を報告します。
ノートと画像を物理的に移動する操作は、この設定追加には含まれません。

### 一括で更新・反映する

```shell
tkn-excel-note pull --source workbooks --context --dry-run
tkn-excel-note pull --source workbooks --context
tkn-excel-note push --source workbooks --dry-run
tkn-excel-note push --source workbooks
tkn-excel-note status --source workbooks
```

AI が不要な更新では `--context` を省略します。
`pull` は対象を明示して実行します。`pull` のみ、または対象なしの `pull --context` / `pull --dry-run` はヘルプ表示のみで、設定の読み込み・探索・保存・AI 呼び出しを行いません。
全 source を処理するときだけ `--all-sources` を指定します。

```shell
tkn-excel-note pull --all-sources --dry-run
tkn-excel-note pull --all-sources
```

単体ファイル・`--source ID`・`--all-sources` は併用できません。`--context` を付けた場合も同じ対象指定が必要です。
`push` は従来どおり、ファイル引数と `--source` をともに省略すると全 source が対象です。
`push --note "book.xlsx.md"` で一括管理の中から対象ノートを絞れます。
コマンドは呼び出したときに一度処理し、常駐監視はしません。
通常の同期で Excel や代理ノートを削除することはありません。

## コマンドと結果

| 目的 | コマンド |
| --- | --- |
| 1 ブックまたはフォルダ内の代理ノートを作成・更新 | `pull [workbook] [--context]` |
| 1 ノートまたはフォルダ内のメタデータを Excel に反映 | `push [note]` |
| フォルダの追跡状況を確認 | `status` |
| 保存済みブックのシート一覧 | `workbook list-sheets --workbook book.xlsx` |
| 設定の作成・確認 | `config init` / `config show` |
| 固定 ID を Excel に付与 | `adopt` |
| 元ブックのない代理ノートをバックアップして削除 | `delete-notes` |

`pull --source workbooks --sheet "Sheet1" --context` で解析範囲を選べます。省略時は表示中の全シートです。
オプションの詳細は `tkn-excel-note pull --help` で確認できます。

進捗・差分・人向けの集計は標準エラー出力、同期コマンドの結果は標準出力の 1 行 JSON です。
通常実行では実行レポートも保存します。競合や失敗があっても、先に成功した変更は残ります。
[状態と終了コード](docs/reference/synchronization.md#状態と次の操作)を確認してください。

## 旧版から更新する

```shell
cd "C:\path\to\tkn_excel_note"
uv tool install . --reinstall
tkn-excel-note --version
```

製品・配布名を `tkn-excel-note`、リポジトリ名を `tkn_excel_note` に変更しました。
旧 `export` と `context build` の機能は `pull --context` に統合しました。旧コマンドの別名は提供しません。
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
[CLI](src/excel_catalog_pipeline/cli.py)、[単体入力の解決](src/excel_catalog_pipeline/targets.py)、[同期処理](src/excel_catalog_pipeline/pipeline.py)、[AI 統合](src/excel_catalog_pipeline/ai_pull.py)、[生成処理](src/excel_catalog_pipeline/context.py)が実装の入口です。
[ノートテンプレート](src/excel_catalog_pipeline/note_profiles/tkn-obsidian-v1/template.md)と[context profile](src/excel_catalog_pipeline/context_profiles/)のプロンプト・出力スキーマ・テンプレートをパッケージに同梱します。
コード編集をインストール環境へ直接反映する開発用途では `uv tool install -e . --reinstall` も使えます。
依存・メタデータ・リソース変更後は再インストールします。テストと設定例には架空データだけを使います。
