# tkn-excel-note — Excel を更新可能な代理 Markdown にする

Excel ブックを、生成 AI が参照しやすい Markdown ノートにします。
1 ブックにつき 1 ノートを作り、Excel の更新を取り込み、ノートの Frontmatter で編集したメタデータを Excel に反映できます。
複数ブックを扱う場合は、同じ処理をフォルダ単位で実行できます。

`pull --context` はシートの画像・文字・配置から説明を生成し、ブック全体の概要とシート間の関係も同じノートにまとめます。
AI を使わない `pull` は、メタデータ、シート一覧、抽出テキストを更新します。
Excel が原本で、Markdown は参照・検索・編集可能なメタデータのための代理ノートです。

## インストールする

Python 3.11 以降と uv を用意し、リポジトリのフォルダで実行します。

```shell
cd "C:\path\to\tkn_excel_note"
uv tool install .
tkn-excel-note --help
```

入力は `.xlsx` / `.xlsm` です。通常の `pull` / `push` に Excel の起動や AI 接続は不要です。
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
- ブック全体の概要、シート間の関係、不明点。
- 各シートの説明と、解析に使用した画像への相対リンク。

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
ファイル引数と `--source` をともに省略すると、設定済みの全 source が対象です。
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

`pull --sheet "Sheet1" --context` で解析範囲を選べます。省略時は表示中の全シートです。
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
既存のシート説明は再利用でき、初回の `pull --context` でブック全体の説明を追加します。
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

内部 Python パッケージ名は既存の `excel_catalog_pipeline` を維持しています。
[CLI](src/excel_catalog_pipeline/cli.py)、[単体入力の解決](src/excel_catalog_pipeline/targets.py)、[同期処理](src/excel_catalog_pipeline/pipeline.py)、[AI 統合](src/excel_catalog_pipeline/ai_pull.py)、[生成処理](src/excel_catalog_pipeline/context.py)が実装の入口です。
[ノートテンプレート](src/excel_catalog_pipeline/note_profiles/tkn-obsidian-v1/template.md)と生成プロンプトはパッケージに同梱します。
コード編集をインストール環境へ直接反映する開発用途では `uv tool install -e . --reinstall` も使えます。
依存・メタデータ・リソース変更後は再インストールします。テストと設定例には架空データだけを使います。
