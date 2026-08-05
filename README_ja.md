# Excel Catalog Pipeline

[English](README.md)

`excel-catalog`は、`.xlsx` / `.xlsm` workbookをMarkdown代理ノートへ投影し、
選択したmetadataを双方向に安全に同期するCLIです。Excelはsource workbookのまま、
Markdownを検索用catalog entry兼metadata編集面として使います。

変更を伴う処理はすべて、既定ではreportだけを作るdry-runです。workbookや代理ノートを
自動削除しません。

## 必要環境

- WindowsまたはPython 3.11以上が動く環境
- [`uv`](https://docs.astral.sh/uv/)
- `.xlsx`または`.xlsm`のsource workbook

metadataの検査にはExcel、Excel COM、Obsidianは不要です。Obsidianが関係するのは、
代理ノートrename時にbacklink更新が必要な場合だけです。

## インストール

通常は、次のcommandでeditable installationを行います。例示している`C:\path\to\tkn_excel_catalog_pipeline`は、このrepositoryの実際のfolder pathへ置き換えてください。

```console
uv tool install -e "C:\path\to\tkn_excel_catalog_pipeline"
excel-catalog config show
```

`-e`（`--editable`）を指定すると、installされた`excel-catalog`はrepository内の
source codeを直接参照します。そのため、通常のsource code変更や`git pull`の内容は、
再installせずに反映されます。2つ目のcommandは、install後に現在の設定を表示し、
CLI entry pointと設定解決が動作することを確認します。必要に応じて
`excel-catalog --version`と`excel-catalog --help`も確認できます。

repository folderを移動・renameした場合、dependencyを変更した場合、または
`pyproject.toml`のpackage metadataやCLI entry pointを変更した場合は、editable
installationを作り直します。

```console
uv tool install -e "C:\path\to\tkn_excel_catalog_pipeline" --force
```

repository内の変更を自動反映しないnon-editable installationへ切り替える場合は、
次のcommandを実行します。

```console
uv tool install "C:\path\to\tkn_excel_catalog_pipeline" --force
```

non-editable installationではinstall時点のcodeが使われます。`git pull`などで
repositoryを更新した後は、同じcommandを再実行して変更を反映してください。

## 設定

まずuser共通の設定fileを生成し、例示pathを書き換えます。

```console
excel-catalog config init
```

`~/.tkn/excel_catalog_pipeline/config.yaml`が作成されます。同じ内容のfileがある場合は
変更しないため、再実行しても安全です。編集済みのfileは上書きせずに停止します。
`excel-catalog config init --force`を指定するとtemplateで置き換えるため、既存の編集を
破棄するときだけ使用してください。

主な設定項目は次のとおりです。

| 設定 | 意味 |
| --- | --- |
| `schema_version` | 設定formatのversionです。`1`のまま使用します。 |
| `sources[].id` | Excel sourceを識別する、一意で安定した任意の名前です。 |
| `sources[].path` | source Excel fileを格納するroot folderです。 |
| `sources[].include` | root配下で対象にするworkbookのglob patternです。 |
| `sources[].notes.root` | 代理noteのfolderです。`pull`の出力先であり、`push`の入力元です。 |
| `sources[].profile` | note形式のprofileです。現versionでは`tkn-obsidian-v1`のまま使用します。 |
| `sources[].rename_adapter` | `report-only`はrenameを報告だけし、`filesystem`は明示的に許可したrenameを適用できます。 |
| `sync.max_extracted_text_chars` | 代理noteへ保存するworkbook抽出textの最大文字数です。 |

その他の`sync`の真偽値は、将来の拡張に備えた安全方針です。現versionは未知のnote
metadataを常に保持し、fileを削除しません。workbook renameの適用には、設定とは別に
`push --write-excel --allow-rename`も必要です。

Windows pathはYAMLのsingle quoteで囲む方法を推奨します。例:
`'C:\path\to\excel-workbooks'`。single quote内ではbackslashを`\\`へ二重化する必要は
ありません。

設定は次の順で読み込まれます。

1. `~/.tkn/excel_catalog_pipeline/config.yaml`
2. `./.tkn/config.yaml`
3. `--config`で指定するfile

後の設定が前を上書きし、個別CLI optionが最優先です。相対pathはcurrent working
directory基準です。実行fileを作らず解決結果を確認できます。

```console
excel-catalog config show
excel-catalog --config C:\path\to\config.yaml config show
```

source IDは一意である必要があります。schemaは複数rootを表現でき、`--source <id>`で
1つに限定できます。

## 基本操作

inventory確認:

```console
excel-catalog status
```

ExcelからMarkdownへの変更を確認し、review後に適用:

```console
excel-catalog pull
excel-catalog pull --write-notes
```

MarkdownからExcelへの変更を確認し、review後にbackup付きで適用:

```console
excel-catalog push
excel-catalog push --write-excel
```

`push`のfile別status:

| status | 意味 |
| ------ | ---- |
| `unchanged` | workbookと代理ノートに同期対象の差分がありません。個別ログには表示せず、summaryに総数だけを表示します。 |
| `would-write` | 代理ノート側の変更をworkbookへ書き込む予定です。dry-runのため、まだ書き込んでいません。 |
| `written` | workbookへの書き込みと検証が完了しました。 |
| `missing-source` | 代理ノートに一意に対応するworkbookが見つかりません。workbookは変更しません。 |
| `pull-required` | workbook側に取り込むべき変更があります。`push`では変更せず、`pull`で確認します。 |
| `conflict` | base、workbook、代理ノートの比較で競合しました。内容を確認し、必要な場合だけ`--prefer-note`または`--prefer-source`を指定します。 |
| `duplicate-id` | 同じ`TknExcelCatalogId`を持つworkbookが複数あり、一意に対応付けできません。 |
| `rename-required` | rename要求がありますが、明示許可または設定されたadapterでの処理が必要です。 |
| `rename-error` | rename先の名前、衝突、または関連fileの処理で問題が発生しました。 |
| `read-error` | workbookまたは代理ノートの検出・読み取りに失敗しました。 |
| `write-error` | workbookへの書き込みまたは書き込み後の検証に失敗しました。可能な範囲でrollbackします。 |

`sourcePath`は、開始時に表示される選択sourceの`path`を基準にした相対パスです。
root pathを各行で繰り返さず、reportをsourceの移動に対して扱いやすくするためです。
`missing-source`は対応するworkbook自体がないため`sourcePath`を持たず、ログには代理ノートの
file nameだけを表示します。

特定ノートだけを対象にする例:

```console
excel-catalog push --note example.xlsx.md
excel-catalog push --note example.xlsx.md --write-excel
```

stable workbook IDの不足を確認し、明示的にcustom propertyへ付与:

```console
excel-catalog adopt
excel-catalog adopt --write-excel
```

## metadata契約

| Markdown代理ノート | Excel core property            |
| ------------------ | ------------------------------ |
| `schemaVersion: "1.0"` | 現行の代理note Frontmatter契約。workbook metadataではない |
| `title`          | Title                          |
| `description`    | Comments / description         |
| `nouns[0]`       | Categories / category          |
| `nouns[1..]`     | Tags / keywords                |
| `sourceFileName` | metadataではなく明示rename要求 |

代理ノート名は`<workbook-name>.xlsx.md`または`<workbook-name>.xlsm.md`です。生成管理する
本文sectionは`excel-catalog` markerで囲みます。未知のFrontmatter fieldとmarker外の
手書き本文は保持します。
`schemaVersion`がないlegacy代理noteも引き続き読めます。review後の明示的なnote書込みで
現profileのversionを追加し、dry-runではnoteを変更しません。

生成Frontmatter契約と本文構造（`schemaVersion`、見出し、section順、default description）は、application-owned profile
`src/excel_catalog_pipeline/note_profiles/tkn-obsidian-v1/template.md`をsource of truthとします。
このfileはpackage resourceとして配布し、user設定にはしません。Pythonはresourceの読込と
validation、動的なworkbook内容の生成、安全なmarker置換を担当します。

## conflictとrename

前回一致時のbase、現在のworkbook、現在のnoteをfield単位で三方向比較します。双方が
異なる値へ変わった場合は終了code `2`で停止し、mtimeによるlast-write-winsは行いません。
review後に`--prefer-source`または`--prefer-note`を明示できます。

workbook renameには次をすべて要求します。

- Frontmatter `sourceFileName`の編集
- `push --write-excel --allow-rename`
- 同じ拡張子で、Windows上有効かつ衝突しないfilename
- 代理ノートも調整できるrename adapter

`rename_adapter: report-only`では`rename-required`で停止します。`filesystem`はbacklink
更新なしの直接Markdown renameを許容できる場合だけ使ってください。Obsidian固有renameは、
外部adapterが設定されるまではreport-onlyです。

## 出力と安全性

run report:

```text
~/.tkn/excel_catalog_pipeline/state/runs/<run-id>/
  summary.json
  actions.csv
  details.json
```

base stateは`~/.tkn/excel_catalog_pipeline/state/sync-state.json`、backupは隣接する
`backups/`です。dry-runはreportを作成しますが、workbook、note、state、cacheを変更しません。

workbook writeは次の契約です。

- `.xlsx` / `.xlsm`だけを正式対応する。
- replace前にbackupする。
- 選択したOOXML metadata entryだけを書き換える。
- VBAと無関係なZIP entryを保持する。
- ZIP integrityと書き込み後metadataを再検証する。
- digital signature付きOOXML packageはwrite拒否する。
- lock中sourceを安全にreplaceできなければ失敗にする。

暗号化workbook、`.xls`、`.xlsb`、削除同期、format変換、図形や二次元layoutの意味解析は
MVP対象外です。cell text抽出は検索補助であり、内容の完全表現ではありません。

## consoleと終了code

人向け進捗はstderrへ`[LEVEL] message`として出します。`push`では、処理対象のsource ID、
workbook path、include pattern、notes設定を最初に表示します。その後、`unchanged`以外の
fileごとの結果を確定するたびに、noteのfile nameと相対`sourcePath`を表示し、最後に
indent付きのsummaryを表示します。`unchanged`は個別表示せず、最終summaryに総数だけを表示します。stdoutには
従来どおりcompact JSONを1件だけ出します。`--quiet`、`--verbose`、`--no-color`を提供し、
`NO_COLOR`も尊重します。

- `0`: 成功
- `1`: 実行・validation・partial write error
- `2`: 未解決conflict
- `3`: 設定または対象選択error

## 開発と検証

```console
uv sync --locked
uv run pytest
uv run ruff check .
uv run mypy src
uv build
```

testはsynthetic workbookだけを使用します。private path、実workbook metadata、credential、
Vault内容をfixtureや文書へ追加しないでください。
