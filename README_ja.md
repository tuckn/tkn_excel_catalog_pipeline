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
| `sources[].recursive` | `true`ならサブフォルダも探索します。未設定時の既定値は`false`で、root直下だけを読みます。 |
| `sources[].include` | 探索する階層で対象にするworkbookのglob patternです。 |
| `sources[].ignore` | source root基準で対象外にするfileまたはfolderのglob patternです。 |
| `sources[].notes.root` | 代理noteのfolderです。`pull`の出力先であり、`push`の入力元です。 |
| `sources[].profile` | note形式のprofileです。現versionでは`tkn-obsidian-v1`のまま使用します。 |
| `sources[].rename_adapter` | rename・folder移動の処理方法です。既定の`report-only`はfileを変更せず、`filesystem`は後述の明示的なwrite optionと組み合わせて直接変更します。 |
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

サブフォルダを読み、不要なfolderを除外する設定例です。patternの区切りにはOSに関係なく`/`を使います。`ignore`は`path`からの相対pathに適用され、folder全体は`archive/**`のように指定できます。

```yaml
sources:
  - id: personal-excel
    path: 'C:\path\to\excel-workbooks'
    recursive: true
    include:
      - "*.xlsx"
      - "*.xlsm"
    ignore:
      - "archive/**"
      - "**/Temp/**"
    notes:
      root: 'C:\path\to\obsidian-vault\reference\entities\files\Excel'
      profile: tkn-obsidian-v1
      rename_adapter: filesystem
```

`recursive: true`では、source rootからの相対folderを代理note側にも再現します。例えば`2008/所有してきたCPUのベンチマーク比較.xlsx`は、notes rootの`2008/所有してきたCPUのベンチマーク比較.xlsx.md`となり、Frontmatterは`sourceFileName: 2008/所有してきたCPUのベンチマーク比較.xlsx`になります。既存workbookを年folderへ移動した後は、まず`pull`で予定を確認し、review後に`pull --write-notes`を実行してください。追跡済み代理noteのfolder移動を直接適用するには`rename_adapter: filesystem`が必要です。

### rename・folder移動の設定

rename・folder移動は、変更した場所によって次の2方向があります。

| 変更した場所 | CLIが検出する要求 | 許可時にCLIが変更する対象 |
| --- | --- | --- |
| Explorerなどでworkbookをrename・移動した後に`pull` | sourceの相対pathに合わせた代理noteのrename・移動 | 代理note |
| 代理noteのFrontmatter `sourceFileName`を編集した後に`push` | source root内でのworkbookのrename・移動 | workbookと代理note |

`rename_adapter`は、この要求をfilesystem上で直接処理してよいかをsource単位で指定します。

- `report-only`: 既定値です。rename・folder移動が必要でもfileを変更せず、actionを`rename-required`としてrun reportへ記録します。
- `filesystem`: backlink更新を伴わない通常のfilesystem操作で、workbookまたは代理noteを直接rename・移動できるようにします。

`filesystem`を設定しただけではfileは変更されません。実際の変更には、方向に応じて次のwrite optionも必要です。

| 方向 | 必要な設定とcommand |
| --- | --- |
| workbookを先にrename・移動し、代理noteを追従させる | `rename_adapter: filesystem`と`pull --write-notes` |
| `sourceFileName`を編集し、workbookと代理noteをrename・移動する | `rename_adapter: filesystem`と`push --write-excel --allow-rename` |

optionなしの`pull`と`push`はdry-runです。`sync.allow_source_rename`は現versionのrename許可判定には使用されないため、上記の`rename_adapter`とcommand optionを使用してください。

`rename-required`の対象、現在のsource path・note path、理由は`~/.tkn/excel_catalog_pipeline/state/runs/<run-id>/actions.csv`と`details.json`へ記録されます。`pull`で代理noteの移動が必要な場合は、希望するnote pathと衝突情報も`details.json`に入ります。`push`側の希望する相対pathは編集したFrontmatter `sourceFileName`で確認します。command終了時にreport folderがconsoleへ表示され、`push`ではfile単位の`rename-required`もstderrへ表示されます。

`filesystem`はMarkdown fileを直接移動し、Obsidian backlinkを更新しません。そのため既定値は`report-only`です。backlink更新が必要なVaultではreportを確認し、Obsidian上で手動renameするか、対応する外部adapterが追加されるまで`report-only`を使用してください。

## 基本操作

設定したsource内のExcel workbookと代理ノートを走査し、追跡状況、重複ID、
読み取りエラーを確認:

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
| `schemaVersion: "2.0"` | 現行の代理note Frontmatter契約。workbook metadataではない |
| `title`          | Title                          |
| `subject`        | Subject                        |
| `author`         | Author / creatorを単一文字列で保持 |
| `keywords[]`     | Keywords / tags                |
| `categories[]`   | Category / categories          |
| `comments`       | Comments / description         |
| `sourceCreated`  | Created。source-ownedかつpull専用 |
| `sourceModified` | Modified。source-ownedかつpull専用 |
| `sourceFileName` | source root基準の相対workbook path。編集時は明示rename/move要求 |

`description`は代理ノート専用の任意説明で、Excelへpushしません。Frontmatterのlist値は
quote付きObsidian linkとして保持し、Excelへはlink記法を外して`; `で結合します。
commaはtermの一部として保持します。`sourceCreated`と`sourceModified`は`pull`でのみ
更新し、`push`では現在のnote値を無視・保持します。
生成する`type`、source日時、`date`、`updated`、`noteId`の値はquoteなしのplain YAML
scalarとし、`schemaVersion`はquote付き文字列のまま保持します。

代理ノート名は`<workbook-name>.xlsx.md`または`<workbook-name>.xlsm.md`です。再帰探索時はsourceの相対folder構造をnotes root配下に再現します。生成管理する
本文sectionは`excel-catalog` markerで囲みます。未知のFrontmatter fieldとmarker外の
手書き本文は保持します。workbook core metadataとstable custom IDはFrontmatterに置くため、
schema 2.0では`## Excel Metadata`を生成せず、明示的なnote書込み時に旧managed sectionを
削除します。
`schemaVersion`がない、またはschema 1.0の`noun` / `nouns`を使うlegacy代理noteも
引き続き読めます。review後の明示的なnote書込みで現profileのversionを追加し、dry-runでは
noteを変更しません。

生成Frontmatter契約と本文構造（`schemaVersion`、見出し、section順、default description）は、application-owned profile
`src/excel_catalog_pipeline/note_profiles/tkn-obsidian-v1/template.md`をsource of truthとします。
このfileはpackage resourceとして配布し、user設定にはしません。Pythonはresourceの読込と
validation、動的なworkbook内容の生成、安全なmarker置換を担当します。

## conflictとrename

前回一致時のbase、現在のworkbook、現在のnoteをfield単位で三方向比較します。双方が
異なる値へ変わった場合は終了code `2`で停止し、mtimeによるlast-write-winsは行いません。
review後に`--prefer-source`または`--prefer-note`を明示できます。

rename・folder移動の検出方向、必要なoption、report先は「設定」の「rename・folder移動の設定」を参照してください。rename先はsource root内に収まり、現在と同じ拡張子で、Windows上有効かつ衝突しない相対pathである必要があります。

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
- workbookの上書き前にbackupする。
- application用のOS一時directoryで置換用OOXMLを構築する。
- 選択したOOXML metadata entryだけを書き換える。
- VBAと無関係なZIP entryを保持する。
- 上書き前後にZIP integrityとmetadataを再検証する。
- file自体を置換せず既存fileの内容だけを上書きし、OSを問わずfile identityと作成日時を保持する。
- 上書きまたは書き込み後検証に失敗した場合は、backupの内容と日時を復元する。
- digital signature付きOOXML packageはwrite拒否する。
- lock中sourceを安全に上書きできなければ失敗にする。

暗号化workbook、`.xls`、`.xlsb`、削除同期、format変換、図形や二次元layoutの意味解析は
MVP対象外です。cell text抽出は検索補助であり、内容の完全表現ではありません。

## consoleと終了code

人向け進捗はstderrへ`[LEVEL] message`として出します。`push`では、処理対象のsource ID、
workbook path、include pattern、notes設定を最初に表示します。その後、`unchanged`以外の
fileごとの結果を確定するたびに、noteのfile nameと相対`sourcePath`を表示し、最後に
indent付きのsummaryを表示します。`unchanged`は個別表示せず、最終summaryに総数だけを表示します。stdoutには
従来どおりcompact JSONを1件だけ出します。`--quiet`、`--verbose`、`--no-color`を提供し、
`NO_COLOR`も尊重します。

`status`は、設定したsourceごとに件数をまとめ、未追跡workbook、未追跡代理ノート、
重複ID、未対応file、読み取りエラーなど、確認が必要な項目だけを一覧表示します。
追跡済みworkbookは個別表示せず、件数だけを表示します。確認項目の相対pathは分類ごとに
最大20件を表示し、全件の詳細はrun reportに保存します。

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
