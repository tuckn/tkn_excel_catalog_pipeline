# Excel Catalog Pipeline

[English](README.md)

`tkn-excel-catalog`は、`.xlsx` / `.xlsm` workbookをMarkdown代理ノートへ投影し、
選択したmetadataを双方向に安全に同期するCLIです。Excelはsource workbookのまま、
Markdownを検索用catalog entry兼metadata編集面として使います。
ExcelからMarkdownへの反映は`pull`、MarkdownからExcelへの反映は`push`です。
前回両者が一致した同期stateを基準に変更方向とconflictを判定します。
変更を伴う処理はすべて、既定ではreportだけを作るdry-runであり、
workbookや代理ノートを自動削除しません。

## 必要環境

- WindowsまたはPython 3.11以上が動く環境
- [`uv`](https://docs.astral.sh/uv/)
- `.xlsx`または`.xlsm`のsource workbook

metadataの検査にはExcel、Excel COM、Obsidianは不要です。Obsidianが関係するのは、
代理ノートrename時にbacklink更新が必要な場合だけです。
旧`.xls` workbookは主CLIの対象外です。desktop版Excelを使う変換方法は
「仕様」の「旧`.xls` workbookの変換」を参照してください。

## インストール

通常は、次のcommandでinstallします。例示している`C:\path\to\tkn_excel_catalog_pipeline`は、このrepositoryの実際のfolder pathへ置き換えてください。

```console
cd "C:\path\to\tkn_excel_catalog_pipeline"
uv tool install .
tkn-excel-catalog --help
```

このinstallationでは、install時点のcode、package resource、dependencyが`uv`の
tool環境へ格納され、repository内の変更は自動反映されません。最後のcommandは、
install後にCLI entry pointが動作することを確認します。現在の設定は
`tkn-excel-catalog config show`、versionは`tkn-excel-catalog --version`で確認できます。

`git pull`などでrepositoryを更新した後は、次のcommandで再インストールします。

```console
cd "C:\path\to\tkn_excel_catalog_pipeline"
uv tool install . --reinstall
tkn-excel-catalog --help
```

実行名は`excel-catalog`から`tkn-excel-catalog`へ変更されました。旧versionをinstall済みの
場合は、上記の`--reinstall`付きcommandでinstall済みentry pointを更新できます。

`--reinstall`により、更新後のcode、package resource、dependencyをtool環境へ確実に
反映します。`--force`はtool installation自体を強制しますが、同じversionのpackageを
必ず再構築・再installするoptionではありません。

## 設定

### 設定fileを作成する

まずuser共通の設定fileを生成し、例示pathを書き換えます。

```console
tkn-excel-catalog config init
```

`~/.tkn/excel_catalog_pipeline/config.yaml`が作成されます。同じ内容のfileがある場合は
変更しないため、再実行しても安全です。編集済みのfileは上書きせずに停止します。
`tkn-excel-catalog config init --force`を指定するとtemplateで置き換えるため、既存の編集を
破棄するときだけ使用してください。

最初には`sources[].id`、`sources[].path`、`sources[].notes.root`を実際の値に
書き換えれば、基本操作を始められます。

### 主な設定項目

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
| `sources[].notes.profile` | note形式のprofileです。現versionでは`tkn-obsidian-v1`のまま使用します。 |
| `sources[].notes.frontmatter_term_format` | `keywords`と`categories`の形式です。既定の`obsidian-link`は`[[term]]`、`plain`は通常文字列で出力します。 |
| `sources[].notes.rename_adapter` | rename・folder移動の処理方法です。既定の`report-only`はfileを変更せず、`filesystem`は後述の明示的なwrite optionと組み合わせて直接変更します。 |
| `sync.max_extracted_text_chars` | 代理noteへ保存するworkbook抽出textの最大文字数です。 |

その他の`sync`の真偽値は、将来の拡張に備えた安全方針です。現versionは未知のnote
metadataを常に保持し、fileを削除しません。workbook renameの適用には、設定とは別に
`push --write-excel --allow-rename`も必要です。

Windows pathはYAMLのsingle quoteで囲む方法を推奨します。例:
`'C:\path\to\excel-workbooks'`。single quote内ではbackslashを`\\`へ二重化する必要は
ありません。

### 設定内容を確認する

設定は次の順で読み込まれます。

1. `~/.tkn/excel_catalog_pipeline/config.yaml`
2. `./.tkn/config.yaml`
3. `--config`で指定するfile

後の設定が前を上書きし、個別CLI optionが最優先です。相対pathはcurrent working
directory基準です。実行fileを作らず解決結果を確認できます。

```console
tkn-excel-catalog config show
tkn-excel-catalog --config C:\path\to\config.yaml config show
```

### サブfolderを読む設定例

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
      frontmatter_term_format: obsidian-link
      rename_adapter: filesystem
```

`recursive: true`では、source rootからの相対folderを代理note側にも再現します。例えば`2008/所有してきたCPUのベンチマーク比較.xlsx`は、notes rootの`2008/所有してきたCPUのベンチマーク比較.xlsx.md`となり、Frontmatterは`sourceFileName: 2008/所有してきたCPUのベンチマーク比較.xlsx`になります。既存workbookを年folderへ移動した後は、まず`pull`で予定を確認し、review後に`pull --write-notes`を実行してください。追跡済み代理noteのfolder移動を直接適用するには`rename_adapter: filesystem`が必要です。

rename・folder移動の方向、許可option、Obsidian backlinkの注意点は、
「仕様」の「rename・folder移動」を参照してください。

## 基本操作

変更を伴うcommandは、まずwrite optionなしのdry-runで実行します。
consoleのsummaryとreport folderをreviewし、意図した内容であることを確認してから、
同じcommandに明示的なwrite optionを付けて適用します。

### 状態を確認する

設定したsource内のExcel workbookと代理ノートを走査し、追跡状況、重複ID、
読み取りエラーを確認:

```console
tkn-excel-catalog status
```

### ExcelからMarkdownへ反映する

ExcelからMarkdownへの変更を確認し、review後に適用:

```console
tkn-excel-catalog pull
tkn-excel-catalog pull --write-notes
```

通常の`pull`は、Markdownだけで編集されたように見えるmetadataを上書きせず保持します。
過去の誤った同期stateからの復旧など、Excelをすべての差分fieldの正とする場合は、
dry-run reportを確認してから`--prefer-source`を明示します。

```console
tkn-excel-catalog pull --source example --prefer-source
tkn-excel-catalog pull --source example --prefer-source --write-notes
```

subcommandの前にglobal option `-v`を付けると、metadataの比較値をfield単位で表示します。
同じ完全な値は、workbookの1 fieldを1行とする`differences.csv`にも保存します。

```console
tkn-excel-catalog -v pull --source example --prefer-source
```

### MarkdownからExcelへ反映する

MarkdownからExcelへの変更を確認し、review後にbackup付きで適用:

```console
tkn-excel-catalog push
tkn-excel-catalog push --write-excel
```

各statusの意味と終了codeは、「仕様」の「console出力、status、終了code」を
参照してください。

### 対象を限定する

特定ノートだけを対象にする例:

```console
tkn-excel-catalog push --note example.xlsx.md
tkn-excel-catalog push --note example.xlsx.md --write-excel
```

### workbook IDを付与する

stable workbook IDの不足を確認し、明示的にcustom propertyへ付与:

```console
tkn-excel-catalog adopt
tkn-excel-catalog adopt --write-excel
```

## 仕様

### 同期modelと管理state

このCLIはExcelとMarkdownの二者だけを比較するのではなく、次の3つを
workbookごと、metadata fieldごとに比較します。

```text
現在のExcel workbook metadata -------\
前回一致した値 (sync-state.json) ---+-> CLIの三方向比較 -> pull / push / conflict
現在のMarkdown代理ノート ---------/
```

| 構成要素 | 役割 |
| --- | --- |
| Excel workbook | 実データである`.xlsx` / `.xlsm`と、現在のOOXML metadataを保持します。 |
| Markdown代理ノート | workbookの検索用catalog entryであり、`push`でExcelへ戻すmetadataの編集面です。 |
| `~/.tkn/excel_catalog_pipeline/state/sync-state.json` | CLIが管理する同期stateです。workbook ID、source / note path、前回両者が一致した`baseMetadata`などを保持します。userがmetadataを編集するfileではありません。 |

例えば、前回一致したtitleだけをExcel側で変更した場合は`pull`、
Markdown側だけで変更した場合は`push`の対象です。両方が前回値から
別々の値へ変わった場合は`conflict`とし、勝手に上書きしません。
明示的なwrite後は、ExcelとMarkdownが実際に一致したfieldだけを新しい
baseとして`sync-state.json`へ記録します。dry-runはこのstateを更新しません。
runごとの`summary.json`や`differences.csv`は判定のreview用reportであり、
上記の継続的な同期stateとは別です。

### metadata契約

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
既定ではquote付きObsidian linkとして保持します。通常文字列にする場合は
`sources[].notes.frontmatter_term_format: plain`を設定します。どちらの形式もExcelへは
link記法を含まないtermとして`; `で結合します。
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

### conflictの判定と解決

前回一致時のbase、現在のworkbook、現在のnoteをfield単位で三方向比較します。双方が
異なる値へ変わった場合は終了code `2`で停止し、mtimeによるlast-write-winsは行いません。
review後に`--prefer-source`または`--prefer-note`を明示できます。
`pull`の`--prefer-source`は、note側だけが変更されたと判定されたmetadataも現在のworkbook値で
置き換えます。このoptionがない通常の`pull`では、Markdown編集の黙示的な破棄を防ぐため、
そのnote値を保持します。
部分的な`pull`または`push`の後は、workbookと代理noteが実際に一致したfieldだけbaseを
更新します。反対方向に残る未同期変更は同期済みにせず、次のcommandへ引き継ぎます。

### rename・folder移動

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

rename先はsource root内に収まり、現在と同じ拡張子で、Windows上有効かつ衝突しない相対pathである必要があります。
`filesystem`はMarkdown fileを直接移動し、Obsidian backlinkを更新しません。そのため既定値は`report-only`です。backlink更新が必要なVaultではreportを確認し、Obsidian上で手動renameするか、対応する外部adapterが追加されるまで`report-only`を使用してください。

### run reportと同期state

run report:

```text
~/.tkn/excel_catalog_pipeline/state/runs/<run-id>/
  summary.json
  actions.csv
  details.json
  differences.csv
```

base stateは`~/.tkn/excel_catalog_pipeline/state/sync-state.json`、backupは隣接する
`backups/`です。dry-runはreportを作成しますが、workbook、note、state、cacheを変更しません。
file別reportでは、`sourceToNoteFields`が`pull`でworkbookから代理noteへ反映するfield、
`noteToSourceFields`が`push`で代理noteからworkbookへ反映するfieldを示します。
`changedFields`は、そのreportを作成したcommandが反映または反映予定としたfieldです。
`differences.csv`には`baseValue`、`excelValue`、`noteValue`、元の三方向判定
`direction`、明示操作による`plannedDirection`を記録します。JSONを読まなくても、
空欄による削除やExcelを正とした復旧内容を確認できます。

### workbook writeの安全性

書き込みは次の契約です。

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

### 旧`.xls` workbookの変換

Windows用の補佐scriptはdesktop版Excelを使い、旧`.xls` workbookをcatalogへ取り込む前に
変換します。元fileは残し、既存の変換先を上書きしません。VBAを含むworkbookは`.xlsm`、
それ以外は`.xlsx`として保存します。desktop版Excelが必要なのは、この変換scriptだけです。

まずdry-runで予定を確認し、次に明示的なwrite switchを付けて実行します。

```console
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath "C:\path\to\legacy-workbooks"
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath "C:\path\to\legacy-workbooks" -Write
```

`-SourcePath`は、変換元の`.xls`があるfolderを指定する引数であり、出力先の指定では
ありません。
`-OutputDirectory`を省略すると、変換後のworkbookは各`.xls`と同じfolderへ作成されます。
別folderへまとめて出力する場合は`-OutputDirectory`、サブfolderも対象にする場合は
`-Recurse`を指定します。明示した出力先は単一folderとなるため、同名の変換先は
collisionとして報告し、書き込みません。進捗はstderr、最後のcompact JSON resultは
集計だけを含む短い1件としてstdoutへ出力します。拡張子が厳密に`.xls`のfileだけを
検査し、既存の`.xlsx`と`.xlsm`は対象外とします。

### console出力、status、終了code

人向け進捗はstderrへ`[LEVEL] message`として出します。pull、push、adoptの最後には、
status件数とreport pathをindent付きsummaryで表示します。`pull`では、`unchanged`以外の
代理noteごとに、絶対pathの`notePath`と相対`sourcePath`を表示します。`push`では、処理対象のsource ID、
workbook path、include pattern、notes設定を最初に表示します。その後、`unchanged`以外の
fileごとの結果を確定するたびに、noteのfile nameと相対`sourcePath`を表示し、最後に
indent付きのsummaryを表示します。`unchanged`は個別表示せず、最終summaryに総数だけを表示します。
同期commandはstdoutの末尾にraw JSONを追加しません。最終summaryには`summary.json`の絶対pathを
表示し、`details.json`とCSVのreview artifactも、表示されたreport folderへ従来どおり保存します。`config show`は解決済み設定を
JSONとして引き続き表示します。`-v` / `--verbose`では、fieldごとのExcel値、
Markdown値、base値、適用方向も表示します。`--quiet`、`--verbose`、`--no-color`を提供し、
`NO_COLOR`も尊重します。

`status`は、設定したsourceごとに件数をまとめ、未追跡workbook、未追跡代理ノート、
重複ID、未対応file、読み取りエラーなど、確認が必要な項目だけを一覧表示します。
追跡済みworkbookは個別表示せず、件数だけを表示します。確認項目の相対pathは分類ごとに
最大20件を表示し、全件の詳細はrun reportに保存します。

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

`sourcePath`は、開始時に表示される選択sourceの`path`を基準にした相対pathです。
root pathを各行で繰り返さず、reportをsourceの移動に対して扱いやすくするためです。
`missing-source`は対応するworkbook自体がないため`sourcePath`を持たず、ログには代理ノートの
file nameだけを表示します。

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
