# 代理ノートの管理

通常の作成・同期は [README](../../README.md) を参照してください。サムネイル、削除、旧形式の変換など、必要になった操作を説明します。

## Obsidian Bases のカードにサムネイルを表示する

初回の既定方式（`embedded`）では、`pull` は Excel に保存されたブックのサムネイルをローカルで PNG に変換し、
`notes.dir/img/excel-cover-<画像ハッシュ>.png` として保存します。
Frontmatter の `cover` には Vault 内の画像へのリンクを設定します。以下の `catalog` は、Vault 内で代理ノートを保存するフォルダ名の例です。
`.xlsx` / `.xlsm` のパッケージサムネイルが対象で、シート上の挿入画像は選びません。
Windows では EMF / WMF（Excel の標準 WMF を含む）と PNG / JPEG に対応します。
Excel の起動、AI 呼び出し、外部サービスへの画像送信は不要です。

```yaml
cover: "[[catalog/img/excel-cover-<画像ハッシュ>.png]]"
```

以下は既存 Base の `views` に追加する設定の抜粋です。Obsidian でカードビューを作成しても構いません。
画像プロパティに `cover` を指定します。既存のフィルターはそのまま使えます。

```yaml
  - type: cards
    name: Cards View
    order:
      - file.name
      - title
      - description
    image: note.cover
    imageFit: contain
    cardSize: 200
```

対象と変更予定を確認してから反映します。

```shell
tkn-excel-note pull --source workbooks --dry-run
tkn-excel-note pull --source workbooks
```

- 埋め込み方式の `--dry-run` は画像変換まで検証しますが、画像・ノート・同期状態を保存しません。
- 埋め込み方式でサムネイルがないブックは `cover` が空になります。画像のないカードも表示されます。
- 手動で設定した `cover` は保持します。自動抽出へ戻す場合は `cover` を空にしてください。
- 自動設定した画像は次回の `pull` で更新し、欠損・破損した出力画像も再生成します。
- 変換できないサムネイルは警告を出し、既存の `cover` と通常のメタデータ同期を保持します。
- 同期状態の `managedCover` で自動設定したリンクを識別します。状態を失った場合、既存の非空の `cover` は手動指定として保持します。
- 元の Excel と過去に出力した画像は削除しません。埋め込み方式の画像は最大辺 1200 px の PNG です。

カードの画像プロパティはローカル添付へのリンクを受け付けます（[Obsidian 公式ヘルプ](https://help.obsidian.md/bases/views/cards)）。

## シートの指定範囲から cover を作る

`--cover sheet` は Windows のデスクトップ版 Excel で保存済みのシートを描画し、指定範囲を1枚の PNG にします。埋め込みサムネイルの有無に左右されません。AI 呼び出しや外部への画像送信はありません。

```shell
tkn-excel-note pull "C:\path\to\book.xlsx" --cover sheet --dry-run
tkn-excel-note pull "C:\path\to\book.xlsx" --cover sheet
tkn-excel-note pull "C:\path\to\book.xlsx" --cover sheet --cover-sheet "概要" --cover-range "B2:R51" --cover-width 3000
tkn-excel-note pull --source workbooks --cover sheet
```

| オプション | 初回の既定値・意味 |
| --- | --- |
| `--cover sheet` | シート画像方式へ切り替えます。 |
| `--cover embedded` | 埋め込みサムネイル方式へ切り替えます。 |
| `--cover auto` | 既定。最後に成功した方式を継承し、記録がない場合は `embedded` を使います。 |
| `--cover-sheet` | 先頭の表示ワークシート。名前を指定する場合も表示シートが対象です。非表示シート・チャートシートは対象外です。 |
| `--cover-range` | `A1:Q50`。単一の矩形 A1 範囲。シート名付き参照・複数範囲・行列全体は指定できません。 |
| `--cover-width` | 幅 2400 px。600～4000 の整数。高さは範囲の描画結果に合わせます。 |

`--cover-sheet` / `--cover-range` / `--cover-width` はシート方式で使います。既にシート cover を生成したブックなら `--cover sheet` の再指定は不要です。
`--sheet` は AI 解析対象の指定で、cover のシート選択には使いません。`--context` と併用できます。

生成に成功すると、同期状態の `coverGeneration` に方式・シート選択・範囲・幅・描画処理の版・保存済みブックと画像のハッシュを記録します。
通常の `pull` でオプションを省略しても、埋め込み方式へ戻りません。条件とブックが同じで出力画像も正常なら、Excel を起動せず再利用します。
セル以外の保存内容変更もブックの変更として扱うため、メタデータ変更や `push` の後にも再描画する場合があります。
生成結果の画像が以前と同じでも、変更した範囲などの条件は記録します。

手動の cover は `--cover sheet` でも上書きしません。自動管理に戻す場合は Frontmatter の `cover` を空にします。
状態が失われた場合も非空の cover は保護します。失敗時は警告を出し、以前の cover と成功時の生成記録を保持して、通常のメタデータ同期を続けます。
存在しないシート、Excel 未導入、描画失敗などが対象です。古い画像を自動削除することはありません。

`--dry-run` は保存済みブックから対象シートと範囲の指定、再生成の要否を確認します。Excel の起動、一時ブック・PDF・PNG の作成、ノート・同期状態・レポートの保存は行いません。
実際の描画可否や正確な高さは実行時に確定します。結果の `details.cover` と stderr に対象・条件・`planned` / `cached` / `rendered` / `warning` を表示します。

画像化は一時コピーで行い、元の Excel ファイルと開いている作業用 Excel は変更しません。未保存の編集は含めません。
マクロ・イベント・リンク更新・再計算・起動時の外部データ更新を抑止する仕組みは、シート説明の画像化と共通です。
保存済みの印刷範囲・タイトル行列・ヘッダー・フッターを使わず、指定範囲を PDF に描画して PNG 化します。
Excel の印刷倍率と PDF 出力については [PageSetup.Zoom](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.zoom) と [Worksheet.ExportAsFixedFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.exportasfixedformat) の仕様に従います。

指定範囲内の空白セルは維持し、用紙の余白だけを取り除きます。印刷でのセル寸法を正確に取得するため、一時コピーに測定枠を加え、PDF からその枠を除去してから画像化します。
グリッド線と行列見出しは表示せず、罫線は表示します。非表示行列や図形の印刷設定などは Excel の印刷結果に従うため、画面をそのまま撮影した画像とは異なります。
結合セル・図形などが指定範囲の境界をまたぐ場合は、必要に応じて範囲を広げてください。

出力は最大1600万画素、描画範囲の幅・高さはそれぞれ12000ポイントまでです。複数ページに分かれた結果を一部だけ採用することはありません。
超過時は範囲や画像幅を小さくしてください。入力ブックのサイズ上限には `generation.max_workbook_mb`（既定100 MiB）を使います。
小さなカードでは全体を見分ける用途、画像を開いたときには細部を読む用途に向きます。

## 元ブックがない代理ノートを削除する

`delete-notes` は、元ブックの不存在を確認できた代理ノートと、そのノートに対応する同期記録を削除します。
通常実行で削除し、`--dry-run` は予定だけを表示します。ノート ID を調べる必要はありません。

ノート名を指定する場合:

```shell
tkn-excel-note delete-notes --source workbooks --note "example.xlsx.md" --dry-run
tkn-excel-note delete-notes --source workbooks --note "example.xlsx.md"
```

元ブックがないノートをまとめて削除する場合:

```shell
tkn-excel-note delete-notes --source workbooks --all-missing --dry-run
tkn-excel-note delete-notes --source workbooks --all-missing
```

`--note` はノートの絶対パス、ノートルートからの相対パス、元ブックの相対パス（`sourceFileName`）、`noteId`、`sourceId` でも指定でき、繰り返し指定できます。
同名が複数ある場合は相対パスまたは絶対パスで区別してください。各指定が一意に一致しない場合は削除しません。
`--note` と `--all-missing` はどちらか一方が必須です。`--source` を省略すると全設定 source のノートを対象にします。

個別指定でも、元ブックが存在するノートは削除しません。一括指定ではそれらを対象から外します。
「ブックと一意に対応しない」という警告だけでは削除せず、記録したパス、ブックの固定 ID、内容の照合情報で存在を確認します。
移動・リネーム先を見落とさないよう、この確認では全設定 source の入力ルート内のサブフォルダーと除外ファイルも読み取ります。
ノートの選択範囲には設定の `recursive`、`include`、`ignore` を適用します。
一括指定では、Frontmatter の `sourceRoot` が別の source を示すノートを `skipped-note` として対象外にします。
入力ルートの不在、読み取り失敗、辿れないリンク先、同期記録の欠落・曖昧さ、重複ノートなどがあれば、削除開始前にその実行全体を停止します。
現在のノートと同期記録のノート先が異なる場合、旧パスにも別実体のノートが残っていれば重複として停止します。

削除前に、対象ノートの全文と同期記録を `~/.tkn/excel_note/state/backups/deleted-notes/<run-id>/` に保存します。
手書き本文もバックアップに含みます。Excel ブック、添付画像、シート説明の処理記録は削除しません。
ノートの削除に成功した後で同期記録の該当項目だけを解除するため、次回の `pull` に削除済みノートの `missing-source` が残りません。
元ブックが後日復元された場合、次回の `pull` は新しい代理ノートを作成します。

削除中の書き込み失敗や通常の中断では、削除済みノートの復元を試みます。同期コマンドを同時に実行せず、結果を確認してください。
プロセスの強制終了などで復元できなかった場合、バックアップの `manifest.json` に元のノートパス、バックアップ名、同期記録のキーがあります。
これに従ってノートを戻し、必要ならバックアップの `sync-state.json` から対応する項目を戻します。他の同期を行った後に記録全体を上書きすると、その後の更新を失うため避けてください。
通常実行の結果にはバックアップと実行レポートの保存先を表示し、標準出力には JSON を1件返します。`--dry-run` はバックアップ・レポートも作成しません。

## 旧 `.xls` を変換する

Windows 用の [Convert-XlsToOpenXml.ps1](../../scripts/Convert-XlsToOpenXml.ps1) は、デスクトップ版 Excel を使って `.xls` を変換します。
元ファイルを残し、VBA を含む場合は `.xlsm`、それ以外は `.xlsx` にします。
既存の変換先は上書きしません。

リポジトリのフォルダで、入力パスを置き換えて実行します。

```shell
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath "C:\path\to\legacy-workbooks" -DryRun
.\scripts\Convert-XlsToOpenXml.ps1 -SourcePath "C:\path\to\legacy-workbooks"
```

`-SourcePath` は入力フォルダです。出力先を省略すると元の `.xls` と同じ場所に作成します。
`-OutputDirectory` は別の出力先、`-Recurse` はサブフォルダも対象にする指定です。
明示した出力先は単一フォルダとなるため、同名の変換先があれば衝突として報告します。

このスクリプトは `-DryRun` でも Excel を起動し、元ファイルを読み取り専用で開いて VBA と出力形式を確認します。
出力フォルダやブックは作成せず、マクロ・イベント・リンク更新・警告表示は無効にします。
拡張子が厳密に `.xls` のものだけを検査し、既存の `.xlsx` / `.xlsm` は対象外です。
進捗は標準エラー出力、最後の集計は標準出力に短い JSON 1 件で出します。
