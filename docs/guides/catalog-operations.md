# 代理ノートの管理

通常の作成・同期は [README](../../README.md) を参照してください。サムネイル、削除、旧形式の変換など、必要になった操作を説明します。

## Obsidian Bases のカードにサムネイルを表示する

`pull` は Excel に保存されたブックのサムネイルをローカルで PNG に変換し、
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
tkn-excel-note pull --dry-run
tkn-excel-note pull
```

- `--dry-run` は画像変換まで検証しますが、画像・ノート・同期状態を保存しません。
- サムネイルがないブックは `cover` が空になります。画像のないカードも表示されます。
- 手動で設定した `cover` は保持します。自動抽出へ戻す場合は `cover` を空にしてください。
- 自動設定した画像は次回の `pull` で更新し、欠損・破損した出力画像も再生成します。
- 変換できないサムネイルは警告を出し、既存の `cover` と通常のメタデータ同期を保持します。
- 同期状態の `managedCover` で自動設定したリンクを識別します。状態を失った場合、既存の非空の `cover` は手動指定として保持します。
- 元の Excel と過去に出力した画像は削除しません。画像は最大辺 1200 px の PNG です。

カードの画像プロパティはローカル添付へのリンクを受け付けます（[Obsidian 公式ヘルプ](https://help.obsidian.md/bases/views/cards)）。

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
