# 設定リファレンス

初回設定は [README のセットアップ](../../README.md#セットアップ)にあります。ここでは、入力範囲や保存先を変えるときに使う現行の設定をまとめます。

## 設定ファイルの優先順位

CLI は次の順に設定を読み、後の値を優先します。

1. 組み込みの既定値。
2. ユーザー共通の `~/.tkn/excel_note/config.yaml`。
3. 実行時の作業フォルダにある `.tkn/config.yaml`。
4. `--config` で指定したファイル。

`generation` などの設定項目は重ねられます。`sources` と `include` / `ignore` のような一覧は、上位ファイルで指定するとその項目全体を置き換えます。同じ source の設定を複数ファイルに分けて部分的に追加することはできません。
相対パスは設定ファイルの場所ではなく、コマンドを実行した作業フォルダから解決します。
`tkn-excel-note config show` で読み込まれたファイルと有効な値を確認できます。

保存領域は `~/.tkn/excel_note/` です。旧領域を使っていた場合は、[移行手順](../../README.md#旧版から更新する)に沿ってフォルダ全体を名前変更します。実際の設定ファイルは `config show` で確認できます。
単体ファイルの `pull <workbook>` / `push <note>` に `sources` 登録は不要です。設定済み source に属する場合はその設定を使い、未登録の単体入力では用語を `plain` で出力します。

## 探索範囲

| 設定 | 動作 |
| --- | --- |
| `schema_version` | 同梱例の値は `"1.2.0"` です。 |
| `sources.<id>` | 入力ブックと代理ノートの保存先をまとめる名前です。`--source <id>` で選びます。 |
| `sources.<id>.workbooks_dir` | 入力ブックのルートフォルダです。 |
| `sources.<id>.recursive` | 既定は `false`。`true` にするとサブフォルダも調べます。 |
| `sources.<id>.include` | 既定は `['*.xlsx', '*.xlsm']`。対象にするブックのパターンです。 |
| `sources.<id>.ignore` | 既定は空。対象から外すパスのパターンです。 |
| `sources.<id>.notes.dir` | `pull` が代理ノートを保存し、`push` が読み取るフォルダです。 |
| `sources.<id>.notes.profile` | 既定は `tkn-obsidian-v1`。同梱のノート形式です。 |
| `sources.<id>.notes.frontmatter_term_format` | 既定は `obsidian-link`。`keywords` と `categories` を Obsidian リンクにします。通常文字列にするには `plain` を指定します。 |
| `sources.<id>.notes.rename_adapter` | 既定は `report-only`。パス変更を実行するには `filesystem` を設定します。条件は[名前変更と移動](synchronization.md#名前変更と移動)を参照してください。 |
| `sync.max_extracted_text_chars` | 既定は `12000`。1 ブックから代理ノートへ抽出する文字情報の最大文字数です。 |

`include` と `ignore` のパターンでは、入力ルートからの相対パスに `/` を使います。`ignore` に `archive/**` を指定すると `archive` 以下を除外できます。
`recursive: true` の場合は、入力ルートからの相対フォルダを代理ノート側にも再現します。たとえば `2026/example.xlsx` は `notes.dir` 以下の `2026/example.xlsx.md` になります。

以下は、サブフォルダを含めて探索し、一部を除外する場合の `sources` 部分の設定例です。

```yaml
sources:
  workbooks:
    workbooks_dir: 'C:\path\to\excel-workbooks'
    recursive: true
    include:
      - "*.xlsx"
      - "*.xlsm"
    ignore:
      - "archive/**"
      - "**/Temp/**"
    notes:
      dir: 'C:\path\to\excel-notes'
      profile: tkn-obsidian-v1
      frontmatter_term_format: obsidian-link
      rename_adapter: report-only
```

Windows パスを YAML のシングルクォートで囲むと、バックスラッシュを二重化せずに書けます。

## 同期と削除の設定

同梱テンプレートには `sync.pull_preserves_user_metadata`、`sync.delete_missing_notes`、`sync.delete_missing_workbooks`、`sync.allow_source_rename` が含まれています。これらは現行コードで設定値として読み取られますが、同期処理の分岐には使われません。値を変更してもノートの手書き項目の保持、ファイルの自動削除、名前変更の許可は切り替わりません。
通常の同期は Excel ブックや代理ノートを自動削除せず、未知の Frontmatter 項目と管理マーカー外の本文を保持します。元ブックのない代理ノートを削除するときだけ、[明示的な `delete-notes`](../guides/catalog-operations.md#元ブックがない代理ノートを削除する)を使います。
名前変更の許可は `sources.<id>.notes.rename_adapter` と `push --allow-rename` で指定します。

## cover の設定

トップレベルの `cover` は、単体・source・全 source の `pull` に共通で適用します。
シート画像方式を既定にする最小設定は次のとおりです。

```yaml
cover:
  mode: sheet
```

| キー | 組み込み値 | 意味 |
| --- | --- | --- |
| `cover.mode` | `auto` | `auto` は最後に成功した方式を継承。記録がなければ `embedded`。`sheet` でシート画像、`embedded` で埋め込み画像を指定します。 |
| `cover.sheet` | `null` | 保存済みのシート選択を継承。新規は先頭の表示ワークシート。文字列でシート名を指定します。 |
| `cover.range` | `null` | 保存済みの範囲を継承。新規は `A1:Q50`。 |
| `cover.width` | `null` | 保存済みの幅を継承。新規は 2400 px。600～4000 の整数。 |

値の決定順は CLI の明示指定 → 設定ファイルの非 null 値 → ブックごとの最後の成功記録 → 初回既定値です。
`mode` は `auto` のときに記録を継承します。設定ファイルの `mode: embedded` を明示すると既存の自動生成 cover も埋め込み方式へ切り替えます。手動 cover は保護します。
`null` は初回値へのリセットではありません。全条件を初回値に戻すには、一度 `--cover embedded` を成功させてから `--cover sheet` を実行し、設定ファイルの個別指定も外してください。
`range` や `width` を設定ファイルに指定すると、過去に CLI でブックごとに選んだ条件より優先されます。個別指定を維持したい場合は `mode: sheet` だけにします。
シート方式は Windows とデスクトップ版 Excel が必要です。`generation` の AI 設定は不要です。[利用方法と画像の仕様](../guides/catalog-operations.md#シートの指定範囲から-cover-を作る)を参照してください。

## AI 生成の設定

`pull --context` は `generation` セクションを使います。AI の接続先、認証、モデルは GenAI Bridge の共有設定で管理します。[生成時の設定と保存内容](../guides/sheet-content.md#準備と実行)を参照してください。
