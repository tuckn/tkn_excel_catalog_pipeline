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
`export <file-or-folder>` に `sources` 登録は不要です。入力と出力先は引数だけで決まり、`sources`・`sync`・`cover` の設定を使いません。共通の `generation` 設定は抽出上限・AI 接続・profileに使います。
`pull` / `push` / `status` / `adopt` / `delete-notes` は設定済み source が対象で、`--source ID` が必須です。

## 探索範囲

| 設定 | 動作 |
| --- | --- |
| `schema_version` | 同梱例の値は `"2.1.0"` です。 |
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

トップレベルの `cover` は、選択した source の `pull` に適用します。`export` は cover を生成しません。
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

`export --context` と `pull --source ID --context` は `generation` セクションを使います。AI の接続先、認証、モデルは GenAI Bridge の共有設定で管理します。[生成時の設定と保存内容](../guides/sheet-content.md#準備と実行)を参照してください。


## contextのprofile

generator を選択しない場合、`generation.prompt_profile` の既定は `default-ja`（日本語）です。英語は `default-en` を指定します。生成言語はprofile内の `template.md` / `workbook-template.md` の `language` によって決まり、それぞれのプロンプトの `{{language}}` に渡されます。見出しはテンプレート本文・`labels` が定義します。ブック本文やOSによる言語の自動判定は行いません。

```yaml
generation:
  prompt_profile: default-en
```

`generation.language` と `prompt_profile: auto` / `--profile auto` は廃止しました。残っていると設定または引数エラーになります。[旧言語設定からの移行](../guides/sheet-content.md#旧言語設定からの移行)を参照してください。旧スキーマの読み込み条件も同節に記載しています。

`generation.prompt_profile` は文章のprofile名（既定 `default-ja`）、`generation.profile_dirs` はユーザー定義profileを置いた親フォルダの一覧（既定 `[]`）です。CLIの `export --context --profile <名前>` / `pull --source ID --context --profile <名前>` が、選択したgeneratorまたは共通設定のprofileより優先します。`profile_dirs` は上位の設定で一覧全体を置き換え、相対パスは実行時の作業フォルダから解決します。

```yaml
generation:
  prompt_profile: technical-notes
  profile_dirs:
    - 'C:\path\to\profiles'
  bridge_profile: codex-default
```

この例では `C:\path\to\profiles\technical-notes\` を読みます。指定順で最初の同名フォルダを使い、なければ同梱の `default-ja` / `default-en` を探索します。profileの検証は `--context` で行い、通常のpullやconfig showはAIを呼び出しません。構成ファイル・出力形式・変更時の再利用範囲は[contextのprofile](../guides/sheet-content.md#contextのprofile)を参照してください。


## 名前付きgeneratorと補足文

AI 接続の `bridge_profile`、文章の `prompt_profile`、Bridge の `overrides`、共通の補足文を名前付きgeneratorとして保存できます。
`profile_dirs`、画像分割、抽出上限などは `generation` 直下の共通設定です。

```yaml
schema_version: "2.1.0"
generation:
  default_generator: my-codex-def
  generators:
    my-codex-def:
      bridge_profile: codex-default
      prompt_profile: default-ja
      overrides:
        timeout_seconds: 600
      reference:
        text: |
          技術調査や設計検討のノートを扱います。
        files:
          - ./references/technical-terms.md
    my-codex-en:
      bridge_profile: codex-default
      prompt_profile: default-en
  max_reference_chars: 20000
sources:
  workbooks:
    workbooks_dir: 'C:\path\to\workbooks'
    notes:
      dir: 'C:\path\to\notes'
    generation:
      generator: my-codex-def
      reference:
        text: "このsourceの案件背景を記載します。"
        files:
          - ./references/project-background.md
```

| 項目 | 選択・合成の規則 |
| --- | --- |
| generator | CLI `--generator` → `sources.<id>.generation.generator` → `generation.default_generator`。未指定なら従来の共通設定を使います。 |
| 文章のprofile | CLI `--profile` → 選択したgeneratorの `prompt_profile`。generator未選択なら共通設定の `prompt_profile`。 |
| 補足文 | 選んだgeneratorの text・files → sourceの text・files → CLIの `--reference` → `--reference-file` を追加します。別generatorの補足は含めません。 |
| 補足の無効化 | `--no-reference` は全補足を無効化し、補足ファイルも開きません。直接文・ファイル指定との併用はエラーです。 |
| ファイルの相対パス | 設定内の `reference.files` は、その値を書いた `config.yaml` のフォルダが基準です。CLIは実行フォルダが基準です。 |
| 文字数の上限 | `generation.max_reference_chars` は1つの生成入力に渡す全補足本文の合計文字数。既定20000。超過時はAI呼び出し・書き込み前にエラーとなり、切り詰めません。 |

`export` はsource設定を参照しません。既定generatorまたは `--generator` と、CLIの補足指定を使います。
`--generator`、`--reference`、`--reference-file`、`--no-reference` は `--context` が必要です。
直接文とファイルは併用でき、それぞれ複数回指定できます。ファイルは空でないUTF-8テキストで、BOMにも対応します。

generator内で省略した値は `bridge_profile: codex-default`、`prompt_profile: default-ja`、`overrides: {}`、補足なしです。
名前付きgeneratorを選ぶと、その接続・profile・overridesを使います。従来の共通 `bridge_profile` / `prompt_profile` / `overrides` は、generatorを選ばない場合に使います。
Bridge共有設定との上書き関係は従来どおりです。

設定ファイルのレイヤー間ではgeneratorを名前で合成し、同じgeneratorの未指定項目は下位設定を継承します。
`reference.text` は明示した上位値で置き換え、`reference.files` は一覧全体を置き換えます。`text: ""` / `files: []` でその項目を空にできます。
`sources` は従来どおりレイヤー間で全体を置き換えます。未知の項目、未定義generator、型の違いはエラーです。
`config show` は補足設定を表示しますが、補足ファイル本文の読み取りやAI呼び出しは行いません。

`schema_version: "2.0.0"` と対応する旧1系の設定は引き続き読み込めます。結果は `"2.1.0"` に正規化し、既存ファイルは自動で書き換えません。
