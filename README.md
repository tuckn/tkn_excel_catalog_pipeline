# Excel Catalog Pipeline — Excel を Markdown で検索・整理する

Excel ブックを元の形式で保ちながら、内容の検索や分類に使う Markdown ノートを作成する CLI です。
`.xlsx` / `.xlsm` からタイトル・作成者などのメタデータ、シート一覧、セルの文字情報を取り出します。
Markdown 側で編集したメタデータを Excel へ戻すこともできます。

初めて使う場合は、[対象範囲](#対象範囲)から[最初のノートを作成する](#最初のノートを作成する)まで順に進めてください。
導入後は[日常の更新](#日常の更新)と[コマンド一覧](#コマンド一覧)から必要な操作を探せます。

## 得られる結果

たとえば、`example.xlsx` に製品比較のシートがある場合、`pull` で `example.xlsx.md` を作成できます。
以下は生成ノートの内容を示す架空の抜粋です。実際のノートには識別子や管理用のマーカーも入ります。

```markdown
---
type: Excel
schemaVersion: "2.1"
title: 製品比較
subject: 次期モデルの検討
author: Example Author
keywords:
  - "[[製品比較]]"
sourceFileName: example.xlsx
---

# 製品比較

## Workbook Map

（シートの一覧）

## Extracted Text

（セルから抽出した比較項目や説明文）
```

ノートを Obsidian やテキスト検索で探し、`sourceFullPath` から元の Excel ファイルを確認できます。
ノート先頭の YAML 領域（Frontmatter）で `subject` や `keywords` を編集し、`push` で Excel のプロパティへ反映できます。
セルや図形の編集は Excel で行います。

## 対象範囲

| 目的 | 操作 | 得られるもの・変更対象 |
| --- | --- | --- |
| Excel を検索・分類する | `pull` | メタデータ、シート一覧、抽出文字を含む Markdown ノート |
| ノートで整理したメタデータを Excel に戻す | `push` | Excel のプロパティ更新とバックアップ |
| シートの配置や図形の意味も文章にする | 任意の `context build` | シート画像を使った AI の説明文と、その根拠画像 |
| 作成・確認済みのシート説明を取り込む | 任意の `context import` | 元の文章を保った説明文と画像のコピー |

ブックに対応して作る Markdown ファイルを、この文書では**代理ノート**と呼びます。
入力フォルダとノートの保存先を組にした設定が **source** です。
`--source` には、その組を識別する `sources[].id` を指定します。

通常の同期はローカルで完結し、Excel 本体や Obsidian の起動、認証、ネットワーク通信、生成 AI を必要としません。
`context build` は選択シートの情報を Codex サービスへ送信します。
利用条件は[AI でシートの説明を追加する](#ai-でシートの説明を追加する)で確認してください。

主 CLI の対象は `.xlsx` と `.xlsm` です。
暗号化ブック、`.xls`、`.xlsb`、削除の同期には対応しません。
セルの文字抽出は検索補助であり、書式・図形・配置を含むブック全体の完全な再現ではありません。
旧 `.xls` は、[変換用スクリプト](#旧-xls-を変換する)で別途変換できます。

## セットアップ

### 必要なもの

- Python 3.11 以上と `uv`。
- 入力にする `.xlsx` / `.xlsm` と、代理ノートを保存するフォルダ。
- このリポジトリのローカルコピー。

以下の手順は Windows の PowerShell 用です。
基本の同期処理は Excel COM を使いませんが、この手順では他 OS 上の実動作確認までは扱いません。
デスクトップ版 Microsoft Excel は、任意の `context build` と旧 `.xls` の変換に必要です。

### インストールする

例示パスを、このリポジトリの実際の保存先へ置き換えて実行します。

```powershell
cd "C:\path\to\tkn_excel_catalog_pipeline"
uv tool install .
tkn-excel-catalog --help
```

ヘルプが表示されれば、CLI の起動を確認できています。
通常のインストールは、その時点のコードと同梱ファイルを `uv` のツール環境へ格納します。
リポジトリを編集しただけではインストール済み CLI に反映されません。
更新時の操作は[更新と移行](#更新と移行)を参照してください。

### 設定ファイルを作成する

ユーザー共通の設定ファイルを作成します。

```powershell
tkn-excel-catalog config init
```

保存先は `~/.tkn/excel_catalog_pipeline/config.yaml` です。
`~` はユーザーのホームフォルダを表し、Windows では `$HOME` に相当します。
既存ファイルが同梱テンプレートと同じなら変更せず、編集済みなら上書きせず停止します。

### 入力と出力先を設定する

作成した設定をエディタで開きます。

```powershell
notepad "$HOME\.tkn\excel_catalog_pipeline\config.yaml"
```

以下は初回利用に必要な設定ファイル全体の最小例です。
2 つの `C:\path\to\...` を実際の入力フォルダとノート保存先へ置き換えてください。
同梱テンプレートをそのまま使う場合も、`sources` 内の該当項目を編集すれば始められます。

```yaml
schema_version: 1
sources:
  - id: personal-excel
    path: 'C:\path\to\excel-workbooks'
    notes:
      root: 'C:\path\to\catalog-notes'
```

この例では、入力フォルダ直下の `.xlsx` / `.xlsm` が対象です。
サブフォルダも含める場合は[探索範囲の設定](#探索範囲の設定)で `recursive: true` を指定します。
`personal-excel` は任意の一意な設定名です。変更した場合は、以降の `--source personal-excel` も同じ名前にします。

### 読み込まれた設定を確認する

設定を保存し、入力・出力先が意図した値になっているか確認します。

```powershell
tkn-excel-catalog config show
```

最も優先度が高い設定ファイルの絶対パスに続けて、解決済みの設定を JSON で表示します。
`loadedConfigFiles` は、実際に読み込んだ設定ファイルの一覧です。
作業フォルダに `.tkn/config.yaml` があるとユーザー共通設定より優先されるため、ここで対象パスを確認してください。
読み込み順の詳細は[設定の優先順位](#設定の優先順位)にあります。

## 最初のノートを作成する

**通常の `pull`、`push`、`adopt` は書き込みを行います。変更せず予定を調べる場合は `--dry-run` を付けます。**
最初は `pull` で代理ノートを作成します。この操作は元の Excel ファイルを変更しません。

まず、対象と作成予定のノートを確認します。

```powershell
tkn-excel-catalog pull --source personal-excel --dry-run
```

新しいブックには `would-create` が表示されます。
対象パスが正しく、読み取りエラーや競合がなければ通常実行します。

```powershell
tkn-excel-catalog pull --source personal-excel
```

`notes.root` に `<ブック名>.xlsx.md` または `<ブック名>.xlsm.md` が作成されます。
たとえば `example.xlsx` には `example.xlsx.md` が対応します。
完了後は、次の点を確認してください。

1. 画面の集計で `created` または `updated` の件数を確認し、エラー・競合がないことを確かめます。
2. 表示された `notePath` のノートを開き、タイトル、`Workbook Map`、`Extracted Text` を確認します。
3. 詳細が必要な場合は、最後に表示された `summary.json` と同じフォルダの `actions.csv` を開きます。

差分がなければ `unchanged` になります。
対象が 0 件でも処理を終えるため、最初の実行でノートができない場合は `path`、`recursive`、`include`、`ignore` を見直してください。
終了コードだけで全対象の同期完了とは判断せず、[結果の読み方](#結果の読み方)も確認します。

## 日常の更新

この CLI は呼び出したときに一度処理し、常駐監視はしません。
`--source` を省略した同期コマンドは、設定済みの全 source を対象にします。

### Excel 側の変更を取り込む

Excel を保存してから、代理ノートを更新します。

```powershell
tkn-excel-catalog pull --source personal-excel
```

ノートだけで編集されたメタデータは保持します。
両方で同じ項目を別の値へ変更していた場合は、該当ブックを競合として報告します。
他の対象で成功した処理は残るため、全体が一括で取り消されるわけではありません。

### ノート側の変更を Excel に戻す

代理ノートの Frontmatter を編集し、対象を確認してから反映します。
以下の `example.xlsx.md` は、実際に作成されたノート名に置き換えます。

```powershell
tkn-excel-catalog push --source personal-excel --note "example.xlsx.md" --dry-run
```

`would-write` の対象と差分を確認したら、書き込みます。

```powershell
tkn-excel-catalog push --source personal-excel --note "example.xlsx.md"
```

元ブックをバックアップしてからプロパティを更新し、書き込み後に検証します。
反映できる項目は[メタデータの対応](#メタデータの対応)を参照してください。
`--note` は繰り返し指定でき、ノートのパス・名前、`sourceFileName`、`sourceId` で絞り込めます。
省略時は、選択した source の全代理ノートが対象です。

### 追跡状況を確認する

```powershell
tkn-excel-catalog status --source personal-excel
```

未追跡のブックやノート、重複 ID、未対応ファイル、読み取りエラーを表示します。
ブックとノートは変更しませんが、実行レポートを保存します。
追跡済みブックは件数のみ、確認が必要な項目は分類ごとに最大 20 件の相対パスを表示します。
全件の詳細はレポートで確認できます。

### 再実行・中断後の確認

同期済みの内容は再実行で判定し直します。
通常実行では差分がなくても実行レポートを作成します。
`--dry-run` は設定・入力・既存の同期記録を読み、変更方向、競合、パス、保護条件を確認しますが、ブック、ノート、同期記録、キャッシュ、バックアップ、レポートを書きません。
アプリケーションの一時ファイルも残しません。

中断・失敗後は表示されたエラーとレポートを確認し、同じ対象を `--dry-run` で再確認します。
ロック、競合、権限などの原因を解消してから通常実行してください。
プレビューは確認した時点の状態なので、通常実行時には入力を読み直し、衝突や保護条件を再確認します。

## コマンド一覧

| 目的 | コマンド・主な引数 | 詳細 |
| --- | --- | --- |
| 共通設定を作る | `config init` | [設定ファイルの作成](#設定ファイルを作成する) |
| 有効な設定を読む | `config show` | [設定の優先順位](#設定の優先順位) |
| 追跡状況を調べる | `status --source <id>` | [追跡状況](#追跡状況を確認する) |
| Excel からノートへ反映する | `pull --source <id>` | [初回実行](#最初のノートを作成する) |
| ノートから Excel へ反映する | `push --source <id> --note <note>` | [日常の更新](#ノート側の変更を-excel-に戻す) |
| ブックに固定 ID を付ける | `adopt --source <id>` | [名前変更と移動](#名前変更と移動) |
| シート名を調べる | `context sheets --source <id> --workbook <path>` | [AI による説明](#ai-でシートの説明を追加する) |
| 選択シートの説明を生成する | `context build --source <id> --workbook <path> --sheet <name>` | [AI による説明](#ai-でシートの説明を追加する) |
| 既存の説明を取り込む | `context import --source <id> --workbook <path> --sheet <name> --markdown <path>` | [Markdown の取り込み](#確認済みの-markdown-を取り込む) |

全体のオプションはサブコマンドの前に置きます。
たとえば詳細表示は `tkn-excel-catalog -v pull --dry-run`、設定の指定は `tkn-excel-catalog --config "C:\path\to\config.yaml" config show` です。
各コマンドの引数は `tkn-excel-catalog push --help` などで確認できます。

## 設定の詳細

### 設定の優先順位

次の順で読み込み、後の値を優先します。

1. 組み込みの既定値。
2. `~/.tkn/excel_catalog_pipeline/config.yaml`。
3. 現在の作業フォルダにある `.tkn/config.yaml`。
4. 全体オプション `--config` で指定したファイル。

辞書は項目ごとに統合しますが、`sources` などのリストは後のファイルの値で全体を置き換えます。
相対パスは設定ファイルの場所ではなく、現在の作業フォルダを基準に解決します。
個別の CLI オプションによる指定は、その対象となる動作で優先されます。
`config show` はファイルを作成せず、読み込んだファイルがない場合は組み込み値を使っていることを表示します。

### 探索範囲の設定

| 設定 | 意味・省略時の動作 |
| --- | --- |
| `schema_version` | 設定形式の版。`1` を指定します。 |
| `sources[].id` | 必須。source の一意な名前です。継続利用中は安定した名前を使います。 |
| `sources[].path` | 必須。入力ブックのルートフォルダです。 |
| `sources[].recursive` | 既定は `false`。`true` でサブフォルダも探索します。 |
| `sources[].include` | 既定は `["*.xlsx", "*.xlsm"]`。対象のパターンです。空リストは指定できません。 |
| `sources[].ignore` | 既定は `[]`。入力ルートからの相対パスに適用する除外パターンです。 |
| `sources[].notes.root` | 必須。`pull` の出力先であり、`push` の入力元です。 |
| `sources[].notes.profile` | 既定は `tkn-obsidian-v1`。同梱のノート形式を使います。 |
| `sources[].notes.frontmatter_term_format` | 既定は `obsidian-link`。`keywords` と `categories` を `[[用語]]` にします。`plain` は通常の文字列です。 |
| `sources[].notes.rename_adapter` | 既定は `report-only`。`filesystem` でファイルの直接移動を許可します。[名前変更と移動](#名前変更と移動)を参照してください。 |
| `sync.max_extracted_text_chars` | 既定は `12000`。1 ブックから抽出する検索用文字情報の最大文字数です。正の整数を指定します。 |

`sources` が空のままでは同期を実行できません。
未知の設定キーや、不正な型・値はエラーになります。
その他の `sync` の真偽値は将来の拡張用であり、現実装は未知のノート情報を保持し、ブックやノートを削除しません。
`sync.allow_source_rename` は現在の移動許可の判定には使いません。

以下は、既存設定の `sources` を置き換える例です。ほかの設定は保持してください。
パターンの区切りは OS に関係なく `/` を使います。
Windows パスは YAML のシングルクォートで囲むと、バックスラッシュを二重化せずに記載できます。

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
      root: 'C:\path\to\catalog-notes'
      profile: tkn-obsidian-v1
      frontmatter_term_format: obsidian-link
      rename_adapter: report-only
```

この設定では、たとえば `2026/example.xlsx` のノートは `notes.root` 配下の `2026/example.xlsx.md` になります。
`sourceFileName` も `2026/example.xlsx` です。
追跡開始後のブック移動にノートを追従させる場合は、[名前変更と移動](#名前変更と移動)の条件が必要です。

## 同期の仕組みと保護

### メタデータの対応

| 代理ノートの項目 | Excel 側の値・用途 |
| --- | --- |
| `title` | タイトル（Title） |
| `subject` | 件名（Subject） |
| `author` | 作成者（Author / creator）。単一文字列です。 |
| `keywords[]` | キーワード（Keywords / tags） |
| `categories[]` | 分類（Category） |
| `comments` | コメント（Comments / description） |
| `sourceCreated` | Excel の作成日時。`pull` だけで更新します。 |
| `sourceModified` | Excel の更新日時。`pull` だけで更新します。 |
| `sourceFileName` | 入力ルートからの相対パス。編集すると名前変更・移動の要求になります。 |
| `sourceFullPath` | 元ブックの絶対パス。名前変更の要求には使いません。 |
| `description` | 代理ノート全体の概要。Excel には反映しません。 |
| `schemaVersion` | 代理ノートの形式の版。現在は文字列 `"2.1"` です。 |

`keywords` と `categories` は、Excel へ戻すときに Obsidian のリンク記法を外し、`; ` で結合します。
カンマは用語の一部として保持します。
`push` ではノート側の `sourceCreated` / `sourceModified` を Excel へ書かず、ノートの現在値も保持します。

生成する `type`、元ファイルの日時、`date`、`updated`、`noteId` は引用符なしの YAML 値、`schemaVersion` は引用符付き文字列です。
本文の管理部分は `excel-catalog` マーカーで囲みます。
未知の Frontmatter 項目と、マーカー外の手書き本文は保持します。
旧 `Overview` / `Workbook Path` の移行だけは、[ノート形式の移行](#ノート形式の移行)に従います。

### 競合の判定と解決

たとえば前回のタイトルが「製品比較」で、Excel だけを「製品比較 2026」に変えた場合は `pull` の対象です。
ノートだけで変えた場合は `push` の対象です。
両方で別のタイトルに変えた場合は、更新日時だけで優先側を決めず、競合として報告します。

この判定には、前回 Excel とノートが一致した値を保存した同期記録（`sync-state.json`）を使います。
現在の Excel、現在のノート、前回の一致値を項目ごとに比べる方式です。
通常実行後も、実際に一致した項目だけを基準値として更新します。
反対方向の未反映の変更は、次の操作に引き継ぎます。
同期記録がなく、既存の Excel とノートの値が異なる場合も競合になります。

競合の内容は、全体オプション `-v` またはレポートの `differences.csv` で確認します。
Excel を正としてノートをそろえる場合は、次の順で実行します。

```powershell
tkn-excel-catalog -v pull --source personal-excel --prefer-source --dry-run
tkn-excel-catalog pull --source personal-excel --prefer-source
```

**`pull --prefer-source` は、ノートだけで編集した項目も Excel の値で置き換えます。**
通常の `pull` はその編集を保持します。
ノートを正として競合を解決する場合は、`push --prefer-note` を指定します。
いずれも対象と差分を確認してから使い、`--prefer-source` と `--prefer-note` は同時に指定しません。

### 名前変更と移動

ブックを移動しても識別しやすくするため、`adopt` で未付与のブックに固定 ID（カスタムプロパティ `TknExcelCatalogId`）を付けられます。
初回の `pull` の必須手順ではありません。
通常実行はバックアップ後にブックを書き換えるため、ID の付与が目的の場合に実行します。

```powershell
tkn-excel-catalog adopt --source personal-excel --dry-run
tkn-excel-catalog adopt --source personal-excel
```

名前変更・移動には、次の 2 方向があります。

| 先に変更する場所 | CLI が変更する対象 | 必要な指定 |
| --- | --- | --- |
| エクスプローラーなどで元ブックを移動する | 代理ノートを元ブックの相対パスへ追従させる | `rename_adapter: filesystem` と通常の `pull` |
| ノートの `sourceFileName` を編集する | 元ブックと代理ノートを移動する | `rename_adapter: filesystem` と通常の `push --allow-rename` |

既定の `report-only` はパスを変更せず、必要な移動を `rename-required` として報告します。
`filesystem` の設定だけでは移動せず、対応するコマンドの通常実行が必要です。
プレビューには `--dry-run` を追加します。`push` のプレビューでも `--allow-rename` を残すと、通常実行と同じ移動先の検証を行えます。

移動先は入力ルート内の有効な相対パスで、拡張子を保ち、既存ファイルと衝突しない必要があります。
`filesystem` は Obsidian の被リンクを更新しません。
被リンクの維持が必要なら `report-only` で報告を確認し、Obsidian 上での手動操作を検討してください。

現在のパスと理由は画面に表示し、通常実行では `actions.csv` と `details.json` にも記録します。
`pull` の希望移動先と衝突情報は `details.json`、`push` の希望相対パスは編集した `sourceFileName` で確認します。
プレビューではレポートを作らないため、項目単位の比較には `-v` を使います。

### Excel への書き込みを保護する仕組み

- 上書き前に元ブックをバックアップします。
- OS のアプリケーション用一時フォルダで OOXML の更新内容を組み立て、対象のメタデータだけを書き換えます。
- VBA と、変更対象ではない ZIP 内の項目を保持します。
- 上書き前後に ZIP の整合性とプロパティを読み直して検証します。
- ファイル自体を別ファイルに置き換えず内容を上書きし、ファイルの識別情報と作成日時を維持します。
- 上書きや事後検証に失敗した場合は、バックアップの内容と日時の復元を試みます。
- デジタル署名付き OOXML への書き込みは拒否します。ロックなどで安全に上書きできない場合も失敗として扱います。

## 結果の読み方

### 画面表示と出力形式

進捗や結果の集計は、標準エラー出力へ `[LEVEL] message` の形で表示します。
`pull` は変更などがあるノートの絶対 `notePath` と相対 `sourcePath` を表示します。
`push` は冒頭に入力・ノート設定を表示し、処理結果をノート名と相対 `sourcePath` で表示します。
`unchanged` は個別表示せず、最後の集計に件数を出します。
`missing-source` は元ブックが見つからないため `sourcePath` がなく、ノート名だけを表示します。

標準出力の形式はコマンドで異なります。

| コマンド | 標準出力 | 保存される実行結果 |
| --- | --- | --- |
| `status`、通常の `pull` / `push` / `adopt` | JSON を出しません。画面用の結果は標準エラー出力です。 | 実行レポート |
| `pull` / `push` / `adopt` の `--dry-run` | 同上 | 保存しません。予定・理由・競合を画面で確認します。 |
| `config show` | 設定ファイルの見出しと、インデント付き JSON | 保存しません。出力全体は単独の JSON ではありません。 |
| `config init`、`context` の各コマンド | 1 件のコンパクトな JSON | 設定やシート処理の結果は各節を参照してください。 |

`-v` / `--verbose` は Excel・ノート・基準値・予定方向の詳細を表示します。
`--quiet` は情報ログを抑制し、`--no-color` または環境変数 `NO_COLOR` は色を無効にします。
これらはサブコマンドの前に指定します。

### 状態と次の操作

| 状態 | 意味・次の操作 |
| --- | --- |
| `unchanged` | 同期対象の差分がありません。 |
| `would-create` / `would-update` | `pull --dry-run` でのノート作成・更新予定です。対象を確認して通常実行します。 |
| `created` / `updated` | ノートを作成・更新しました。 |
| `would-write` / `written` | Excel への書き込み予定／書き込みと検証の完了です。 |
| `would-adopt` / `adopted` | 固定 ID の付与予定／付与の完了です。 |
| `missing-source` | ノートに一意に対応するブックがありません。元ファイルの場所を確認します。 |
| `missing-note` | 追跡済みのノートがありません。`pull` の作成予定を確認します。 |
| `pull-required` / `push-required` | 反対方向に未反映の変更があります。該当方向のプレビューで確認します。 |
| `conflict` | 両側の変更や基準値の欠落で競合しています。[競合の解決](#競合の判定と解決)を参照します。 |
| `duplicate-id` | 複数のブックに同じ固定 ID があります。対応付けを確認します。 |
| `rename-required` | 名前変更・移動に必要な設定や許可を確認します。 |
| `rename-error` | 移動先の名前、衝突、ファイル操作の失敗を確認します。 |
| `read-error` / `write-error` | 読み取り／書き込み・事後検証に失敗しました。表示理由とバックアップを確認します。 |

終了コードは、通常 `0` が成功、`1` が実行・検証・部分的な書き込みエラー、`2` が未解決の競合、`3` が設定・対象選択などのエラーです。
引数の構文エラーも引数解析処理が `2` を返すため、画面のメッセージと合わせて判断します。
また、`rename-required` などが残っていても終了コード `0` になることがあります。
`statusCounts` と対象ごとの状態を確認してください。

### レポートと同期記録の保存先

実行レポートは、ノートの内容とは別の、処理結果の記録です。
`status` と通常の同期コマンドは、次のファイルを作成します。

```text
~/.tkn/excel_catalog_pipeline/state/runs/<run-id>/
  summary.json
  actions.csv
  details.json
  differences.csv
```

| ファイル | 確認できること |
| --- | --- |
| `summary.json` | 全体の状態、変更・競合・エラー件数、状態別件数 |
| `actions.csv` | ファイルごとの結果、対象パス、変更項目、理由 |
| `details.json` | ファイルごとの詳細情報 |
| `differences.csv` | 項目ごとの前回一致値 `baseValue`、Excel 値 `excelValue`、ノート値 `noteValue`、変更方向 |

`sourceToNoteFields` は Excel からノート、`noteToSourceFields` はノートから Excel への変更項目です。
`changedFields` は、そのコマンドで反映した、または反映予定の項目です。
`differences.csv` の `direction` は比較による変更方向、`plannedDirection` は優先指定も含めた適用予定方向です。
空欄で値を消す変更も、ここで確認できます。

全体オプション `--report-dir` でレポートの親フォルダを変更できます。
同期の `--dry-run` ではこの指定は無効で、レポートを作りません。
`context` にも適用せず、シート処理専用の保存先を使います。

継続利用するデータは、次のように扱います。

| 保存領域 | 内容・失った場合の影響 |
| --- | --- |
| 入力の `path` | 元の Excel ブックです。代理ノートだけでは元の内容を復元できません。 |
| `notes.root` | 代理ノートと、シート説明の画像です。手書き本文や取り込み済みの説明を含むため、元ブックと別に保持します。 |
| `~/.tkn/excel_catalog_pipeline/config.yaml` | 入出力先などの設定です。移行時は設定を引き継いでパスを確認します。 |
| `state/sync-state.json` | ブック・ノートの対応と前回一致値です。失うと、既存の差分を競合として扱うことがあります。手編集や安易な削除を避けてください。 |
| `state/backups/` | Excel 更新前のバックアップです。過去の内容への復旧に使います。 |
| `state/runs/` | 各実行のレポートです。削除するとその実行の調査記録を失います。 |
| `state/context/` | シート説明の保護・再利用用の記録、取り込み時のバックアップ、使用量の記録です。欠落時は既存の説明の置き換えを保護することがあります。 |

上表の `state/` は `~/.tkn/excel_catalog_pipeline/state/` です。
環境を移す場合は、元ブック、ノートと画像、設定、同期・シート処理の記録を一緒に保全してください。
ノートには元ファイルの絶対パスや抽出した本文が入り、レポートにもパスやメタデータが入ります。
実データを含む生成物を公開リポジトリへ追加しないでください。

## AI でシートの説明を追加する

`context build` は、Excel の描画画像とセル・図形の文字・位置情報から、シートの説明を作成します。
既存の代理ノートに、シートごとの説明と画像へのリンクを追加します。
通常の `pull` から自動実行されることはありません。

**選択シートの情報を、ログイン済みの Codex サービスへ送信します。**
送信できる内容か確認してから実行してください。
利用枠や料金はアカウント・契約によって異なり、この CLI は金額を見積もりません。

### 準備と実行

Windows、デスクトップ版 Microsoft Excel、インストール・ログイン済みの Codex CLI、画像入力に対応するモデルが必要です。
同梱設定のモデルは `gpt-5.6-sol`、推論量は `medium` です。
実際に利用できるモデルを設定し、画像化用の追加依存を含めてインストールします。

```powershell
cd "C:\path\to\tkn_excel_catalog_pipeline"
uv tool install ".[context]" --reinstall
tkn-excel-catalog context --help
```

事前に `pull` で代理ノートを作成してください。
以下の `example.xlsx` と `Solutions` は、対象ブックの相対パスと実際のシート名に置き換えます。
まず保存済みのシート一覧を確認します。この操作は読み取り専用です。

```powershell
tkn-excel-catalog context sheets --source personal-excel --workbook "example.xlsx"
```

生成前の検証は、次のコマンドで行います。

```powershell
tkn-excel-catalog context build --source personal-excel --workbook "example.xlsx" --sheet "Solutions" --dry-run
```

保存済みブック、シート選択、既存ノート、再利用・手動編集の保護条件を確認します。
ファイルの保存、Excel の起動、画像化、認証、AI 呼び出しは行いません。
画像数とトークン数は、この段階では確定できません。

説明を生成してノートへ追加します。

```powershell
tkn-excel-catalog context build --source personal-excel --workbook "example.xlsx" --sheet "Solutions"
```

`--sheet` は完全一致の名前で繰り返し指定できます。
`--all-sheets` は表示中の全シートを選び、非表示シートは名前を明示した場合だけ対象になります。
未対応のシート種別は AI を呼ぶ前に拒否します。
相対ブックパスは選択 source の入力ルート基準です。絶対パスでも、その範囲内で `include` / `ignore` の条件を満たす必要があります。
全シートを事前検証した後に順番に処理し、最初の失敗で停止します。先に成功したシートは残ります。

生成後はノート内の説明を元シートと見比べてください。
小さな文字、矢印の接続先、埋め込みオブジェクト、説明のない色などは誤って解釈される場合があります。
生成指示では原文と推測を区別させますが、結果の正確さを保証するものではありません。

### 保存済みの内容と画像化

Excel を開いたままでも、最後に保存された内容を読み取ります。
画面上の未保存の編集は対象外です。
Windows の共有読み取りを使い、取得中の保存を検出した場合は再実行を案内します。

一時コピーを別の Excel インスタンスで開き、マクロ・イベント・リンク更新、自動再計算、起動時の外部データ更新を抑止して描画します。
元ブックやユーザーの Excel を変更・終了せず、一時コピーは保存せず閉じます。
Excel 4 マクロシートや、UTF-8 以外のブック・接続情報 XML は画像化できません。

表示中のセル、結合セル、図形から描画範囲を決め、UsedRange や印刷範囲外の図形も対象にします。
詳細画像は一部を重ねて分割し、上から下、同じ高さでは左から右に並べます。
全体図も付け、必要なら全体図の PDF ページを連結します。
極端に離れた領域は座標付きの詳細画像にし、詳細領域が想定外に複数ページとなる場合は欠落を避けるため停止します。
AI には領域を Z 字順に読み、矢印・色・配置を解釈するよう指示します。

### 保存先と手動編集の保護

各シートの文章は `context-<sheetId>` 管理セクションに保存します。
Frontmatter、セクション外の文章、他シートの説明は保持します。通常の `pull` もこの説明を保持します。
画像と抽出根拠は、ノートと同じ階層の固定の `img` フォルダに保存します。

```text
example.xlsx.md
img/
  <book-key>/sheet-<id>/<generation>/
    001.png
    002.png
    evidence.json
```

Markdown 内のリンクは相対パスです。移動時はノートと `img` を一緒に扱ってください。
サブフォルダのノートも、そのノートと同階層の `img` を使います。
識別子でブック・シート間の衝突を防ぎ、古い画像は自動削除しません。
PDF とブックのコピーは一時ファイル、PNG と抽出根拠の JSON は保存対象です。

シート内容、関連する内部データ、設定、生成指示の版が同じで、画像と生成本文が保持されていれば、前回の結果を再利用します。
この場合、Excel の起動も AI 呼び出しもなく、トークン消費は 0 です。
共有書式の変更は複数シートの再生成につながる場合があります。無関係なセル文字列の変更は通常影響しません。

`--force` は再利用せずに生成し、**指定したシートの管理セクション内の手動編集も置き換えます。**
通常は手直しや対応する記録の欠落を検出すると既存本文を保護します。
生成中にノートが編集された場合は、`--force` でも上書きせず停止します。

処理記録と使用量は `~/.tkn/excel_catalog_pipeline/state/context/` に保存します。
ノート単位のロックで同時生成を防ぎます。強制終了後のロックは、処理が動いていないことを確認してから除去してください。
AI 呼び出しは自動再試行しません。失敗・タイムアウト・不正な応答では既存ノートを保持します。
保存途中の失敗で未使用の画像が残る場合があります。

### 使用量と生成設定

呼び出しごとの入力・出力・キャッシュ入力・推論トークンと所要時間を表示し、結果 JSON にその実行の合計を含めます。
キャッシュ入力は入力の内数、推論トークンは出力の内数なので重複加算しません。
`turn.completed` の使用量だけを合計し、累積イベントは加算しません。
不明・中断時は `unknown` / `null` とし、判明した部分的な使用量は記録します。
使用量の記録には生成指示の本文、生成本文、認証情報を保存しません。

設定ファイルのトップレベル `context` で生成条件を変更します。省略時は以下の値です。

| キー（`context.` 以下） | 既定値 | 意味・単位 |
| --- | --- | --- |
| `executable` | `codex` | 呼び出す Codex CLI |
| `model` | `gpt-5.6-sol` | 画像入力に対応するモデル |
| `reasoning_effort` | `medium` | `low` / `medium` / `high` / `xhigh` |
| `language` | `Japanese` | 生成文の言語 |
| `timeout_seconds` | `600` | AI 呼び出しのタイムアウト秒数 |
| `max_images` | `24` | 1 シートの全体図を含む画像数の上限。最大 `100` |
| `tile_width_points` / `tile_height_points` | `1200` / `800` | 詳細画像の幅・高さ。Excel のポイント単位 |
| `overlap_points` | `80` | 隣接画像の重なり。幅・高さの小さい方の半分未満 |
| `image_dpi` | `150` | 画像の解像度。`72`～`300` dpi |
| `max_cells` / `max_objects` | 各 `10000` | 1 シートの抽出対象セル・オブジェクト数の上限 |
| `max_input_chars` | `200000` | シートの抽出根拠・送信内容の文字数上限 |
| `max_workbook_mb` | `100` | 入力ブックのサイズ上限。1 単位は 1024 × 1024 バイト |

数値項目は正の整数です。
上限を超えた場合は内容を切り捨てずエラーにします。対象と必要量を確認してから設定を変更してください。
条件の変更により、前回の生成結果を再利用できなくなる場合があります。

## 確認済みの Markdown を取り込む

`context import` は、作成・確認済みのシート説明を、見出しと画像リンクの調整だけで代理ノートへ取り込みます。
Excel の起動、画像化、AI 呼び出し、`[context]` の追加依存は不要です。
事前に `pull` で代理ノートを作成してください。

以下のブック・シート名と Markdown パスを実際の値に置き換え、まず検証します。

```powershell
tkn-excel-catalog context import --source personal-excel --workbook "example.xlsx" --sheet "Solutions" --markdown "C:\path\to\Solutions.context.md" --dry-run
```

問題がなければ取り込みます。

```powershell
tkn-excel-catalog context import --source personal-excel --workbook "example.xlsx" --sheet "Solutions" --markdown "C:\path\to\Solutions.context.md"
```

入力は `# タイトル` 形式の H1 が 1 つある Markdown にします。
H1 は `## <シート名> (sheetId: <id>)`、H2 は H3 というように 1 段下げます。
入力 H6 は変換できないため拒否します。コードブロックとインラインコードは保持します。
新しいセクションは `Workbook Map` の後、`Extracted Text` の前へ置き、既存セクションは同じ位置で更新します。

相対画像リンクと参照形式の画像は `img/<book-key>/sheet-<id>/import-<hash>/` にコピーします。
ローカル画像は入力 Markdown のフォルダ配下に必要です。絶対パス、範囲外の参照、未対応のローカル形式は拒否します。
HTTP(S)、メール、見出しへのリンクは取得せず保持します。
入力 Markdown と画像は変更しません。

入力 Frontmatter は `provenance.json` に出典として保存し、代理ノートの ID やメタデータ全体を置き換えません。
入力にブック・シートの識別情報があれば照合します。
記録されたブックのハッシュが現在と異なる場合は警告し、過去の出典情報として保持します。

更新前の代理ノートはシート処理の記録フォルダへバックアップします。
同一内容の再取り込みは変更なしです。欠けた画像は再取り込みで復元できます。
取り込み後に編集された本文や画像を置き換える場合は `--force` が必要です。
`--dry-run` はノート・画像・バックアップ・処理記録・ロックを作成しません。

取り込んだ説明はブックが変わっても通常の `context build` で上書きせず、`retained` として保持します。
その場合の AI 呼び出しとトークン消費は 0 です。
**`context build --force` は、取り込んだ説明を AI の生成文で置き換えます。**

## 更新と移行

### インストール済み CLI を更新する

リポジトリの更新後は再インストールし、起動と版を確認します。

```powershell
cd "C:\path\to\tkn_excel_catalog_pipeline"
uv tool install . --reinstall
tkn-excel-catalog --help
tkn-excel-catalog --version
```

AI によるシート説明を使う場合は、インストール対象を `".[context]"` にします。
`--reinstall` でコード・同梱ファイル・依存を反映します。
`--force` はツールのインストール自体を強制する指定で、同じ版の再構築を目的とした代替にはしません。
旧実行名 `excel-catalog` を使っている場合も再インストールし、呼び出し名を `tkn-excel-catalog` へ変更します。

### 古い実行手順を見直す

`0.2.0` 以降、オプションなしの `pull`、`push`、`adopt` は書き込みを行います。
`0.1.x` の書き込みなしの動作に依存したスクリプトやタスクスケジューラでは、確認だけの呼び出しに `--dry-run` を付けてください。
旧 `--write-notes` / `--write-excel` は現在も互換用に受け付けますが、非推奨の警告を出します。
通常実行と同じ書き込みを行い、`--dry-run` とは併用できません。

`config init --force` は、編集済みの共通設定を同梱テンプレートで置き換えます。
通常の更新では既存設定を保持し、新しい設定項目が必要な場合だけ[同梱設定](src/excel_catalog_pipeline/resources/config.example.yaml)を参照してください。

### ノート形式の移行

現行形式は `schemaVersion: "2.1"` です。
版のないノートや、旧 `noun` / `nouns` を使う形式も読み取れます。
通常のノート書き込みで現行形式へ移行し、`--dry-run` はノートを変更しません。
メタデータは Frontmatter に置き、旧管理セクション `Excel Metadata` は生成せず、書き込み時に除去します。

通常の同期によるノート書き込み、または `context import` では、旧 `Overview` と `Workbook Path` を移行します。
`Overview` の独自文章は `description` に移し、既存の概要と違う場合は追記して保持します。
既知の汎用的な下書き文は除去します。
元ブックの絶対パスは `sourceFullPath` に移し、パス節の補足文章は本文に残します。旧節からの読み取りにも対応します。

`context import --description "概要"` は、概要を指定内容で明示的に置き換えます。
通常の取り込みで変更する Frontmatter は `description`、`sourceFullPath`、`schemaVersion` だけで、それ以外の項目・コメント・日時・ID を保持します。
対象となったノートだけを移行し、Vault 全体を自動で一括変更しません。

### 旧 `.xls` を変換する

Windows 用の [Convert-XlsToOpenXml.ps1](scripts/Convert-XlsToOpenXml.ps1) は、デスクトップ版 Excel を使って `.xls` を変換します。
元ファイルを残し、VBA を含む場合は `.xlsm`、それ以外は `.xlsx` にします。
既存の変換先は上書きしません。

リポジトリのフォルダで、入力パスを置き換えて実行します。

```powershell
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
旧 `-Write` は現在も互換用に受け付ける非推奨の指定で、通常実行と同じ動作です。`-DryRun` とは併用できません。

## 開発と検証

開発用依存をそろえ、合成データでテストします。

```powershell
cd "C:\path\to\tkn_excel_catalog_pipeline"
uv sync --locked
uv run pytest
uv run ruff check .
uv run mypy src
uv build
```

コード編集をツール環境へ反映しながら開発する場合は、通常のインストールと別に editable 方式を選べます。

```powershell
uv tool install -e . --reinstall
```

依存関係、パッケージ情報、実行コマンドの定義を変更した場合は再インストールしてください。
テストには架空のブックだけを使い、個人の実パス、実ブックの情報、認証情報、Vault の内容を追加しません。

仕様を調べたり変更したりする場合は、次のファイルを入口にしてください。

| 対象 | 確認先 |
| --- | --- |
| コマンドと画面出力 | [cli.py](src/excel_catalog_pipeline/cli.py) |
| 設定の検証と既定値 | [config.py](src/excel_catalog_pipeline/config.py)、[models.py](src/excel_catalog_pipeline/models.py)、[同梱設定](src/excel_catalog_pipeline/resources/config.example.yaml) |
| 同期と競合判定 | [pipeline.py](src/excel_catalog_pipeline/pipeline.py) |
| ノートの Frontmatter・見出し・節の順序 | [同梱テンプレート](src/excel_catalog_pipeline/note_profiles/tkn-obsidian-v1/template.md) |
| シート説明の生成・取り込み | [context.py](src/excel_catalog_pipeline/context.py)、[context_import.py](src/excel_catalog_pipeline/context_import.py) |
| 検証例 | [tests](tests/) |
| 利用許諾 | [LICENSE](LICENSE) |

ノートのテンプレートはパッケージに含める管理用ファイルです。ユーザー設定としては編集しません。
Python 側が動的なブック内容の生成、形式の検証、管理マーカーの置き換えを担当します。
