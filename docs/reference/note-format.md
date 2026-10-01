# ノートの形式

`export` で書き出す Markdown と、同期で作成する代理ノートの Frontmatter・本文の構成を定義します。
ノートを読み取るスクリプトや Obsidian の Bases を作るときに参照してください。
使い方は [README](../../README.md) を参照してください。

## 書き出しと代理ノートの違い

| 観点 | 書き出し（`export`） | 代理ノート（`pull`） |
| --- | --- | --- |
| `type` | `ExcelExport` | `Excel` |
| `schemaVersion` | `"1.0"` | `"2.1"` |
| 保存先 | 既定は各ブックの隣。単体のブックは `--output` で変更可能 | `sources.<id>.notes.dir` |
| 管理マーカー | なし | ツールが管理する本文を `excel-catalog` マーカーで囲む |
| 更新 | `--force` で全体を置き換え | 管理マーカー内と同期対象の Frontmatter だけを更新 |
| `push` の対象 | 対象外 | 対象 |

## 代理ノートの Frontmatter

代理ノートは、同梱のテンプレート `tkn-obsidian-v1` から作成します。
テンプレートにない Frontmatter 項目を追加した場合、ツールはその項目を変更せずに保持します。

| 項目 | 内容 | 更新するコマンド |
| --- | --- | --- |
| `type` | 常に `Excel` | 作成時のみ |
| `schemaVersion` | 代理ノートの形式の版。引用符付きの文字列 `"2.1"` | `pull` |
| `title`、`subject`、`author`、`keywords`、`categories`、`comments` | Excel の文書プロパティ。編集すると `push` で Excel へ反映できます。対応は「[メタデータの対応](synchronization.md#メタデータの対応)」を参照してください。 | `pull` / `push` |
| `description` | ノート全体の概要。Excel には反映しません。 | 利用者 |
| `cover` | カード表示用の画像へのリンク | `pull` |
| `sourceRoot` | source の ID | `pull` |
| `sourceFileName` | 入力フォルダからの相対パス。編集すると名前変更・移動の要求になります。 | `pull` |
| `sourceFullPath` | 元ブックの絶対パス | `pull` |
| `sourceId` | `<source の ID>:<ブックの固定 ID または相対パス>` | `pull` |
| `sourceCreated` / `sourceModified` | Excel の作成日時・更新日時 | `pull` |
| `date` / `updated` | ノートの作成日時・更新日時 | `pull` |
| `noteId` | ノートの識別子（UUID）。作成後は変わりません。 | 作成時のみ |
| `contextStatus` | AI の説明の状態。値の意味は「[説明の状態](../guides/sheet-content.md#説明の状態)」を参照してください。 | `pull` |

`--context` で説明を生成すると、次の項目も追加します。

| 項目 | 内容 |
| --- | --- |
| `contextAnalyzedSheets` | 説明があるシートの一覧 |
| `contextOmittedSheets` | 説明をまだ生成していないシートの一覧 |
| `contextStaleSheets` | 説明の生成後に内容が変わったシートの一覧 |
| `contextUnverifiedSheets` | 手直しや記録の欠落により、Excel との対応を確認できないシートの一覧 |
| `contextGeneratedAt` | ブック要約の生成日時 |
| `contextSourceFingerprint` | ブック要約の生成時点のブックの内容のハッシュ値 |

いずれも、直近の実行で選んだシートだけでなく、それまでに生成した説明全体を表します。
これらの項目は Excel へ書き戻しません。

`type`、日時、`noteId` は引用符なしの YAML の値、`schemaVersion` は引用符付きの文字列で出力します。
`keywords` と `categories` は、`sources.<id>.notes.frontmatter_term_format` が `obsidian-link` なら `[[...]]` 形式のリンク、`plain` なら通常の文字列の一覧です。
`files` 項目は同期に使いません。
既存のノートにある場合は、その値を保持します。

## 書き出しの Frontmatter

| 項目 | 内容 |
| --- | --- |
| `type` / `schemaVersion` | `ExcelExport` / `"1.0"` |
| `title`、`subject`、`author`、`keywords`、`categories`、`comments` | Excel の文書プロパティ。タイトルが空の場合はファイル名（拡張子なし） |
| `sourceFileName` | 拡張子付きのファイル名 |
| `sourceFullPath` | 元ブックの絶対パス |
| `sourceCreated` / `sourceModified` | Excel の作成日時・更新日時 |
| `sourceSnapshotSha256` | 読み取ったブックのファイルの SHA-256 ハッシュ値 |
| `generatedAt` | 書き出した日時（UTC） |
| `generationMethod` | `--context` ありなら `context`、なしなら `text` |
| `contextStatus` | AI の説明の状態。`--context` なしでは `not-generated` |
| `extractedSheets` | 文字を抽出したワークシートの一覧 |
| `unsupportedSheets` | チャートシートなど、抽出に対応していないシートの一覧 |

`--context` を指定した場合は、代理ノートと同じ `context` で始まる項目も記録します。
`sourceId`、`noteId`、`sourceRoot` は持たないため、同期の対象として識別されません。

## 本文の構成

`--context` で説明を生成した代理ノートは、次の構成になります。
書き出しの場合も同じ構成で、管理マーカーだけを除きます。

```markdown
# <タイトル>

## ブック要約

## シート

### <シート名>
使用profile: default-ja

#### シート要約
#### 結論
#### 要点
#### 内容
##### <内容に応じた見出し>
#### 不確実な点
#### 出典画像

## Extracted Text

## Workbook Map
```

見出しの文言と言語は、文章の profile のテンプレートが定義します。
上記は既定の `default-ja` の場合です。

- シートは、ブック内のシートの並び順に配置します。
- 「シート要約」は必ず出力します。「結論」「要点」「内容」「不確実な点」は、該当する内容がなければ省略します。
- 「内容」の下の見出しは、シートの内容に応じて AI が構成します。
- 「出典画像」には、AI に送ったシート画像へのリンクを置きます。画像はノートと同じフォルダの `img/` にあり、相対パスで参照します。
- `--context` を使わない場合、本文は「Workbook Map」と「Extracted Text」だけです。

「Extracted Text」には、ワークシートごとに、セル位置付きの保存値、数式、読み取れる図形の文字を出力します。
非表示のシートも対象です。
数式は再計算しません。
代理ノートでは、1ブックあたりの文字数を `sync.max_extracted_text_chars`（既定 12,000 文字）で打ち切ります。
書き出しでは打ち切りません。

## Workbook Map

Workbook Map は、シートごとの範囲や要素の数を示す表です。
保存済みのブックのファイルから直接集計するため、Excel の起動や AI は使いません。
非表示のシートも集計し、文字の抽出上限には影響されません。

| 列 | 意味 |
| --- | --- |
| `Sheet` / `ID` / `State` | シート名、シートの内部 ID、表示状態（`visible`、`hidden` など） |
| `Stored range` | ブックに保存されている使用範囲。書式だけのセルを含む場合があります。 |
| `Content range` | 値または数式があるセルを囲む最小の矩形。空白だけの文字列も値として扱います。 |
| `Populated cells` | 値または数式があるセルの数。0、`FALSE`、結果が保存されていない数式も数えます。書式だけのセルは数えません。 |
| `Tables` | Excel で定義したテーブルの数。見た目だけの表は数えません。 |
| `Shapes` | 図形とコネクタの数。グループは、中の要素を個別に数えます。 |
| `Images` / `Charts` | シート上に配置した画像とグラフの数。同じ画像を複数か所に配置した場合は、それぞれ数えます。 |
| `Notes` | 取得に失敗した項目や、対応していない項目の理由 |

値の表記は次のとおりです。

| 表記 | 意味 |
| --- | --- |
| `0` | 対象が存在しないことを確認しました。 |
| `—` | 範囲が空です。 |
| `unknown` | 取得できなかった、または対応していません。 |

VML、OLE オブジェクト、フォームコントロールなど、数え方に対応していない要素を検出したシートでは、部分的な数を総数と誤解しないよう、図形・画像・グラフの数を `unknown` にします。
チャートシートなど、ワークシート以外のシートは未対応として記録します。

## 同期記録のシート情報

同期では、Workbook Map と同じ集計結果を、同期記録 `~/.tkn/excel_note/state/sync-state.json` のブックごとの `sheetInventory` にも保存します。

```json
{
  "schemaVersion": 1,
  "sheets": [
    {
      "name": "手順",
      "sheetId": "1",
      "state": "visible",
      "path": "xl/worksheets/sheet1.xml",
      "stored_range": "A1:H40",
      "content_range": "B2:G38",
      "populated_cells": 52,
      "tables": 0,
      "shapes": 14,
      "images": 0,
      "charts": 0,
      "warnings": []
    }
  ]
}
```

`path` はブックのファイル（ZIP 形式）内のシートの位置です。
取得できなかった値は `null`、空であることを確認した範囲は空文字列です。
上記の値は説明用の例です。
