# Excel Catalog Pipeline

[日本語](README_ja.md)

`excel-catalog` projects `.xlsx` and `.xlsm` workbooks into Markdown proxy notes and
safely synchronizes selected metadata in both directions. Excel remains the source
workbook; Markdown is a searchable catalog entry and metadata editing surface.

The default for every mutating workflow is a report-only dry run. The tool never
deletes a workbook or proxy note.

## Requirements

- Windows or another Python 3.11+ platform
- [`uv`](https://docs.astral.sh/uv/)
- `.xlsx` or `.xlsm` source workbooks

Excel, Excel COM, and Obsidian are not required for metadata inspection. Obsidian is
only relevant when proxy-note renames must update backlinks.

## Install

The normal installation is editable. Replace
`C:\path\to\tkn_excel_catalog_pipeline` with the actual repository folder.

```console
uv tool install -e "C:\path\to\tkn_excel_catalog_pipeline"
excel-catalog config show
```

With `-e` (`--editable`), the installed `excel-catalog` directly references the
repository source. Normal source changes and `git pull` updates therefore take effect
without reinstalling. The second command verifies the CLI entry point and effective
configuration. You can also run `excel-catalog --version` and `excel-catalog --help`.

Recreate the editable installation after moving or renaming the repository folder,
changing dependencies, or changing package metadata or the CLI entry point in
`pyproject.toml`:

```console
uv tool install -e "C:\path\to\tkn_excel_catalog_pipeline" --force
```

To use a non-editable installation that does not automatically reflect repository
changes, run:

```console
uv tool install "C:\path\to\tkn_excel_catalog_pipeline" --force
```

A non-editable installation uses the code captured at installation time. Run the same
command again after `git pull` or another repository update.

## Configure

Create the user-wide configuration, then edit its example paths:

```console
excel-catalog config init
```

This creates `~/.tkn/excel_catalog_pipeline/config.yaml`. Re-running the command is
safe: an identical file is left unchanged, and an edited file is not overwritten.
`excel-catalog config init --force` explicitly replaces an existing edited file with
the template, so use it only when discarding those edits is intentional.

The main settings are:

| Setting | Meaning |
| --- | --- |
| `schema_version` | Configuration format version; keep `1`. |
| `sources[].id` | A user-chosen, unique, stable name for one Excel source. |
| `sources[].path` | Root folder containing the source Excel workbooks. |
| `sources[].include` | Glob patterns for workbooks below that root. |
| `sources[].notes.root` | Proxy-note folder: `pull` writes here and `push` reads from here. |
| `sources[].profile` | Note-format profile; keep `tkn-obsidian-v1` in the current version. |
| `sources[].rename_adapter` | `report-only` only reports renames; `filesystem` may apply an explicitly authorized rename. |
| `sync.max_extracted_text_chars` | Maximum extracted workbook text stored in a proxy note. |

The remaining `sync` booleans document safety policy for future extension. The
current version always preserves unknown note metadata, never deletes files, and
still requires `push --write-excel --allow-rename` before a workbook rename can be
applied.

For Windows paths, YAML single quotes are recommended, for example
`'C:\path\to\excel-workbooks'`. Inside single quotes, backslashes do not need to be
doubled.

Configuration is merged in this order:

1. `~/.tkn/excel_catalog_pipeline/config.yaml`
2. `./.tkn/config.yaml`
3. a file passed with `--config`

Later files override earlier files. Individual command options have final precedence.
Relative paths are resolved from the current working directory. Inspect the effective
configuration without creating runtime files:

```console
excel-catalog config show
excel-catalog --config C:\path\to\config.yaml config show
```

Source IDs must be unique. The schema supports multiple roots; use `--source <id>` to
limit a run to one root.

## Basic use

Inventory configured workbooks and proxy notes:

```console
excel-catalog status
```

Plan Excel-to-Markdown changes, then apply reviewed note writes:

```console
excel-catalog pull
excel-catalog pull --write-notes
```

Plan Markdown-to-Excel metadata changes, then apply reviewed workbook writes:

```console
excel-catalog push
excel-catalog push --write-excel
```

Per-file `push` statuses:

| Status | Meaning |
| ------ | ------- |
| `unchanged` | The workbook and proxy note have no synchronized differences. Individual logs omit these files; the summary reports only the total. |
| `would-write` | Proxy-note changes are planned for the workbook. Dry-run mode has not written them yet. |
| `written` | The workbook write and post-write verification completed. |
| `missing-source` | No unique workbook matches the proxy note. No workbook is modified. |
| `pull-required` | The workbook has changes that should be reviewed through `pull`; `push` does not modify it. |
| `conflict` | The base, workbook, and proxy note comparison found a conflict. Review it before using `--prefer-note` or `--prefer-source`. |
| `duplicate-id` | More than one workbook has the same `TknExcelCatalogId`, so matching is ambiguous. |
| `rename-required` | A rename was requested but requires explicit permission or handling by the configured adapter. |
| `rename-error` | The rename target, collision checks, or related file operation failed. |
| `read-error` | Workbook or proxy-note discovery or reading failed. |
| `write-error` | The workbook write or post-write verification failed. The command rolls back where possible. |

`sourcePath` is relative to the selected source `path` shown at startup. This avoids
repeating the source root on every line and keeps reports usable if the source root moves.
A `missing-source` result has no matching workbook and therefore no `sourcePath`; its log
shows only the proxy-note filename.

Limit a push to one proxy note:

```console
excel-catalog push --note example.xlsx.md
excel-catalog push --note example.xlsx.md --write-excel
```

Assign missing stable workbook IDs. This changes the OOXML custom properties only in
write mode and creates a backup first:

```console
excel-catalog adopt
excel-catalog adopt --write-excel
```

## Metadata contract

| Markdown proxy note | Excel core property                   |
| ------------------- | ------------------------------------- |
| `title`           | Title                                 |
| `description`     | Comments / description                |
| `nouns[0]`        | Categories / category                 |
| `nouns[1..]`      | Tags / keywords                       |
| `sourceFileName`  | Explicit rename request, not metadata |

Proxy notes are named `<workbook-name>.xlsx.md` or `<workbook-name>.xlsm.md`.
Generated body sections are enclosed by `excel-catalog` markers. Unknown Frontmatter
fields and text outside those markers are preserved.

The generated body structure, headings, section order, and default description are
owned by the application profile at
`src/excel_catalog_pipeline/note_profiles/tkn-obsidian-v1/template.md`. It is shipped
as a package resource and is not a user configuration file. Python owns resource
loading and validation, dynamic workbook content, and safe marker replacement.

## Conflict and rename behavior

Synchronization uses field-level three-way comparison between the previous base,
current workbook, and current note. Different changes on both sides return exit code
`2`; no last-write-wins rule is applied. After review, `--prefer-source` or
`--prefer-note` can resolve that run explicitly.

Workbook rename requires all of the following:

- an edited `sourceFileName`
- `push --write-excel --allow-rename`
- a valid same-extension Windows filename with no collision
- a configured rename adapter capable of coordinating the proxy note

`rename_adapter: report-only` stops with `rename-required`. `filesystem` is suitable
only when direct Markdown rename without backlink maintenance is acceptable. An
Obsidian-specific rename remains report-only until an external adapter is configured.

## Outputs and safety

Run reports are stored under:

```text
~/.tkn/excel_catalog_pipeline/state/runs/<run-id>/
  summary.json
  actions.csv
  details.json
```

Synchronization base state is `~/.tkn/excel_catalog_pipeline/state/sync-state.json`.
Backups are under the adjacent `backups/` directory. Dry runs may create reports, but
do not modify workbooks, notes, synchronization state, or cache.

Workbook writes:

- support `.xlsx` and `.xlsm` only;
- create a backup before replacement;
- rewrite only selected OOXML metadata package entries;
- preserve VBA and unrelated ZIP entries;
- validate ZIP integrity and reread written properties;
- refuse digitally signed OOXML packages;
- fail when a locked source cannot be safely replaced.

Encrypted workbooks, `.xls`, `.xlsb`, deletion synchronization, format conversion,
and workbook layout/shape semantics are outside the MVP. Extracted cell text is a
search aid, not a complete workbook representation.

## Console and exit codes

Human-readable progress goes to stderr as `[LEVEL] message`. For `push`, this includes
the selected source configuration (`id`, workbook path, include patterns, and note
settings), one result line as each non-unchanged file result is determined, and an
indented final summary. Per-file lines use the note filename and relative `sourcePath`
instead of repeating the full note path. Unchanged files appear only as a count in the summary.
One compact JSON result still goes to stdout. Use `--quiet`, `--verbose`, or
`--no-color`; `NO_COLOR` is honored.

- `0`: success
- `1`: execution, validation, or partial write error
- `2`: unresolved conflict
- `3`: configuration or selection error

## Development

```console
uv sync --locked
uv run pytest
uv run ruff check .
uv run mypy src
uv build
```

Tests use synthetic workbooks only. Do not add private paths, real workbook metadata,
credentials, or Vault content to fixtures or documentation.
