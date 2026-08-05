---
type: Excel
schemaVersion: "1.0"
title: {{ title }}
description: {{ description }}
nouns: {{ nouns }}
files: {{ files }}
sourceRoot: {{ source_root }}
sourceFileName: {{ source_file_name }}
sourceId: {{ source_id }}
date: {{ date }}
updated: {{ updated }}
noteId: {{ note_id }}
---

# {{ title }}

## Overview

{{ description | Excel workbookの検索・管理用代理ノート。 }}

<!-- excel-catalog:begin workbook-path -->
## Workbook Path

{{ workbook_path }}
<!-- excel-catalog:end workbook-path -->

<!-- excel-catalog:begin workbook-map -->
## Workbook Map

{{ workbook_map }}
<!-- excel-catalog:end workbook-map -->

<!-- excel-catalog:begin extracted-text -->
## Extracted Text

{{ extracted_text }}
<!-- excel-catalog:end extracted-text -->

<!-- excel-catalog:begin excel-metadata -->
## Excel Metadata

{{ excel_metadata }}
<!-- excel-catalog:end excel-metadata -->
