from __future__ import annotations

import zipfile
from pathlib import Path


def create_workbook(
    path: Path,
    *,
    title: str = "Example title",
    subject: str = "Example subject",
    creator: str = "Example author",
    description: str = "Example description",
    category: str = "Engineer, Myself",
    keywords: str = "Excel; Catalog",
    created: str = "2025-12-01T00:00:00Z",
    modified: str = "2026-01-01T00:00:00+09:00",
    workbook_id: str = "",
    macro_enabled: bool = False,
) -> Path:
    core = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <dc:title>{title}</dc:title><dc:subject>{subject}</dc:subject>
 <dc:creator>{creator}</dc:creator><dc:description>{description}</dc:description>
 <cp:category>{category}</cp:category><cp:keywords>{keywords}</cp:keywords>
 <dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
 <dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified>
</cp:coreProperties>"""
    content_type = (
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        if macro_enabled
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    )
    content_types = f'''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
 <Override PartName="/xl/workbook.xml" ContentType="{content_type}"/>
 <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
 <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>'''
    root_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>"""
    workbook = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""
    shared = """<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
 <si><t>Catalog entry</t></si>
</sst>"""
    sheet = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c></row></sheetData>
</worksheet>"""
    custom = ""
    if workbook_id:
        custom = f"""<?xml version="1.0" encoding="UTF-8"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
 <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" name="TknExcelCatalogId"><vt:lpwstr>{workbook_id}</vt:lpwstr></property>
</Properties>"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("docProps/core.xml", core)
        if custom:
            archive.writestr("docProps/custom.xml", custom)
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/sharedStrings.xml", shared)
        archive.writestr("xl/worksheets/sheet1.xml", sheet)
        archive.writestr("xl/styles.xml", b"styles-unchanged")
        if macro_enabled:
            archive.writestr("xl/vbaProject.bin", b"synthetic-vba-payload\x00\xff")
    return path
