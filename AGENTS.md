# AGENTS.md

## Project guardrails

- Never delete a workbook or proxy note as part of synchronization.
- The explicit `delete-notes` command may delete selected proxy notes only after
  confirming their workbooks are missing, backing up the notes and synchronization
  state, and removing only the corresponding state entries. Never delete workbooks.
- Preserve unknown Frontmatter fields and text outside managed Markdown markers.
- Preserve all OOXML package entries not explicitly changed.
- Keep progress on stderr and one compact result on stdout.
- Do not put private paths, workbook metadata, or real user data in tracked files.
