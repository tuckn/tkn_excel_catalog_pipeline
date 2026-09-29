from __future__ import annotations

import json
import zipfile
from dataclasses import replace
from pathlib import Path

import pytest
from tkn_genai_bridge import Profile

import excel_catalog_pipeline.cli as cli
import excel_catalog_pipeline.context as context
import excel_catalog_pipeline.paths as paths
from excel_catalog_pipeline.adapters.markdown import read_note
from excel_catalog_pipeline.adapters.ooxml import read_core_properties
from excel_catalog_pipeline.models import AppConfig, SourceConfig, SyncConfig
from tests.helpers import create_workbook
from tests.test_context import rewrite_package


@pytest.fixture
def setup(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(paths, "app_root", lambda: tmp_path / "application")
    config = AppConfig("1.1.0", (), SyncConfig())
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: config)
    monkeypatch.setattr(context, "resolve_profile", lambda config: Profile())
    real_plan = context.generation_plan
    monkeypatch.setattr(context, "generation_plan", lambda profile, **kwargs: real_plan(profile))
    calls = []

    def render(snapshot, evidence, output, config):
        output.mkdir(parents=True)
        (output / "001.png").write_bytes(b"image")
        return [{"image": "001.png", "range": "$A$1:$B$1", "overview": False}]

    def generate(config, evidence, images, usage, logger, *, on_usage, stage="sheet", **kwargs):
        calls.append((stage, evidence))
        on_usage({"inputTokens": 12, "outputTokens": 4, "durationSeconds": 0.1})
        return (
            "### Summary\n\nWorkbook meaning."
            if stage == "workbook"
            else "### Details\n\nSheet meaning."
        )

    monkeypatch.setattr(context, "render_sheet", render)
    monkeypatch.setattr(context, "generate_markdown", generate)
    book = create_workbook(tmp_path / "books" / "nested" / "book.xlsx")
    output = tmp_path / "notes" / "custom.md"
    return config, book, output, calls


def pull(book, output=None, *options):
    arguments = ["pull", str(book)]
    if output is not None:
        arguments += ["--output", str(output)]
    return cli.main(arguments + list(options))


def change_cell(book):
    with zipfile.ZipFile(book) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").replace(b"<v>42</v>", b"<v>43</v>")
    book.write_bytes(rewrite_package(book.read_bytes(), {"xl/worksheets/sheet1.xml": sheet}))


def test_single_create_update_push_preserves_identity_and_user_text(setup):
    _, book, output, calls = setup
    before = book.read_bytes()
    assert pull(book, output) == 0
    first = read_note(output)
    assert first.frontmatter["type"] == "Excel"
    assert first.frontmatter["contextStatus"] == "not-generated"
    assert book.read_bytes() == before and not calls
    output.write_text(
        output.read_text(encoding="utf-8").replace("title: Example title", "title: Edited title")
        + "\n## My notes\nKeep this.\n",
        encoding="utf-8",
    )
    assert pull(book) == 0
    assert not book.with_name(book.name + ".md").exists()
    assert read_note(output).note_id == first.note_id
    assert "Keep this." in read_note(output).body
    body_before_push = read_note(output).body
    assert cli.main(["push", str(output)]) == 0
    assert read_note(output).body == body_before_push
    with zipfile.ZipFile(book) as archive:
        assert read_core_properties(archive)["title"] == "Edited title"
    assert read_note(output).note_id == first.note_id
    assert "Keep this." in read_note(output).body


def test_ai_creates_overview_and_frontmatter_reuses_cache_then_marks_stale(setup):
    _, book, output, calls = setup
    assert pull(book, output, "--ai") == 0
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "current"
    assert note.frontmatter["contextAnalyzedSheets"] == ["Data"]
    assert "Workbook meaning." in note.body and "Sheet meaning." in note.body
    assert note.body.index("context-workbook") < note.body.index("context-1")
    assert [stage for stage, _ in calls] == ["sheet", "workbook"]
    assert pull(book, None, "--ai") == 0
    assert len(calls) == 2
    change_cell(book)
    assert pull(book) == 0
    assert read_note(output).frontmatter["contextStatus"] == "stale"
    assert "Workbook meaning." in read_note(output).body and len(calls) == 2
    assert pull(book, None, "--ai") == 0
    assert len(calls) == 4 and read_note(output).frontmatter["contextStatus"] == "current"


def test_first_ai_dry_run_does_not_create_note_assets_or_state(setup, capsys):
    _, book, output, calls = setup
    before = book.read_bytes()
    assert pull(book, output, "--ai", "--dry-run") == 0
    result = json.loads(capsys.readouterr().out)
    assert result["usage"]["calls"] == 0
    assert not output.parent.exists() and not paths.app_root().exists() and not calls
    assert book.read_bytes() == before


@pytest.mark.parametrize("phrase", ["Sheet meaning.", "Workbook meaning."])
def test_edited_ai_sections_are_protected_before_writes(setup, phrase):
    _, book, output, calls = setup
    assert pull(book, output, "--ai") == 0
    output.write_text(
        output.read_text(encoding="utf-8").replace(phrase, "Reviewed text.")
        + "\nPersonal annotation.\n",
        encoding="utf-8",
    )
    before = output.read_bytes()
    assert pull(book, None, "--ai") == 1
    assert output.read_bytes() == before and len(calls) == 2
    assert pull(book, None, "--ai", "--force") == 0
    assert "Personal annotation." in output.read_text(encoding="utf-8")


def test_single_note_can_join_batch_without_duplication_or_losing_base(setup, monkeypatch):
    config, book, output, _ = setup
    assert pull(book, output) == 0
    note_id = read_note(output).note_id
    output.write_text(
        output.read_text(encoding="utf-8").replace("title: Example title", "title: My title"),
        encoding="utf-8",
    )
    source = SourceConfig(
        "library",
        book.parent.parent,
        ("*.xlsx",),
        output.parent,
        recursive=True,
        frontmatter_term_format="plain",
    )
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: replace(config, sources=(source,)))
    assert cli.main(["pull", "--source", "library"]) == 0
    note = read_note(output)
    assert note.note_id == note_id and note.frontmatter["title"] == "My title"
    assert note.frontmatter["sourceFileName"] == "nested/book.xlsx"
    assert list(output.parent.rglob("*.md")) == [output]
    state = json.loads(paths.state_path().read_text(encoding="utf-8"))
    assert len(state["entries"]) == 1
    assert cli.main(["push", str(output)]) == 0
    with zipfile.ZipFile(book) as archive:
        assert read_core_properties(archive)["title"] == "My title"


def test_existing_unrelated_markdown_is_not_overwritten_even_with_force(setup):
    _, book, output, calls = setup
    output.parent.mkdir()
    output.write_text("My unrelated document.", encoding="utf-8")
    assert pull(book, output, "--ai", "--force") == 3
    assert output.read_text() == "My unrelated document." and not calls
    assert not paths.state_path().exists()


def test_unknown_sheet_fails_before_first_note_is_written(setup):
    _, book, output, calls = setup
    assert pull(book, output, "--ai", "--sheet", "Missing") == 1
    assert not output.exists() and not paths.state_path().exists() and not calls


def test_ai_failure_keeps_previous_overview_and_marks_stale(setup, monkeypatch):
    _, book, output, _ = setup
    assert pull(book, output, "--ai") == 0
    change_cell(book)
    original_generate = context.generate_markdown

    def fail(*args, **kwargs):
        if kwargs.get("stage") == "workbook":
            raise RuntimeError("overview failed")
        return original_generate(*args, **kwargs)

    monkeypatch.setattr(context, "generate_markdown", fail)
    assert pull(book, None, "--ai") == 1
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "stale"
    assert "Workbook meaning." in note.body


def test_batch_ai_uses_same_note_format(setup, monkeypatch):
    config, book, output, calls = setup
    create_workbook(book.parent / "second.xlsx")
    source = SourceConfig("library", book.parent, ("*.xlsx",), output.parent)
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: replace(config, sources=(source,)))
    assert cli.main(["pull", "--source", "library", "--ai"]) == 0
    notes = list(output.parent.glob("*.md"))
    assert len(notes) == 2 and len(calls) == 4
    assert all(read_note(note).frontmatter["contextStatus"] == "current" for note in notes)


def test_conflicting_changes_do_not_overwrite_either_side(setup):
    _, book, output, _ = setup
    assert pull(book, output) == 0
    output.write_text(
        output.read_text(encoding="utf-8").replace("title: Example title", "title: Note edit"),
        encoding="utf-8",
    )
    create_workbook(book, title="Excel edit")
    before = book.read_bytes(), output.read_bytes()
    assert cli.main(["push", str(output)]) == 2
    assert (book.read_bytes(), output.read_bytes()) == before


def test_application_store_uses_new_name_even_if_legacy_directory_exists(tmp_path, monkeypatch):
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: tmp_path))
    assert paths.app_root() == tmp_path / ".tkn" / "excel_note"
    legacy = tmp_path / ".tkn" / "excel_catalog_pipeline"
    legacy.mkdir(parents=True)
    assert paths.app_root() == tmp_path / ".tkn" / "excel_note"
    assert not (tmp_path / ".tkn" / "excel_note").exists()


def test_old_generation_commands_are_not_available():
    for command in ("export", "context"):
        with pytest.raises(SystemExit) as failure:
            cli.build_parser().parse_args([command])
        assert failure.value.code == 2


def add_second_sheet(book, *, hidden=False):
    with zipfile.ZipFile(book) as archive:
        state = ' state="hidden"' if hidden else ""
        workbook = archive.read("xl/workbook.xml").replace(
            b"</sheets>", f'<sheet name="Second" sheetId="2" r:id="rId3"{state}/></sheets>'.encode()
        )
        relations = archive.read("xl/_rels/workbook.xml.rels").replace(
            b"</Relationships>",
            b'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
        )
        second = archive.read("xl/worksheets/sheet1.xml")
    book.write_bytes(
        rewrite_package(
            book.read_bytes(),
            {
                "xl/workbook.xml": workbook,
                "xl/_rels/workbook.xml.rels": relations,
                "xl/worksheets/sheet2.xml": second,
            },
        )
    )


def test_only_changed_sheet_is_regenerated_and_overview_is_refreshed(setup):
    _, book, output, calls = setup
    add_second_sheet(book)
    assert pull(book, output, "--ai") == 0
    original = context.existing_block(output.read_text(encoding="utf-8"), "2")
    assert len(calls) == 3
    change_cell(book)
    assert pull(book, None, "--ai") == 0
    assert [(stage, evidence["sheet"]) for stage, evidence in calls[3:]] == [
        ("sheet", "Data"),
        ("workbook", "Workbook overview"),
    ]
    assert context.existing_block(output.read_text(encoding="utf-8"), "2") == original
    before = output.read_bytes()
    assert pull(book, None, "--ai") == 0
    assert output.read_bytes() == before and len(calls) == 5


def test_hidden_sheet_is_omitted_until_explicitly_selected(setup):
    _, book, output, calls = setup
    add_second_sheet(book, hidden=True)
    assert pull(book, output, "--ai") == 0
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "partial"
    assert note.frontmatter["contextOmittedSheets"] == ["Second"]
    assert len(calls) == 2
    assert pull(book, None, "--ai", "--sheet", "Data", "--sheet", "Second") == 0
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "current"
    assert note.frontmatter["contextAnalyzedSheets"] == ["Data", "Second"]
    assert len(calls) == 4


@pytest.mark.parametrize("edited", [False, True])
def test_removed_sheet_context_is_retired_but_manual_edits_are_protected(setup, edited):
    _, book, output, _ = setup
    add_second_sheet(book)
    assert pull(book, output, "--ai") == 0
    if edited:
        text = output.read_text(encoding="utf-8")
        block = context.existing_block(text, "2")
        output.write_text(
            text.replace(block, block.replace("Sheet meaning.", "My reviewed meaning.")),
            encoding="utf-8",
        )
    create_workbook(book)
    before = output.read_bytes()
    assert pull(book, None, "--ai") == (1 if edited else 0)
    if edited:
        assert output.read_bytes() == before
        assert pull(book, None, "--ai", "--force") == 0
    note = read_note(output)
    assert "context-2" not in note.body
    assert note.frontmatter["contextStatus"] == "current"


def test_single_command_protects_a_stable_id_used_by_another_existing_workbook(setup):
    _, book, output, _ = setup
    create_workbook(book, workbook_id="same-id")
    assert pull(book, output) == 0
    before = output.read_bytes(), paths.state_path().read_bytes()
    other = create_workbook(book.parent / "copy.xlsx", workbook_id="same-id")
    assert pull(other) == 2
    assert (output.read_bytes(), paths.state_path().read_bytes()) == before
    assert not other.with_name(other.name + ".md").exists()


def test_relative_sheet_listing_needs_no_note_or_state(setup, monkeypatch, capsys):
    _, book, _, calls = setup
    monkeypatch.chdir(book.parent.parent)
    unrelated = book.with_name(book.name + ".md")
    unrelated.write_text("An unrelated document.", encoding="utf-8")
    assert cli.main(["workbook", "list-sheets", "--workbook", "nested/book.xlsx"]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["sheets"][0]["name"] == "Data"
    assert not paths.app_root().exists() and not calls
    assert unrelated.read_text() == "An unrelated document."


def test_plain_pull_never_resolves_ai_connection(setup, monkeypatch):
    _, book, output, _ = setup

    def forbidden(*args, **kwargs):
        raise AssertionError("plain pull accessed the AI connection")

    monkeypatch.setattr(context, "resolve_profile", forbidden)
    assert pull(book, output) == 0
