"""The DOCX logic. The property that matters most: accepted text excludes deletions."""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest
from conftest import REVISION, para
from lxml import etree

from docx_mcp import docx_tools


def deleted(text: str) -> str:
    return f"<w:del {REVISION}><w:r><w:delText>{text}</w:delText></w:r></w:del>"


def inserted(text: str) -> str:
    return f"<w:ins {REVISION}><w:r><w:t>{text}</w:t></w:r></w:ins>"


def keep(text: str) -> str:
    return f"<w:r><w:t>{text}</w:t></w:r>"


def only_text(path: str) -> str:
    return docx_tools.read_accepted_text(path, min_length=0)[0]["text"]


class TestAcceptedText:
    def test_deleted_text_is_not_returned(self, docx):
        body = f"<w:p>{keep('The method was ')}{deleted('utterly ')}{keep('sound.')}</w:p>"
        assert only_text(docx(body)) == "The method was sound."

    def test_a_w_t_inside_a_deletion_is_excluded_too(self, docx):
        body = (
            f"<w:p>{keep('Keep ')}<w:del {REVISION}><w:r><w:t>DROPME </w:t></w:r></w:del>"
            f"{keep('this.')}</w:p>"
        )
        assert only_text(docx(body)) == "Keep this."

    def test_a_deletion_nested_inside_an_insertion_is_excluded(self, docx):
        body = f"<w:p><w:ins {REVISION}>{keep('Kept ')}{deleted('gone ')}</w:ins></w:p>"
        assert only_text(docx(body)) == "Kept "

    def test_moved_away_text_is_excluded_and_moved_in_text_is_kept(self, docx):
        body = (
            f"<w:p><w:moveFrom {REVISION}>{keep('old place')}</w:moveFrom></w:p>"
            f"<w:p><w:moveTo {REVISION}>{keep('new place')}</w:moveTo></w:p>"
        )
        rows = docx_tools.read_accepted_text(docx(body), min_length=0)
        assert [r["text"] for r in rows] == ["", "new place"]
        assert all(r["has_tracked_change"] for r in rows)

    def test_insertions_are_returned(self, docx):
        body = f"<w:p>{keep('A ')}{inserted('new ')}{keep('result.')}</w:p>"
        assert only_text(docx(body)) == "A new result."

    def test_a_tracked_change_is_reported_on_its_paragraph(self, docx):
        body = para("Clean paragraph.") + f"<w:p>{keep('Edited ')}{inserted('here.')}</w:p>"
        rows = docx_tools.read_accepted_text(docx(body), min_length=0)
        assert rows[0]["has_tracked_change"] is False
        assert rows[1]["has_tracked_change"] is True

    def test_tabs_and_breaks_keep_words_apart(self, docx):
        body = (
            "<w:p><w:pPr><w:tabs><w:tab w:val='left' w:pos='720'/></w:tabs></w:pPr>"
            "<w:r><w:t>one</w:t><w:tab/><w:t>two</w:t><w:br/><w:t>three</w:t></w:r></w:p>"
        )
        assert only_text(docx(body)) == "one\ttwo\nthree"

    def test_hyperlink_text_is_included(self, docx):
        body = f"<w:p>{keep('See ')}<w:hyperlink>{keep('the report')}</w:hyperlink></w:p>"
        assert only_text(docx(body)) == "See the report"


class TestStructure:
    def test_table_cell_paragraphs_are_included(self, docx):
        body = para("Before.") + f"<w:tbl><w:tr><w:tc>{para('In a cell.')}</w:tc></w:tr></w:tbl>"
        rows = docx_tools.read_accepted_text(docx(body), min_length=0)
        assert [r["text"] for r in rows] == ["Before.", "In a cell."]

    def test_a_text_box_is_its_own_paragraph_and_is_not_counted_twice(self, docx):
        box = f"<w:txbxContent>{para('In a box.')}</w:txbxContent>"
        body = (
            "<w:p><w:r><w:t>Anchor.</w:t>"
            "<mc:AlternateContent>"
            f"<mc:Choice>{box}</mc:Choice><mc:Fallback>{box}</mc:Fallback>"
            "</mc:AlternateContent></w:r></w:p>"
        )
        rows = docx_tools.read_accepted_text(docx(body), min_length=0)
        assert [r["text"] for r in rows] == ["Anchor.", "In a box."]

    def test_an_empty_body_returns_no_rows(self, docx):
        assert docx_tools.read_accepted_text(docx("")) == []


class TestSplitRuns:
    def test_text_is_joined_across_runs(self, docx):
        assert only_text(docx(para("clean ", "aviation ", "fuel"))) == "clean aviation fuel"

    def test_a_term_spanning_runs_is_found(self, docx):
        hits = docx_tools.find_term(docx(para("clean avi", "ation fuel")), r"aviation fuel")
        assert [h["match"] for h in hits] == ["aviation fuel"]


class TestFindTerm:
    def test_it_is_case_insensitive_by_default(self, docx):
        assert len(docx_tools.find_term(docx(para("The Method was sound.")), "method")) == 1

    def test_flags_can_be_overridden(self, docx):
        assert docx_tools.find_term(docx(para("The Method was sound.")), "method", flags="") == []

    def test_several_flags_combine(self, docx):
        path = docx(para("The Method"))
        assert len(docx_tools.find_term(path, "the.method", flags="IGNORECASE, DOTALL")) == 1

    def test_an_unknown_flag_name_is_ignored(self, docx):
        hits = docx_tools.find_term(
            docx(para("The Method")), "method", flags="IGNORECASE,NOT_A_FLAG"
        )
        assert len(hits) == 1

    def test_a_non_flag_attribute_of_re_is_not_used_as_a_flag(self, docx):
        hits = docx_tools.find_term(docx(para("The Method")), "Method", flags="COMPILE,SEARCH")
        assert len(hits) == 1

    def test_limit_caps_the_result_set(self, docx):
        assert len(docx_tools.find_term(docx(para("term " * 30)), "term", limit=5)) == 5

    def test_a_match_inside_a_deletion_is_not_found(self, docx):
        body = f"<w:p>{keep('Kept. ')}{deleted('quarantine')}</w:p>"
        assert docx_tools.find_term(docx(body), "quarantine") == []

    def test_context_surrounds_the_match(self, docx):
        path = docx(para("A" * 80 + " NEEDLE " + "B" * 80))
        hit = docx_tools.find_term(path, "NEEDLE", context_chars=20)[0]
        assert "NEEDLE" in hit["context"]
        assert len(hit["context"]) < 80

    def test_quote_variants_match_straight_and_curly(self, docx):
        path = docx(para("the “word” and the 'word'"))
        pattern = f"{docx_tools.QUOTE_VARIANTS}word{docx_tools.QUOTE_VARIANTS}"
        assert len(docx_tools.find_term(path, pattern)) == 2

    def test_a_bad_pattern_raises(self, docx):
        with pytest.raises(re.error):
            docx_tools.find_term(docx(para("x")), "(")


class TestPagination:
    @pytest.fixture
    def path(self, docx) -> str:
        return docx("".join(para(f"Paragraph number {i} here.") for i in range(10)))

    def test_start_and_limit_window_the_result(self, path):
        rows = docx_tools.read_accepted_text(path, start=2, limit=3, min_length=0)
        assert [r["para_index"] for r in rows] == [2, 3, 4]

    def test_the_index_is_the_document_position(self, path):
        rows = docx_tools.read_accepted_text(path, start=7, limit=5, min_length=0)
        assert rows[0]["para_index"] == 7

    def test_min_length_filters_short_paragraphs(self, docx):
        path = docx(para("hi") + para("a paragraph long enough to survive"))
        assert len(docx_tools.read_accepted_text(path, min_length=10)) == 1
        assert len(docx_tools.read_accepted_text(path, min_length=0)) == 2

    def test_paragraph_range(self, path):
        rows = docx_tools.list_paragraph_range(path, 1, 3)
        assert [r["para_index"] for r in rows] == [1, 2]
        assert rows[0]["char_count"] == len(rows[0]["text"])

    def test_paragraph_range_can_keep_empty_paragraphs(self, docx):
        path = docx(para("One.") + para("") + para("Two."))
        assert len(docx_tools.list_paragraph_range(path, 0, 3)) == 2
        assert len(docx_tools.list_paragraph_range(path, 0, 3, include_empty=True)) == 3


class TestStatsAndVariants:
    def test_stats(self, docx):
        body = para("First paragraph here.") + f"<w:p>{keep('Second ')}{inserted('edited.')}</w:p>"
        stats = docx_tools.get_docx_stats(docx(body))
        assert stats["paragraph_count"] == 2
        assert stats["word_count_approx"] == 5
        assert stats["tracked_change_paragraph_count"] == 1
        assert stats["has_tracked_changes"] is True
        assert stats["comment_count"] == 0
        assert stats["file_size_kb"] > 0

    def test_stats_count_comments_from_their_own_part(self, docx):
        assert docx_tools.get_docx_stats(docx(para("Text."), comments=3))["comment_count"] == 3

    def test_word_count_excludes_deleted_words(self, docx):
        body = f"<w:p>{keep('one two ')}{deleted('three four')}</w:p>"
        assert docx_tools.get_docx_stats(docx(body))["word_count_approx"] == 2

    def test_variants_are_counted_with_locations(self, docx):
        path = docx(para("colour and colour again") + para("one color here"))
        out = docx_tools.check_term_variants(path, ["colour", "color"])
        assert out["counts"] == {"colour": 2, "color": 1}
        assert out["locations"] == {"colour": [0], "color": [1]}
        assert out["total_paragraphs"] == 2

    def test_variants_can_be_case_sensitive(self, docx):
        path = docx(para("Colour colour"))
        assert docx_tools.check_term_variants(path, ["colour"])["counts"]["colour"] == 2
        strict = docx_tools.check_term_variants(path, ["colour"], case_sensitive=True)
        assert strict["counts"]["colour"] == 1

    def test_a_term_is_matched_literally_not_as_a_regex(self, docx):
        path = docx(para("a.b and axb"))
        assert docx_tools.check_term_variants(path, ["a.b"])["counts"]["a.b"] == 1

    def test_variant_counting_ignores_deleted_text(self, docx):
        body = f"<w:p>{keep('colour ')}{deleted('colour colour')}</w:p>"
        assert docx_tools.check_term_variants(docx(body), ["colour"])["counts"]["colour"] == 1


class TestScanMetaNotes:
    @pytest.mark.parametrize("prefix", ["Note:", "TODO:", "FIXME:", "REVIEW:", "[internal]"])
    def test_each_default_prefix_is_recognised(self, docx, prefix):
        hits = docx_tools.scan_meta_notes(docx(para(f"{prefix} working text here.")))
        assert len(hits) == 1
        assert hits[0]["para_index"] == 0

    def test_ordinary_prose_is_not_flagged(self, docx):
        path = docx(para("The results indicate a clear downward trend."))
        assert docx_tools.scan_meta_notes(path) == []

    def test_check_first_n_bounds_the_scan(self, docx):
        body = "".join(para(f"Body paragraph {i}.") for i in range(30))
        path = docx(body + para("Note: far down the document."))
        assert docx_tools.scan_meta_notes(path, check_first_n=5) == []
        assert len(docx_tools.scan_meta_notes(path, check_first_n=40)) == 1

    def test_caller_patterns_replace_the_defaults(self, docx):
        path = docx(para("Note: a default match.") + para("SENTINEL custom marker."))
        hits = docx_tools.scan_meta_notes(path, patterns=[r"^SENTINEL"])
        assert len(hits) == 1
        assert "SENTINEL" in hits[0]["text"]

    def test_an_empty_pattern_list_finds_nothing(self, docx):
        assert docx_tools.scan_meta_notes(docx(para("Note: x")), patterns=[]) == []

    def test_the_matched_pattern_is_reported(self, docx):
        hits = docx_tools.scan_meta_notes(docx(para("TODO: chase the reference.")))
        assert "TODO" in hits[0]["matched_pattern"]

    def test_a_note_inside_a_deletion_is_not_flagged(self, docx):
        body = f"<w:p><w:del {REVISION}><w:r><w:t>Note: struck out.</w:t></w:r></w:del></w:p>"
        assert docx_tools.scan_meta_notes(docx(body)) == []


class TestFileHandling:
    def test_a_missing_file_raises_rather_than_returning_empty(self, tmp_path: Path):
        with pytest.raises(FileNotFoundError):
            docx_tools.read_accepted_text(str(tmp_path / "nope.docx"))

    def test_a_non_zip_file_raises(self, tmp_path: Path):
        bad = tmp_path / "bad.docx"
        bad.write_bytes(b"this is not a zip archive")
        with pytest.raises(zipfile.BadZipFile):
            docx_tools.read_accepted_text(str(bad))

    def test_a_zip_with_no_document_part_is_named(self, tmp_path: Path):
        path = tmp_path / "other.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("notes.txt", "hello")
        with pytest.raises(ValueError, match="not a Word document"):
            docx_tools.read_accepted_text(str(path))

    def test_the_file_is_not_changed(self, docx):
        path = docx(para("Unchanged text."))
        before = Path(path).read_bytes()
        docx_tools.read_accepted_text(path)
        docx_tools.find_term(path, "text")
        docx_tools.get_docx_stats(path)
        docx_tools.scan_meta_notes(path)
        assert Path(path).read_bytes() == before


class TestUntrustedInput:
    def test_an_external_entity_is_not_resolved(self, docx, tmp_path: Path):
        secret = tmp_path / "private.txt"
        secret.write_text("PRIVATE_MARKER", encoding="utf-8")
        prolog = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<!DOCTYPE d [<!ENTITY x SYSTEM "{secret.as_uri()}">]>'
        )
        path = docx("<w:p><w:r><w:t>before &x; after</w:t></w:r></w:p>", prolog=prolog)
        assert "PRIVATE_MARKER" not in str(docx_tools.read_accepted_text(path, min_length=0))

    def test_entity_expansion_does_not_blow_up(self, docx):
        entities = '<!ENTITY a "aaaaaaaaaa">' + "".join(
            f'<!ENTITY {chr(98 + i)} "{("&" + chr(97 + i) + ";") * 10}">' for i in range(8)
        )
        prolog = f'<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE d [{entities}]>'
        path = docx("<w:p><w:r><w:t>&i;</w:t></w:r></w:p>", prolog=prolog)
        # Either outcome is safe: the parser refuses, or the entity stays unexpanded.
        try:
            rows = docx_tools.read_accepted_text(path, min_length=0)
        except etree.XMLSyntaxError:
            return
        assert len(str(rows)) < 1000

    def test_an_oversized_part_is_refused_before_it_is_read(self, docx, monkeypatch):
        path = docx(para("Some text."))
        monkeypatch.setattr(docx_tools, "MAX_PART_BYTES", 10)
        with pytest.raises(ValueError, match="byte limit"):
            docx_tools.read_accepted_text(path)
