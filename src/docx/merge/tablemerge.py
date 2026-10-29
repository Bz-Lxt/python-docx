"""Three-way merge of table content, including vertically/horizontally merged cells.

A whole-table conflict is avoided whenever both branches edited the *same* existing
table. Instead the table is merged structurally:

* ``<w:tblPr>`` (table style and other table properties) and ``<w:tblGrid>`` follow the
  usual rule -- unchanged/one-side/both-identical merge cleanly, differing edits on
  both sides become a conflict located at the table;
* rows are aligned with the same diff3 machinery used for body blocks, so independent
  row insertions and deletions merge cleanly;
* a row changed on both sides is merged cell-by-cell on the layout grid. Vertically
  merged ("vMerge=continue") cells resolve to their anchor cell, and horizontally
  merged (``gridSpan``) cells are matched at their starting column;
* cell properties (``<w:tcPr>`` -- shading, width, merge status, ...) and cell content
  (paragraphs and nested tables) are merged independently, so an edit to cell shading
  on one branch and an edit to the cell text on the other combine without conflict.

Only genuinely incompatible structure (column-count changes, differing spans at the
same position, delete-vs-edit of a row) is escalated as a conflict.
"""

from __future__ import annotations

import copy
from typing import TYPE_CHECKING

from docx.merge.canon import fingerprint
from docx.merge.diff3 import three_way_decisions
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement

if TYPE_CHECKING:
    from docx.merge.merger import BlockMerger
    from docx.merge.refs import ReferenceImporter
    from docx.oxml.table import CT_Row, CT_Tbl, CT_Tc


def grid_rows(tbl: CT_Tbl) -> list[dict[int, tuple[CT_Tc, int, str | None]]]:
    """Return per-row mappings of grid start-column -> (anchor tc, gridSpan, vMerge).

    ``vMerge=continue`` cells are resolved to the anchor ``<w:tc>`` of the vertical
    merge in an earlier row (the third tuple item still carries this row's own
    ``"continue"`` status). Rows returned are aligned with ``tbl.tr_lst``.
    """
    rows: list[dict[int, tuple[CT_Tc, int, str | None]]] = []
    vertical_anchors: dict[int, CT_Tc] = {}
    for tr in tbl.tr_lst:
        col = tr.grid_before
        row: dict[int, tuple[CT_Tc, int, str | None]] = {}
        for tc in tr.tc_lst:
            start = col
            span = tc.grid_span
            v_merge = tc.vMerge
            if v_merge == "continue":
                anchor_tc = vertical_anchors[start]
            else:
                anchor_tc = tc
                if v_merge == "restart":
                    vertical_anchors[start] = tc
                else:
                    vertical_anchors.pop(start, None)
            row[start] = (anchor_tc, span, v_merge)
            col += span
        rows.append(row)
    return rows


class TableMerger:
    """Merge one base table with its two independently-edited versions."""

    def __init__(
        self,
        base_tbl: CT_Tbl,
        left_tbl: CT_Tbl,
        right_tbl: CT_Tbl,
        base_part,
        left_importer: ReferenceImporter,
        right_importer: ReferenceImporter,
        block_merger: BlockMerger,
        location: str,
    ):
        self._base_tbl = base_tbl
        self._left_tbl = left_tbl
        self._right_tbl = right_tbl
        self._base_part = base_part
        self._left_part = left_importer._source  # pyright: ignore
        self._right_part = right_importer._source  # pyright: ignore
        self._left_importer = left_importer
        self._right_importer = right_importer
        self._block_merger = block_merger
        self._location = location

    def merge(self) -> CT_Tbl | None:
        """Return a new merged ``<w:tbl>`` element, or |None| when the table structures
        are incompatible and the caller should report a block-level conflict."""
        col_counts = {
            self._base_tbl.col_count,
            self._left_tbl.col_count,
            self._right_tbl.col_count,
        }
        if len(col_counts) != 1:
            return None

        result_tbl = copy.deepcopy(self._base_tbl)

        # -- check the grid first: an incompatible grid aborts the whole structural
        # merge (the caller reports a single block-level table conflict) --
        if not self._merge_table_grid(result_tbl):
            return None
        self._merge_table_properties(result_tbl)

        merged_rows = self._merge_rows()
        if merged_rows is None:
            return None

        for tr in result_tbl.tr_lst:
            result_tbl.remove(tr)
        for tr in merged_rows:
            result_tbl.append(tr)
        return result_tbl

    # -- table-level properties --------------------------------------------------

    def _merge_table_properties(self, result_tbl: CT_Tbl) -> None:
        chosen, changed, source = self._choose(
            self._base_tbl.tblPr, self._left_tbl.tblPr, self._right_tbl.tblPr
        )
        if not changed:
            return
        if source is None:
            # -- both branches changed tblPr differently --
            self._block_merger.record_conflict(
                self._location,
                "table-properties",
                self._base_tbl.tblPr,
                self._left_tbl.tblPr,
                self._right_tbl.tblPr,
            )
            chosen = self._left_tbl.tblPr
            source = "left"
        new_tblPr = (
            self._importer_for(source).import_tree(chosen) if source else copy.deepcopy(chosen)
        )
        old = result_tbl.tblPr
        old.getparent().replace(old, new_tblPr)

    def _merge_table_grid(self, result_tbl: CT_Tbl) -> bool:
        base_grid = self._base_tbl.tblGrid
        chosen, changed, source = self._choose(
            base_grid, self._left_tbl.tblGrid, self._right_tbl.tblGrid
        )
        if not changed:
            return True
        if source is None:
            # -- the branches reworked the grid differently; let the caller report
            # the single block-level table conflict rather than duplicating it --
            return False
        if source == "left":
            new_grid = self._left_importer.import_tree(chosen)
        elif source == "right":
            new_grid = self._right_importer.import_tree(chosen)
        else:
            new_grid = copy.deepcopy(chosen)
        old = result_tbl.tblGrid
        old.getparent().replace(old, new_grid)
        return True

    # -- rows --------------------------------------------------------------------

    def _merge_rows(self) -> list[CT_Row] | None:
        base_rows = self._base_tbl.tr_lst
        left_rows = self._left_tbl.tr_lst
        right_rows = self._right_tbl.tr_lst

        left_grid = grid_rows(self._left_tbl)
        right_grid = grid_rows(self._right_tbl)

        # -- identify which tree a row belongs to by node identity --
        left_row_set = set(left_rows)
        right_row_set = set(right_rows)
        decisions = three_way_decisions(
            base_rows,
            left_rows,
            right_rows,
            key=lambda tr: fingerprint(
                tr,
                self._left_part
                if tr in left_row_set
                else self._right_part
                if tr in right_row_set
                else self._base_part,
            ),
        )

        left_grid_by_row = dict(zip(left_rows, left_grid))
        right_grid_by_row = dict(zip(right_rows, right_grid))

        rows: list[CT_Row] = []
        for decision in decisions:
            kind = decision[0]
            if kind == "base":
                for tr in decision[1]:
                    rows.append(copy.deepcopy(tr))
            elif kind == "left":
                for tr in decision[1]:
                    rows.append(self._left_importer.import_tree(tr))
            elif kind == "right":
                for tr in decision[1]:
                    rows.append(self._right_importer.import_tree(tr))
            elif kind == "insert":
                for tr in decision[1]:
                    rows.append(self._left_importer.import_tree(tr))
                for tr in decision[2]:
                    rows.append(self._right_importer.import_tree(tr))
            else:
                _, b_lo, _b_hi, b_rows, l_rows, r_rows = decision
                merged_row = None
                if len(b_rows) == len(l_rows) == len(r_rows) == 1:
                    merged_row = self._merge_row(
                        b_rows[0],
                        l_rows[0],
                        r_rows[0],
                        left_grid_by_row[l_rows[0]],
                        right_grid_by_row[r_rows[0]],
                        b_lo,
                    )
                if merged_row is not None:
                    rows.append(merged_row)
                else:
                    self._block_merger.record_conflict(
                        "%s/row[%d]" % (self._location, b_lo),
                        "table-row",
                        b_rows[0] if b_rows else None,
                        l_rows[0] if l_rows else None,
                        r_rows[0] if r_rows else None,
                    )
                    self._emit_conflicting_rows(rows, l_rows, r_rows)
        return rows

    def _emit_conflicting_rows(self, rows, l_rows, r_rows) -> None:
        """Append both alternatives with a complete marker block.

        Markers are paragraphs that must live inside cells, so the start/middle/end
        markers are placed in the first cell of the boundary rows. When one branch
        deleted the rows, the empty side is still delimited so the marker block stays
        well-formed (``<<<<<<<`` / ``=======`` / ``>>>>>>>`` all present).
        """
        left_rows = [self._left_importer.import_tree(tr) for tr in l_rows]
        right_rows = [self._right_importer.import_tree(tr) for tr in r_rows]
        start = _marker_paragraph(self._block_merger.left_marker)
        middle = _marker_paragraph(self._block_merger.middle_marker)
        end = _marker_paragraph(self._block_merger.right_marker)

        if left_rows and right_rows:
            _prefix_first_cell(left_rows[0], [start])
            _suffix_first_cell(left_rows[-1], [middle])
            _suffix_first_cell(right_rows[-1], [end])
        elif right_rows:
            # -- left deleted the rows: an empty left section --
            _prefix_first_cell(right_rows[0], [start, middle])
            _suffix_first_cell(right_rows[-1], [end])
        elif left_rows:
            # -- right deleted the rows: an empty right section --
            _prefix_first_cell(left_rows[0], [start])
            _suffix_first_cell(left_rows[-1], [middle, end])

        rows.extend(left_rows)
        rows.extend(right_rows)

    def _merge_row(
        self,
        base_tr: CT_Row,
        left_tr: CT_Row,
        right_tr: CT_Row,
        left_row_grid: dict[int, tuple[CT_Tc, int, str | None]],
        right_row_grid: dict[int, tuple[CT_Tc, int, str | None]],
        row_index: int,
    ) -> CT_Row | None:
        base_grid = grid_rows_for_tr(base_tr)
        if set(base_grid) != set(left_row_grid) or set(base_grid) != set(right_row_grid):
            return None
        for start, (_tc, span, v_merge) in base_grid.items():
            # -- column positions, spans and vertical-merge status must all line
            # up; any branch reworked the layout -> escalate to a row conflict --
            if left_row_grid[start][1:] != (span, v_merge):
                return None
            if right_row_grid[start][1:] != (span, v_merge):
                return None

        new_tr = copy.deepcopy(base_tr)
        for tc in new_tr.tc_lst:
            new_tr.remove(tc)

        merged_trPr = self._merge_trPr(base_tr, left_tr, right_tr, row_index)
        if merged_trPr is not None:
            tblPrEx = new_tr.find(qn("w:tblPrEx"))
            if tblPrEx is not None:
                tblPrEx.addnext(merged_trPr)
            else:
                new_tr.insert(0, merged_trPr)

        col = base_tr.grid_before
        for base_tc in base_tr.tc_lst:
            start = col
            col += base_tc.grid_span
            if base_tc.vMerge == "continue":
                # -- continuation cells carry no content; keep the base placeholder --
                new_tr.append(copy.deepcopy(base_tc))
                continue
            left_tc = left_row_grid[start][0]
            right_tc = right_row_grid[start][0]
            new_tc = self._merge_cell(base_tc, left_tc, right_tc, row_index, start)
            new_tr.append(new_tc)
        return new_tr

    def _merge_trPr(self, base_tr, left_tr, right_tr, row_index):
        base_trPr = base_tr.find(qn("w:trPr"))
        left_trPr = left_tr.find(qn("w:trPr"))
        right_trPr = right_tr.find(qn("w:trPr"))
        chosen, changed, source = self._choose(base_trPr, left_trPr, right_trPr)
        if not changed:
            return copy.deepcopy(base_trPr) if base_trPr is not None else None
        if source is None:
            self._block_merger.record_conflict(
                "%s/row[%d]/properties" % (self._location, row_index),
                "table-row-properties",
                base_trPr,
                left_trPr,
                right_trPr,
            )
            chosen, source = left_trPr, "left"
        if source == "left":
            return self._left_importer.import_tree(chosen)
        if source == "right":
            return self._right_importer.import_tree(chosen)
        return copy.deepcopy(chosen)

    # -- cells -------------------------------------------------------------------

    def _merge_cell(
        self, base_tc: CT_Tc, left_tc: CT_Tc, right_tc: CT_Tc, row_index: int, col_index: int
    ) -> CT_Tc:
        new_tc = copy.deepcopy(base_tc)
        cell_location = "%s/row[%d]/cell[%d]" % (self._location, row_index, col_index)

        # -- cell properties --
        base_tcPr = base_tc.find(qn("w:tcPr"))
        left_tcPr = left_tc.find(qn("w:tcPr"))
        right_tcPr = right_tc.find(qn("w:tcPr"))
        chosen, changed, source = self._choose(base_tcPr, left_tcPr, right_tcPr)
        if changed:
            if source is None:
                self._block_merger.record_conflict(
                    cell_location + "/properties",
                    "table-cell-properties",
                    base_tcPr,
                    left_tcPr,
                    right_tcPr,
                )
                chosen, source = left_tcPr, "left"
            new_tcPr = (
                self._importer_for(source).import_tree(chosen)
                if source
                else copy.deepcopy(chosen)
            )
            existing = new_tc.find(qn("w:tcPr"))
            if existing is not None:
                new_tc.replace(existing, new_tcPr)
            else:
                new_tc.insert(0, new_tcPr)

        # -- cell content (paragraphs and nested tables) --
        base_items = base_tc.xpath("./w:p | ./w:tbl")
        left_items = left_tc.xpath("./w:p | ./w:tbl")
        right_items = right_tc.xpath("./w:p | ./w:tbl")
        content = self._block_merger.merge_blocks(
            base_items, left_items, right_items, cell_location
        )

        for child in list(new_tc):
            if child.tag != qn("w:tcPr"):
                new_tc.remove(child)
        for element in content:
            new_tc.append(element)
        # -- a table cell must contain at least one block and end with a paragraph --
        children = [c for c in new_tc if c.tag != qn("w:tcPr")]
        if not children or children[-1].tag != qn("w:p"):
            new_tc.append(OxmlElement("w:p"))
        return new_tc

    # -- helpers -----------------------------------------------------------------

    def _importer_for(self, source: str) -> ReferenceImporter:
        return self._left_importer if source == "left" else self._right_importer

    def _choose(self, base, left, right):
        """Return ``(chosen_element, changed, source)``.

        `source` is ``'left'`` / ``'right'`` / ``'base'`` for a clean choice, |None|
        when both branches changed the value differently (a conflict).
        """
        base_xml = self._xml(base, self._base_part)
        left_xml = self._xml(left, self._left_part)
        right_xml = self._xml(right, self._right_part)
        left_changed = left_xml != base_xml
        right_changed = right_xml != base_xml
        if not left_changed and not right_changed:
            return base, False, "base"
        if left_changed and not right_changed:
            return left, True, "left"
        if right_changed and not left_changed:
            return right, True, "right"
        if left_xml == right_xml:
            return left, True, "left"
        return base, True, None

    @staticmethod
    def _xml(element, part) -> str:
        return fingerprint(element, part) if element is not None else ""


def grid_rows_for_tr(tr: CT_Row) -> dict[int, tuple[CT_Tc, int, str | None]]:
    """Mapping of grid start-column -> (tc, span, vMerge) for a single row.

    Unlike :func:`grid_rows`, continuation cells map to themselves because the anchor
    cannot be resolved without the whole table; callers only use the anchor mapping
    from a full-table grid and skip continuations explicitly.
    """
    col = tr.grid_before
    row: dict[int, tuple[CT_Tc, int, str | None]] = {}
    for tc in tr.tc_lst:
        row[col] = (tc, tc.grid_span, tc.vMerge)
        col += tc.grid_span
    return row


def _marker_paragraph(text: str):
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def _prefix_first_cell(tr: CT_Row, paragraphs: list) -> None:
    """Insert `paragraphs` (in order) at the top of the row's first cell."""
    if not tr.tc_lst:
        return
    tc = tr.tc_lst[0]
    tcPr = tc.find(qn("w:tcPr"))
    anchor = tcPr
    for paragraph in paragraphs:
        if anchor is not None:
            anchor.addnext(paragraph)
        else:
            tc.insert(0, paragraph)
        anchor = paragraph


def _suffix_first_cell(tr: CT_Row, paragraphs: list) -> None:
    """Append `paragraphs` (in order) at the bottom of the row's first cell."""
    if not tr.tc_lst:
        return
    tc = tr.tc_lst[0]
    for paragraph in paragraphs:
        tc.append(paragraph)
