"""Three-way document merge orchestration.

:func:`merge_documents` is the public entry point. :class:`BlockMerger` merges ordered
sequences of block-level items (``<w:p>`` and ``<w:tbl>``) and is reused recursively for
table-cell content; :class:`DocumentMerger` prepares an isolated result document and
writes the merged blocks into its body.
"""

from __future__ import annotations

import copy
import io
import re
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Sequence

from docx.api import Document
from docx.merge.canon import fingerprint
from docx.merge.diff3 import three_way_decisions
from docx.merge.refs import ReferenceImporter
from docx.merge.tablemerge import TableMerger
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.table import CT_Tbl

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.document import DocumentPart

DocumentInput = "str | IO[bytes] | DocumentObject"

_CONFLICT_START = "<<<<<<< %s"
_CONFLICT_MIDDLE = "======="
_CONFLICT_END = ">>>>>>> %s"


@dataclass
class MergeConflict:
    """An overlapping edit that could not be reconciled automatically.

    .. py:attribute:: location

        Position of the conflict as a path, e.g. ``"body[4]"`` or
        ``"body[2]/row[1]/cell[0]"``. Indices refer to the base document.

    .. py:attribute:: kind

        What kind of content conflicts, e.g. ``"paragraph"``, ``"table-row"``,
        ``"table-cell-properties"`` or ``"table-properties"``.

    .. py:attribute:: branches

        The two branch names contributing the incompatible alternatives, in the order
        *left* then *right*.

    .. py:attribute:: base_text / left_text / right_text

        Short plain-text summaries of the content from each side (empty string when a
        side deleted the content).
    """

    location: str
    kind: str
    branches: tuple[str, str]
    base_text: str = ""
    left_text: str = ""
    right_text: str = ""


@dataclass
class MergeResult:
    """Outcome of a three-way merge."""

    document: DocumentObject
    conflicts: tuple[MergeConflict, ...] = field(default_factory=tuple)

    @property
    def has_conflicts(self) -> bool:
        """|True| when one or more conflicts were reported."""
        return bool(self.conflicts)


def merge_documents(
    base: DocumentInput,
    left: DocumentInput,
    right: DocumentInput,
    *,
    left_name: str = "left",
    right_name: str = "right",
) -> MergeResult:
    """Merge two independently-edited copies of `base` into a new |Document|.

    `base`, `left`, and `right` may each be a filesystem path, a binary stream, or an
    already-open |Document|. Non-overlapping additions, deletions, and edits made on
    either branch are combined; overlapping, incompatible edits are returned as
    :class:`MergeConflict` objects on :class:`MergeResult` (and both alternatives are
    written into the result delimited by ``<<<<<<<`` / ``>>>>>>>`` marker paragraphs).

    The merge is deterministic -- merging the same inputs twice produces the same
    document -- and every image, hyperlink, numbering and style reference is remapped
    into the result package.
    """
    base_doc = _coerce_document(base)
    left_doc = _coerce_document(left)
    right_doc = _coerce_document(right)
    return DocumentMerger(
        base_doc, left_doc, right_doc, left_name=left_name, right_name=right_name
    ).merge()


def _coerce_document(value: DocumentInput) -> DocumentObject:
    if hasattr(value, "part") and hasattr(value, "save"):
        return value  # type: ignore[return-value]
    return Document(value)


def _marker_paragraph(text: str) -> BaseOxmlElement:
    """A plain ``<w:p><w:r><w:t>text</w:t></w:r></w:p>`` conflict marker."""
    paragraph = OxmlElement("w:p")
    run = OxmlElement("w:r")
    text_node = OxmlElement("w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = text
    run.append(text_node)
    paragraph.append(run)
    return paragraph


def content_summary(element: BaseOxmlElement | None, limit: int = 160) -> str:
    """Return a short single-line plain-text summary of block content."""
    if element is None:
        return ""
    text = " ".join(
        t.text for t in element.iter(qn("w:t")) if t.text  # type: ignore[union-attr]
    )
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


class BlockMerger:
    """Merge three parallel lists of block items (`<w:p>` / `<w:tbl>`)."""

    def __init__(
        self,
        target_part: DocumentPart,
        base_part: DocumentPart,
        left_part: DocumentPart,
        right_part: DocumentPart,
        left_name: str = "left",
        right_name: str = "right",
    ):
        self._target_part = target_part
        self._base_part = base_part
        self._left_part = left_part
        self._right_part = right_part
        self.left_name = left_name
        self.right_name = right_name
        self._left_importer = ReferenceImporter(left_part, target_part)
        self._right_importer = ReferenceImporter(right_part, target_part)
        self.conflicts: list[MergeConflict] = []

    @property
    def left_marker(self) -> str:
        return _CONFLICT_START % self.left_name

    @property
    def middle_marker(self) -> str:
        return _CONFLICT_MIDDLE

    @property
    def right_marker(self) -> str:
        return _CONFLICT_END % self.right_name

    def merge_blocks(
        self,
        base_items: Sequence[BaseOxmlElement],
        left_items: Sequence[BaseOxmlElement],
        right_items: Sequence[BaseOxmlElement],
        location: str = "body",
    ) -> list[BaseOxmlElement]:
        """Return the merged block elements and append any conflicts found."""
        # -- lxml may hand out fresh proxy objects for the same C node, so branch
        # membership uses element equality (node identity), never id() --
        left_set = set(left_items)
        right_set = set(right_items)

        def key(element: BaseOxmlElement) -> str:
            if element in left_set:
                return fingerprint(element, self._left_part)
            if element in right_set:
                return fingerprint(element, self._right_part)
            return fingerprint(element, self._base_part)

        decisions = three_way_decisions(base_items, left_items, right_items, key=key)

        result: list[BaseOxmlElement] = []
        for decision in decisions:
            kind = decision[0]
            if kind == "base":
                result.extend(copy.deepcopy(element) for element in decision[1])
            elif kind == "left":
                result.extend(self._left_importer.import_tree(e) for e in decision[1])
            elif kind == "right":
                result.extend(self._right_importer.import_tree(e) for e in decision[1])
            elif kind == "insert":
                result.extend(self._left_importer.import_tree(e) for e in decision[1])
                result.extend(self._right_importer.import_tree(e) for e in decision[2])
            else:
                _, b_lo, b_hi, b_elements, l_elements, r_elements = decision
                region_location = self._region_location(location, b_lo, b_hi)
                result.extend(
                    self._merge_conflict(
                        region_location, b_elements, l_elements, r_elements
                    )
                )
        return result

    def record_conflict(
        self,
        location: str,
        kind: str,
        base_element: BaseOxmlElement | None,
        left_element: BaseOxmlElement | None,
        right_element: BaseOxmlElement | None,
    ) -> MergeConflict:
        """Record and return a conflict at `location`."""
        conflict = MergeConflict(
            location=location,
            kind=kind,
            branches=(self.left_name, self.right_name),
            base_text=content_summary(base_element),
            left_text=content_summary(left_element),
            right_text=content_summary(right_element),
        )
        self.conflicts.append(conflict)
        return conflict

    def _merge_conflict(self, location, base_elements, left_elements, right_elements):
        # -- a single table changed on both sides can still merge at cell level --
        merged_table = self._try_table_merge(
            location, base_elements, left_elements, right_elements
        )
        if merged_table is not None:
            return [merged_table]

        kind = (
            "table"
            if any(isinstance(e, CT_Tbl) for e in (*base_elements, *left_elements, *right_elements))
            else "paragraph"
        )
        self.record_conflict(
            location,
            kind,
            base_elements[0] if base_elements else None,
            left_elements[0] if left_elements else None,
            right_elements[0] if right_elements else None,
        )
        output: list[BaseOxmlElement] = [_marker_paragraph(self.left_marker)]
        output.extend(self._left_importer.import_tree(e) for e in left_elements)
        output.append(_marker_paragraph(_CONFLICT_MIDDLE))
        output.extend(self._right_importer.import_tree(e) for e in right_elements)
        output.append(_marker_paragraph(self.right_marker))
        return output

    def _try_table_merge(
        self, location, base_elements, left_elements, right_elements
    ) -> BaseOxmlElement | None:
        if not (
            len(base_elements) == len(left_elements) == len(right_elements) == 1
            and isinstance(base_elements[0], CT_Tbl)
            and isinstance(left_elements[0], CT_Tbl)
            and isinstance(right_elements[0], CT_Tbl)
        ):
            return None
        merger = TableMerger(
            base_elements[0],
            left_elements[0],
            right_elements[0],
            self._base_part,
            self._left_importer,
            self._right_importer,
            self,
            location,
        )
        return merger.merge()

    @staticmethod
    def _region_location(location: str, lo: int, hi: int) -> str:
        return f"{location}[{lo}]" if hi - lo <= 1 else f"{location}[{lo}:{hi}]"


class DocumentMerger:
    """Prepare an isolated result document and merge the body blocks into it."""

    def __init__(
        self,
        base: DocumentObject,
        left: DocumentObject,
        right: DocumentObject,
        left_name: str = "left",
        right_name: str = "right",
    ):
        self._base = base
        self._left = left
        self._right = right
        self._left_name = left_name
        self._right_name = right_name

    def merge(self) -> MergeResult:
        result = self._isolated_copy_of_base()
        result_part = result.part

        block_merger = BlockMerger(
            target_part=result_part,
            base_part=self._base.part,
            left_part=self._left.part,
            right_part=self._right.part,
            left_name=self._left_name,
            right_name=self._right_name,
        )

        body = result._element.body
        body.clear_content()  # -- preserves the trailing <w:sectPr> --

        merged_blocks = block_merger.merge_blocks(
            self._base._element.body.inner_content_elements,
            self._left._element.body.inner_content_elements,
            self._right._element.body.inner_content_elements,
        )

        sectPr = body.find(qn("w:sectPr"))
        for element in merged_blocks:
            if sectPr is not None:
                sectPr.addprevious(element)
            else:
                body.append(element)

        return MergeResult(result, tuple(block_merger.conflicts))

    def _isolated_copy_of_base(self) -> DocumentObject:
        """Re-open a serialized copy of the base document so the merge result never
        shares mutable package state with any of the input documents."""
        stream = io.BytesIO()
        self._base.save(stream)
        stream.seek(0)
        return Document(stream)
