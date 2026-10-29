"""Three-way merge of WordprocessingML documents.

Given a *base* document and two documents (*left* and *right*) independently derived
from it, :func:`merge_documents` produces a new document combining the non-overlapping
additions, deletions, and modifications made on each branch. Body paragraphs, tables
(including those with merged cells and applied styles), numbering definitions, images,
and hyperlinks are all merged.

Edits that overlap and cannot be reconciled are reported as :class:`MergeConflict`
objects carrying the conflict location together with the branch each alternative
originated from; both alternatives are also written into the merged document delimited
by conflict-marker paragraphs.

Repeatedly merging the same three documents yields an identical result, and the merged
document can be saved and re-opened with every relationship reference (images,
hyperlinks, numbering and style ids) correctly re-mapped rather than aliased between
branches.
"""

from __future__ import annotations

from docx.merge.merger import DocumentMerger, MergeConflict, MergeResult, merge_documents

__all__ = ["DocumentMerger", "MergeConflict", "MergeResult", "merge_documents"]
