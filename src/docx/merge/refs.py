"""Cross-package reference re-mapping when importing XML content between documents.

Content copied from a branch document into the merged document almost always carries
package-local references:

* ``a:blip/@r:embed`` (and legacy ``v:imagedata/@r:id``) point at image parts;
* ``w:hyperlink/@r:id`` points at an external hyperlink relationship;
* ``w:numPr/w:numId/@w:val`` points at ``<w:num>`` / ``<w:abstractNum>`` definitions in
  the numbering part;
* ``w:pStyle`` / ``w:rStyle`` / ``w:tblStyle`` point at style definitions in the styles
  part.

Re-using a raw rId or numId/styleId from a different package would "cross the wires":
two branches that each added a *different* picture as ``rId8`` would end up sharing a
single relationship. :class:`ReferenceImporter` deep-copies an XML sub-tree and rewires
every reference it contains:

* image parts are imported by content (identical blobs share a part, distinct blobs get
  distinct parts and partnames);
* external hyperlinks are imported by target URL;
* numbering definitions are copied with freshly-allocated abstractNumId/numId values,
  re-using an existing abstract definition when its content is identical;
* style definitions missing from the target are copied; a styleId that already exists
  with *different* content is duplicated under a fresh id and its references are
  rewritten to match (linked/based-on/next chains are followed transitively).

A separate importer is used per source document, but both importers write into the
same target parts, so references can never alias between branches.
"""

from __future__ import annotations

import copy
import io
from typing import TYPE_CHECKING

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml.ns import qn
from docx.parts.image import ImagePart

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.document import DocumentPart

# -- the "v" (VML) prefix is not registered in docx.oxml.ns.nsmap, use Clark notation --
_VML_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"
_R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

_STYLE_REF_TAGS = ("w:pStyle", "w:rStyle", "w:tblStyle")
_STYLE_DEP_TAGS = ("w:basedOn", "w:next", "w:link")


class ReferenceImporter:
    """Deep-copy XML sub-trees from `source_part` into `target_part`, re-mapping all
    package-local references to equivalents valid in the target package."""

    def __init__(self, source_part: DocumentPart, target_part: DocumentPart):
        self._source = source_part
        self._target = target_part
        self._image_rids: dict[str, str] = {}
        self._hyperlink_rids: dict[str, str] = {}
        self._num_ids: dict[str, str] = {}
        self._style_ids: dict[str, str] = {}
        self._style_alias_seq = 0

    # -- public entry point -----------------------------------------------------

    def import_tree(self, element: BaseOxmlElement) -> BaseOxmlElement:
        """Return a deep copy of `element` with all references remapped into the
        target package."""
        new_element = copy.deepcopy(element)
        self._remap_styles(new_element)
        self._remap_numbers(new_element)
        self._remap_relationships(new_element)
        return new_element

    # -- styles -----------------------------------------------------------------

    @property
    def _target_styles_root(self):
        return self._target._styles_part.element  # pyright: ignore[reportPrivateUsage]

    @property
    def _source_styles_root(self):
        try:
            return self._source.part_related_by(RT.STYLES).element
        except KeyError:
            return None

    def _remap_styles(self, tree: BaseOxmlElement) -> None:
        for tag in _STYLE_REF_TAGS:
            for ref in tree.xpath(".//" + tag):
                style_id = ref.get(qn("w:val"))
                if style_id:
                    ref.set(qn("w:val"), self._import_style(style_id))

    def _import_style(self, style_id: str) -> str:
        cached = self._style_ids.get(style_id)
        if cached is not None:
            return cached
        source_root = self._source_styles_root
        source_style = source_root.get_by_id(style_id) if source_root is not None else None
        if source_style is None:
            # -- unresolved reference, leave it untouched --
            self._style_ids[style_id] = style_id
            return style_id

        target_root = self._target_styles_root
        existing = target_root.get_by_id(style_id)
        if existing is not None and self._style_matches(existing, source_style):
            self._style_ids[style_id] = style_id
            return style_id

        new_id = style_id if existing is None else self._unique_style_id(style_id)
        # -- reserve the mapping up-front so linked style pairs cannot recurse --
        self._style_ids[style_id] = new_id

        clone = copy.deepcopy(source_style)
        clone.set(qn("w:styleId"), new_id)
        for tag in _STYLE_DEP_TAGS:
            for dep in clone.xpath("./" + tag):
                dep_val = dep.get(qn("w:val"))
                if dep_val:
                    dep.set(qn("w:val"), self._import_style(dep_val))
        # -- paragraph styles may themselves carry numbering references --
        self._remap_numbers(clone)
        target_root.append(clone)
        return new_id

    @staticmethod
    def _style_matches(existing, candidate) -> bool:
        from docx.merge.canon import canonical_xml

        return canonical_xml(existing) == canonical_xml(candidate)

    def _unique_style_id(self, base_id: str) -> str:
        target_root = self._target_styles_root
        while True:
            self._style_alias_seq += 1
            candidate = "%sMerge%d" % (base_id, self._style_alias_seq)
            if target_root.get_by_id(candidate) is None:
                return candidate

    # -- numbering --------------------------------------------------------------

    @property
    def _target_numbering_root(self):
        # -- accessing .numbering_part creates an (empty) one when absent --
        return self._target.numbering_part.element

    @property
    def _source_numbering_root(self):
        try:
            return self._source.part_related_by(RT.NUMBERING).element
        except KeyError:
            return None

    def _remap_numbers(self, tree: BaseOxmlElement) -> None:
        for num_id_el in tree.xpath(".//w:numPr/w:numId[@w:val]"):
            value = num_id_el.get(qn("w:val"))
            if value and value != "0":
                num_id_el.set(qn("w:val"), self._import_num(value))

    def _import_num(self, num_id_str: str) -> str:
        cached = self._num_ids.get(num_id_str)
        if cached is not None:
            return cached
        source_root = self._source_numbering_root
        source_nums = (
            source_root.xpath('./w:num[@w:numId="%s"]' % num_id_str)
            if source_root is not None
            else []
        )
        if not source_nums:
            self._num_ids[num_id_str] = num_id_str
            return num_id_str
        source_num = source_nums[0]

        abstract_ids = source_num.xpath("./w:abstractNumId/@w:val")
        source_abstract_id = abstract_ids[0]
        source_abstracts = source_root.xpath(
            './w:abstractNum[@w:abstractNumId="%s"]' % source_abstract_id
        )
        source_abstract = source_abstracts[0]

        target_root = self._target_numbering_root
        target_abstract_id = self._matching_abstract_num_id(target_root, source_abstract)
        if target_abstract_id is None:
            target_abstract_id = self._next_abstract_num_id(target_root)
            abstract_clone = copy.deepcopy(source_abstract)
            abstract_clone.set(qn("w:abstractNumId"), str(target_abstract_id))
            first_num = target_root.find(qn("w:num"))
            if first_num is not None:
                first_num.addprevious(abstract_clone)
            else:
                target_root.append(abstract_clone)

        target_num_id = self._next_num_id(target_root)
        num_clone = copy.deepcopy(source_num)
        num_clone.set(qn("w:numId"), str(target_num_id))
        num_clone.xpath("./w:abstractNumId")[0].set(qn("w:val"), str(target_abstract_id))
        target_root.append(num_clone)

        new_value = str(target_num_id)
        self._num_ids[num_id_str] = new_value
        return new_value

    @staticmethod
    def _matching_abstract_num_id(target_root, source_abstract):
        from docx.merge.canon import canonical_xml

        source_xml = canonical_xml(source_abstract)
        for abstract in target_root.xpath("./w:abstractNum"):
            if canonical_xml(abstract) == source_xml:
                return int(abstract.get(qn("w:abstractNumId")))
        return None

    @staticmethod
    def _next_abstract_num_id(root) -> int:
        values = [int(v) for v in root.xpath("./w:abstractNum/@w:abstractNumId")]
        return (max(values) if values else -1) + 1

    @staticmethod
    def _next_num_id(root) -> int:
        values = [int(v) for v in root.xpath("./w:num/@w:numId")]
        return (max(values) if values else 0) + 1

    # -- images and hyperlinks --------------------------------------------------

    def _remap_relationships(self, tree: BaseOxmlElement) -> None:
        for blip in tree.xpath(".//a:blip[@r:embed]"):
            rId = blip.get(qn("r:embed"))
            blip.set(qn("r:embed"), self._import_image_rid(rId))
        for blip in tree.xpath(".//a:blip[@r:link]"):
            rId = blip.get(qn("r:link"))
            blip.set(qn("r:link"), self._import_hyperlink_rid(rId))
        for image_data in tree.iter(_VML_IMAGEDATA):
            rId = image_data.get(_R_ID)
            if rId:
                image_data.set(_R_ID, self._import_image_rid(rId))
        for hyperlink in tree.xpath(".//w:hyperlink[@r:id]"):
            rId = hyperlink.get(qn("r:id"))
            hyperlink.set(qn("r:id"), self._import_hyperlink_rid(rId))

    def _import_image_rid(self, rId: str | None) -> str:
        if rId is None:
            return ""
        cached = self._image_rids.get(rId)
        if cached is not None:
            return cached
        rel = self._source.rels.get(rId)
        if rel is None or rel.is_external:
            # -- linked (external) image, or a dangling reference --
            new_rid = (
                self._import_hyperlink_rid(rId)
                if rel is not None and rel.is_external
                else rId
            )
            self._image_rids[rId] = new_rid
            return new_rid

        source_image_part = rel.target_part
        image_part = self._target.package.get_or_add_image_part(
            io.BytesIO(source_image_part.blob)
        )
        if image_part.sha1 != source_image_part.sha1:
            # -- extremely defensive: sniffer produced different bytes, rebuild --
            image_part = self._copy_image_part(source_image_part)
        new_rid = self._target.relate_to(image_part, RT.IMAGE)
        self._image_rids[rId] = new_rid
        return new_rid

    def _copy_image_part(self, source_image_part) -> ImagePart:
        """Copy an image part that could not be re-decoded from its blob (e.g. WMF),
        preserving its content type and extension under a fresh partname."""
        package = self._target.package
        used_indexes = {part.partname.idx for part in package.image_parts}
        index = 1
        while index in used_indexes:
            index += 1
        partname = PackURI("/word/media/image%d%s" % (index, source_image_part.partname.ext))
        image_part = ImagePart(
            partname, source_image_part.content_type, source_image_part.blob
        )
        package.image_parts.append(image_part)
        return image_part

    def _import_hyperlink_rid(self, rId: str | None) -> str:
        if rId is None:
            return ""
        cached = self._hyperlink_rids.get(rId)
        if cached is not None:
            return cached
        rel = self._source.rels.get(rId)
        if rel is None:
            self._hyperlink_rids[rId] = rId
            return rId
        if rel.is_external:
            new_rid = self._target.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            # -- best effort for internal targets (rare for hyperlinks) --
            new_rid = self._target.relate_to(rel.target_part, rel.reltype)
        self._hyperlink_rids[rId] = new_rid
        return new_rid
