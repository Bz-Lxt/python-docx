"""Canonical fingerprints for aligning identical content across three documents.

A fingerprint is a deterministic, namespace-aware serialization of an XML element
sub-tree. Two elements that "mean" the same thing produce the same fingerprint even
when they live in different packages and therefore differ in incidental details:

* revision (rsid) attributes, which Word rewrites on every save, are ignored;
* relationship references (``r:id`` / ``r:embed``) are normalized to a token derived
  from the *target* of the relationship -- an external URL or a hash of the target
  part blob -- so the same picture/link added independently on two branches matches;
* numbering references are normalized against the numbering definition they point at
  rather than the (package-local) numId;
* style references are normalized against the referenced style definition rather than
  the (package-local) styleId.
"""

from __future__ import annotations

import hashlib
from typing import TYPE_CHECKING

from docx.oxml.ns import nsmap

if TYPE_CHECKING:
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.story import StoryPart

W_URI = nsmap["w"]
R_URI = nsmap["r"]

# -- w:rsid* attributes are session-specific and must not influence identity --
_RSID_LOCALS = frozenset(
    (
        "rsidR",
        "rsidRDefault",
        "rsidP",
        "rsidRPr",
        "rsidTr",
        "rsidTbl",
        "rsidDel",
        "rsidIns",
        "rsidMoveFrom",
        "rsidMoveTo",
        "rsidPaste",
        "rsidSplits",
    )
)

_REF_ATTR_LOCALS = frozenset(("id", "embed", "link"))
_STYLE_REF_TAGS = frozenset(("pStyle", "rStyle", "tblStyle"))


def _local_name(tag: str) -> str:
    return tag[tag.rfind("}") + 1 :] if "}" in tag else tag


def fingerprint(element: BaseOxmlElement, part: StoryPart | None = None) -> str:
    """Return the canonical fingerprint string for `element`.

    `part` is the story part containing the element; it is used to resolve reference
    attributes. When |None|, reference values are compared literally.
    """
    cache = _part_cache(part)
    # -- lxml hands out fresh Python proxies for the same C node, so the element
    # itself (identity equality/hash on the underlying node), not id(), must key
    # the cache; a recycled proxy id would otherwise return a stale fingerprint --
    cached = cache.get(element)
    if cached is None:
        cached = _emit(element, part)
        cache[element] = cached
    return cached


def canonical_xml(element: BaseOxmlElement) -> str:
    """Canonical serialization without any reference normalization."""
    return _emit(element, None)


def _part_cache(part: StoryPart | None) -> dict:
    if part is None:
        return {}
    cache = getattr(part, "_merge_fingerprint_cache", None)
    if cache is None:
        cache = {}
        # -- StoryPart is not ours to grow __slots__ on; stash on the rels object --
        setattr(part, "_merge_fingerprint_cache", cache)
    return cache


def _emit(element: BaseOxmlElement, part: StoryPart | None, out: list[str] | None = None) -> str:
    if out is None:
        out = []
    tag = element.tag
    if not isinstance(tag, str):
        # -- comments / PIs --
        return ""
    out.append("<" + tag)
    tag_local = _local_name(tag)
    for name, value in sorted(element.attrib.items()):
        if name.startswith("{%s}" % W_URI) and _local_name(name) in _RSID_LOCALS:
            continue
        if name.startswith("{%s}" % R_URI) and _local_name(name) in _REF_ATTR_LOCALS:
            value = _relationship_token(part, value)
        elif name == "{%s}val" % W_URI and tag_local in _STYLE_REF_TAGS:
            value = _style_token(part, tag_local, value)
        elif name == "{%s}val" % W_URI and tag_local == "numId":
            value = _numbering_token(part, value)
        out.append(" %s=%r" % (name, value))
    out.append(">")
    if element.text:
        out.append(element.text)
    for child in element:
        _emit(child, part, out)
        if child.tail:
            out.append(child.tail)
    out.append("</%s>" % tag)
    return "".join(out)


def _relationship_token(part: StoryPart | None, rId: str) -> str:
    if part is None:
        return "rId:" + rId
    cache = getattr(part, "_merge_rel_token_cache", None)
    if cache is None:
        cache = {}
        setattr(part, "_merge_rel_token_cache", cache)
    if rId in cache:
        return cache[rId]
    rel = part.rels.get(rId)
    if rel is None:
        token = "rId:" + rId
    elif rel.is_external:
        token = "ext:%s:%s" % (rel.reltype.rsplit("/", 1)[-1], rel.target_ref)
    else:
        digest = hashlib.sha1(rel.target_part.blob).hexdigest()[:16]
        token = "part:%s:%s" % (rel.reltype.rsplit("/", 1)[-1], digest)
    cache[rId] = token
    return token


def _style_token(part: StoryPart | None, tag_local: str, style_id: str) -> str:
    if part is None:
        return f"{tag_local}:{style_id}"
    try:
        styles_root = part._document_part._styles_part.element  # pyright: ignore
    except KeyError:
        return f"{tag_local}:{style_id}"
    style = styles_root.get_by_id(style_id)
    if style is None:
        return f"{tag_local}:{style_id}"
    digest = hashlib.sha1(canonical_xml(style).encode("utf-8")).hexdigest()[:12]
    return f"style:{digest}"


def _numbering_token(part: StoryPart | None, num_id_str: str) -> str:
    if num_id_str == "0" or part is None:
        return "numId:" + num_id_str
    try:
        numbering_root = part._document_part.part_related_by(  # pyright: ignore
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
        ).element
    except KeyError:
        return "numId:" + num_id_str
    nums = numbering_root.xpath('./w:num[@w:numId="%s"]' % num_id_str)
    if not nums:
        return "numId:" + num_id_str
    num = nums[0]
    abstract_ids = num.xpath("./w:abstractNumId/@w:val")
    if not abstract_ids:
        return "numId:" + num_id_str
    abstracts = numbering_root.xpath(
        './w:abstractNum[@w:abstractNumId="%s"]' % abstract_ids[0]
    )
    payload = canonical_xml(abstracts[0]) if abstracts else ""
    payload += "".join(canonical_xml(e) for e in num.xpath("./w:lvlOverride"))
    digest = hashlib.sha1(payload.encode("utf-8")).hexdigest()[:12]
    return "numDef:" + digest
