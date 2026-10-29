"""Three-way sequence merge (diff3) over lists of hashable items.

The algorithm mirrors a line-oriented source-code merge:

1. Align *base* against *left* and *base* against *right* with
   :class:`difflib.SequenceMatcher`, producing change *hunks* for each branch -- a
   hunk replaces or deletes a base interval, or inserts items at a boundary.
2. Hunks from the two branches whose base intervals *overlap* cannot be reconciled
   automatically and become a conflict; hunks with disjoint intervals are applied
   independently. Editing item *i* on one branch and item *j* on the other merges
   cleanly even when no unchanged item sits between them.
3. The same change on both branches is taken once. Different insertions at the same
   boundary are both kept (left first, then right). An insertion strictly inside the
   span of the other branch's replacement joins that hunk into a conflict, since the
   insertion point itself was rewritten.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import Callable, List, Sequence, TypeVar

T = TypeVar("T")

# -- decision tuples --
# ("base",    [items])                      unchanged on both branches
# ("left",    [items])                      only left changed
# ("right",   [items])                      only right changed
# ("insert",  [left_items], [right_items])  both branches inserted here
# ("conflict", b_lo, b_hi, [b], [l], [r])   incompatible overlapping edits
Decision = tuple


@dataclass
class _Hunk:
    """One branch-side change: `items` replace ``base[lo:hi]``.

    An empty `items` list is a deletion; ``lo == hi`` marks a pure insertion at
    boundary `lo`.
    """

    lo: int
    hi: int
    side: str
    items: list
    keys: list

    @property
    def is_insert(self) -> bool:
        return self.lo == self.hi


@dataclass
class _Component:
    """One or more strictly-overlapping range hunks spanning ``base[lo:hi]``."""

    lo: int
    hi: int
    hunks: list[_Hunk] = field(default_factory=list)

    @property
    def sides(self) -> frozenset[str]:
        return frozenset(h.side for h in self.hunks)


def _side_hunks(
    base_keys: Sequence, side_keys: Sequence, side_items: Sequence, side: str
) -> tuple[list, list[_Hunk]]:
    """Return ``(opcodes, hunks)`` for one branch; opcodes are kept for rendering."""
    sm = SequenceMatcher(None, base_keys, side_keys, autojunk=False)
    opcodes = sm.get_opcodes()
    hunks: list[_Hunk] = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            continue
        hunks.append(
            _Hunk(
                i1,
                i2,
                side,
                list(side_items[j1:j2]),
                list(side_keys[j1:j2]),
            )
        )
    return opcodes, hunks


def _components(hunks: list[_Hunk]) -> list[_Component]:
    """Group range hunks whose base intervals strictly overlap.

    Hunks that merely touch at a boundary (e.g. replace ``[0:1)`` on one branch and
    ``[1:2)`` on the other) stay independent.
    """
    ranged = sorted((h for h in hunks if not h.is_insert), key=lambda h: (h.lo, h.hi))
    components: list[_Component] = []
    for hunk in ranged:
        if components and hunk.lo < components[-1].hi:
            component = components[-1]
            component.hunks.append(hunk)
            component.hi = max(component.hi, hunk.hi)
        else:
            components.append(_Component(hunk.lo, hunk.hi, [hunk]))
    return components


def _segment_chunk(opcode, side_items, side_keys, s: int, e: int):
    """Map base interval ``[s:e)`` through one side opcode.

    Returns ``(items, keys, changed)``, or |None| when the segment sits strictly
    inside a replacement hunk whose items cannot be attributed 1:1 to base items
    (e.g. two base items replaced by three side items).
    """
    tag, i1, i2, j1, j2 = opcode
    if tag == "equal":
        a, b = j1 + (s - i1), j1 + (e - i1)
        return list(side_items[a:b]), list(side_keys[a:b]), False
    if (i1, i2) == (s, e):
        return list(side_items[j1:j2]), list(side_keys[j1:j2]), True
    # -- the segment is part of a wider replacement; proportional attribution is
    # only sound when items map one-for-one onto base positions --
    if (j2 - j1) == (i2 - i1):
        a, b = j1 + (s - i1), j1 + (e - i1)
        return list(side_items[a:b]), list(side_keys[a:b]), True
    return None


def _opcode_at(opcodes, s: int, e: int):
    """Return the opcode whose base range covers segment ``[s:e)``."""
    for opcode in opcodes:
        _tag, i1, i2, _j1, _j2 = opcode
        if i1 <= s and e <= i2 and e > i1:
            return opcode
    raise KeyError((s, e))


def _subdivide_component(
    component: _Component,
    left_opcodes,
    right_opcodes,
    base: Sequence,
    left: Sequence,
    right: Sequence,
    kl,
    kr,
):
    """Split a conflict component into minimal per-base-interval decisions.

    Returns a decision list, or |None| when finer attribution is impossible (an
    insertion sits inside the span, or a replacement hunk has unequal cardinality).
    """
    lo, hi = component.lo, component.hi
    if any(h.is_insert and lo < h.lo < hi for h in component.hunks):
        return None

    intersecting = [
        op
        for op in (*left_opcodes, *right_opcodes)
        if op[2] > lo and op[1] < hi
    ]
    bounds = {lo, hi}
    for _tag, i1, i2, _j1, _j2 in intersecting:
        if lo < i1 < hi:
            bounds.add(i1)
        if lo < i2 < hi:
            bounds.add(i2)
    positions = sorted(bounds)

    decisions: List[Decision] = []
    for s, e in zip(positions, positions[1:]):
        left_chunk = _segment_chunk(_opcode_at(left_opcodes, s, e), left, kl, s, e)
        right_chunk = _segment_chunk(_opcode_at(right_opcodes, s, e), right, kr, s, e)
        if left_chunk is None or right_chunk is None:
            return None
        l_items, l_keys, left_changed = left_chunk
        r_items, r_keys, right_changed = right_chunk
        if not left_changed and not right_changed:
            decisions.append(("base", list(base[s:e])))
        elif left_changed and not right_changed:
            if l_items:
                decisions.append(("left", l_items))
        elif right_changed and not left_changed:
            if r_items:
                decisions.append(("right", r_items))
        elif l_keys == r_keys:
            if l_items:
                decisions.append(("left", l_items))
        else:
            decisions.append(
                ("conflict", s, e, list(base[s:e]), l_items, r_items)
            )
    return decisions


def _render(
    opcodes,
    side_items: Sequence,
    side_keys: Sequence,
    lo: int,
    hi: int,
) -> tuple[list, list]:
    """Reconstruct one branch's full version of base interval ``[lo:hi)``."""
    items: list = []
    keys: list = []
    for tag, i1, i2, j1, j2 in opcodes:
        if tag == "equal":
            a, b = max(i1, lo), min(i2, hi)
            if b > a:
                s, e = j1 + (a - i1), j1 + (b - i1)
                items.extend(side_items[s:e])
                keys.extend(side_keys[s:e])
        elif i1 == i2:
            # -- insertion opcode; only those strictly inside the interval --
            if lo < i1 < hi:
                items.extend(side_items[j1:j2])
                keys.extend(side_keys[j1:j2])
        elif i1 < hi and i2 > lo:
            # -- replace/delete; component edges always coincide with hunk edges --
            items.extend(side_items[j1:j2])
            keys.extend(side_keys[j1:j2])
    return items, keys


def three_way_decisions(
    base: Sequence[T],
    left: Sequence[T],
    right: Sequence[T],
    key: Callable[[T], object] | None = None,
) -> List[Decision]:
    """Merge three parallel item sequences into an ordered list of decisions.

    `key` maps an item to a hashable comparison key; items compare by identity when
    not provided.
    """
    if key is None:
        kb: Sequence = list(base)
        kl: Sequence = list(left)
        kr: Sequence = list(right)
    else:
        kb = [key(x) for x in base]
        kl = [key(x) for x in left]
        kr = [key(x) for x in right]

    left_opcodes, left_hunks = _side_hunks(kb, kl, left, "left")
    right_opcodes, right_hunks = _side_hunks(kb, kr, right, "right")
    all_hunks = left_hunks + right_hunks

    components = _components(all_hunks)

    # -- insertions strictly inside a component span render as part of that
    # component (this also blocks subdivision of that component later);
    # everything else is an independent boundary insertion --
    boundary: dict[int, dict[str, _Hunk]] = {}
    for hunk in all_hunks:
        if not hunk.is_insert:
            continue
        host = next(
            (c for c in components if c.lo < hunk.lo < c.hi),
            None,
        )
        if host is not None:
            host.hunks.append(hunk)
        else:
            boundary.setdefault(hunk.lo, {})[hunk.side] = hunk

    decisions: List[Decision] = []
    base_run: list = []

    def flush_base() -> None:
        if base_run:
            decisions.append(("base", base_run.copy()))
            base_run.clear()

    def emit_base(index: int) -> None:
        base_run.append(base[index])

    def emit_inserts(position: int) -> None:
        at = boundary.get(position)
        if not at:
            return
        flush_base()
        lh, rh = at.get("left"), at.get("right")
        if lh is not None and rh is not None:
            if lh.keys == rh.keys:
                decisions.append(("left", lh.items))
            else:
                decisions.append(("insert", lh.items, rh.items))
        elif lh is not None:
            decisions.append(("left", lh.items))
        else:
            assert rh is not None
            decisions.append(("right", rh.items))

    def emit_component(component: _Component) -> None:
        flush_base()
        lo, hi = component.lo, component.hi
        if len(component.sides) == 1:
            side = next(iter(component.sides))
            opcodes = left_opcodes if side == "left" else right_opcodes
            items, _keys = _render(
                opcodes, left if side == "left" else right,
                kl if side == "left" else kr, lo, hi,
            )
            if items:
                decisions.append((side, items))
            return
        l_items, l_keys = _render(left_opcodes, left, kl, lo, hi)
        r_items, r_keys = _render(right_opcodes, right, kr, lo, hi)
        if l_keys == r_keys:
            # -- identical replacement on both branches, take it exactly once --
            if l_items:
                decisions.append(("left", l_items))
            return
        # -- try to narrow the conflict to individual base items, so unrelated
        # edits inside the same alignment region keep merging independently --
        subdivided = _subdivide_component(
            component, left_opcodes, right_opcodes, base, left, right, kl, kr
        )
        if subdivided is not None:
            decisions.extend(subdivided)
        else:
            decisions.append(
                ("conflict", lo, hi, list(base[lo:hi]), l_items, r_items)
            )

    cursor = 0
    for component in sorted(components, key=lambda c: (c.lo, c.hi)):
        for position in range(cursor, component.lo):
            emit_inserts(position)
            emit_base(position)
        emit_inserts(component.lo)
        emit_component(component)
        cursor = component.hi

    for position in range(cursor, len(kb) + 1):
        emit_inserts(position)
        if position < len(kb):
            emit_base(position)

    flush_base()
    return decisions
