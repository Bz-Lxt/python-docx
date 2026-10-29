# pyright: reportPrivateUsage=false

"""Unit tests for the three-way sequence merge in `docx.merge.diff3`."""

from __future__ import annotations

from docx.merge.diff3 import three_way_decisions


def _flatten(decisions):
    """Render a decision list into (tag-marker, item) tuples for order assertions."""
    out = []
    for decision in decisions:
        kind = decision[0]
        if kind in ("base", "left", "right"):
            out.extend((kind, item) for item in decision[1])
        elif kind == "insert":
            out.extend(("left", item) for item in decision[1])
            out.extend(("right", item) for item in decision[2])
        else:
            _, lo, hi, base, left, right = decision
            out.append(("conflict", (lo, hi, tuple(base), tuple(left), tuple(right))))
    return out


class DescribeThreeWayDecisions:
    def it_returns_the_base_when_nothing_changed(self):
        decisions = three_way_decisions(["a", "b"], ["a", "b"], ["a", "b"])

        assert decisions == [("base", ["a", "b"])]

    def it_takes_a_change_made_on_only_one_branch(self):
        left_only = three_way_decisions(["a"], ["A"], ["a"])
        right_only = three_way_decisions(["a"], ["a"], ["A"])

        assert left_only == [("left", ["A"])]
        assert right_only == [("right", ["A"])]

    def it_takes_an_identical_change_on_both_branches_once(self):
        decisions = three_way_decisions(["a"], ["A"], ["A"])

        assert decisions == [("left", ["A"])]

    def it_applies_a_delete_that_only_one_branch_made(self):
        decisions = three_way_decisions(["a", "b"], ["b"], ["a", "b"])

        assert all(kind != "conflict" for kind, *_ in decisions)
        assert [item for d in decisions for item in d[1]] == ["b"]

    def it_merges_independent_edits_to_adjacent_items_without_an_anchor(self):
        # -- left edits items 0 and 2, right edits item 1; no unchanged item sits
        # between the edits, yet they must not collide --
        decisions = three_way_decisions(
            ["a", "b", "c"], ["A", "b", "C"], ["a", "B", "c"]
        )

        assert all(d[0] != "conflict" for d in decisions)
        assert _flatten(decisions) == [
            ("left", "A"),
            ("right", "B"),
            ("left", "C"),
        ]

    def it_reports_overlapping_different_edits_as_a_conflict(self):
        decisions = three_way_decisions(["a"], ["A"], ["B"])

        assert len(decisions) == 1
        kind, lo, hi, base, left, right = decisions[0]
        assert kind == "conflict"
        assert (lo, hi) == (0, 1)
        assert (base, left, right) == (["a"], ["A"], ["B"])

    def it_reports_delete_against_edit_of_the_same_item_as_a_conflict(self):
        # -- both branches touch item 0: left deletes it, right edits it --
        decisions = three_way_decisions(
            ["a", "b"], ["b"], ["A", "b"]
        )

        conflicts = [d for d in decisions if d[0] == "conflict"]
        assert len(conflicts) == 1
        _, lo, hi, base, left, right = conflicts[0]
        assert (lo, hi) == (0, 1)
        assert base == ["a"]
        assert left == []
        assert right == ["A"]

    def it_merges_a_delete_of_one_item_and_an_edit_of_another(self):
        # -- left deletes item 0, right edits item 1: disjoint, no conflict --
        decisions = three_way_decisions(
            ["a", "b"], ["b"], ["a", "B"]
        )

        assert all(d[0] != "conflict" for d in decisions)
        assert _flatten(decisions) == [("right", "B")]

    def it_keeps_both_different_insertions_at_the_same_boundary(self):
        decisions = three_way_decisions(["a"], ["x", "a"], ["y", "a"])

        assert decisions[0] == ("insert", ["x"], ["y"])
        assert decisions[1] == ("base", ["a"])

    def it_dedupes_identical_insertions_at_the_same_boundary(self):
        decisions = three_way_decisions(["a"], ["x", "a"], ["x", "a"])

        assert decisions == [("left", ["x"]), ("base", ["a"])]

    def it_keeps_insertions_at_different_boundaries_in_position(self):
        decisions = three_way_decisions(
            ["a", "b"], ["x", "a", "b"], ["a", "b", "y"]
        )

        assert _flatten(decisions) == [
            ("left", "x"),
            ("base", "a"),
            ("base", "b"),
            ("right", "y"),
        ]

    def it_conflicts_when_an_insert_sits_inside_a_replaced_span(self):
        # -- left replaces both base items wholesale; right inserts between them --
        decisions = three_way_decisions(
            ["a", "b"], ["L"], ["a", "R", "b"]
        )

        conflicts = [d for d in decisions if d[0] == "conflict"]
        assert len(conflicts) == 1
        _, lo, hi, base, left, right = conflicts[0]
        assert (lo, hi) == (0, 2)
        assert left == ["L"]
        assert right == ["a", "R", "b"]

    def it_subdivides_a_wide_replace_hunk_to_attribute_each_item_to_a_branch(self):
        # -- right reworks two adjacent items (one wide replace hunk) while left
        # only edits the first of them: the component is split per item so the
        # right-only edit does not escalate the whole span --
        decisions = three_way_decisions(
            ["a", "b", "c"], ["A", "b", "c"], ["RA", "RB", "c"]
        )

        assert decisions == [
            ("conflict", 0, 1, ["a"], ["A"], ["RA"]),
            ("right", ["RB"]),
            ("base", ["c"]),
        ]

    def it_subdivides_a_wide_replace_hunk_on_the_left_side_as_well(self):
        decisions = three_way_decisions(
            ["a", "b", "c"], ["LA", "LB", "c"], ["a", "B", "c"]
        )

        assert decisions == [
            ("left", ["LA"]),
            ("conflict", 1, 2, ["b"], ["LB"], ["B"]),
            ("base", ["c"]),
        ]

    def it_combines_disjoint_edits_and_insertions(self):
        decisions = three_way_decisions(
            ["a", "b", "c", "d"],
            ["A", "b", "c", "d", "L"],
            ["a", "x", "B", "c", "d"],
        )

        assert all(d[0] != "conflict" for d in decisions)
        assert _flatten(decisions) == [
            ("left", "A"),
            ("right", "x"),
            ("right", "B"),
            ("base", "c"),
            ("base", "d"),
            ("left", "L"),
        ]

    def it_renders_each_side_of_a_multiitem_conflict_with_its_own_changes(self):
        # -- both sides rework the same two items differently --
        decisions = three_way_decisions(
            ["a", "b"], ["LA", "LB"], ["RA", "RB"]
        )

        conflicts = [d for d in decisions if d[0] == "conflict"]
        assert len(conflicts) == 1
        _, lo, hi, base, left, right = conflicts[0]
        assert (lo, hi) == (0, 2)
        assert base == ["a", "b"]
        assert left == ["LA", "LB"]
        assert right == ["RA", "RB"]

    def it_maps_items_through_the_key_callable(self):
        # -- keys equal -> treated as unchanged -> the base item is emitted --
        decisions = three_way_decisions(
            ["a"], ["A"], ["a"], key=str.lower
        )

        assert decisions == [("base", ["a"])]

    def it_is_deterministic_for_the_same_inputs(self):
        args = (["a", "b", "c"], ["A", "b"], ["a", "B", "C"])

        first = three_way_decisions(*args)
        second = three_way_decisions(*args)

        assert first == second
