"""Regression checks for export caches and screenshot selection; no CATIA required."""

import sys
import unittest
from collections import Counter
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "CATIA_BOM_GUI"))
import catia_bom_export_v2 as core


class Collection:
    def __init__(self, items):
        self.items = list(items)
        self.item_calls = 0
        self.count_calls = 0

    @property
    def Count(self):
        self.count_calls += 1
        return len(self.items)

    def Item(self, index):
        self.item_calls += 1
        return self.items[index - 1]


class Product:
    def __init__(self, number, children=()):
        self.PartNumber = number
        self.Name = number + ".1"
        self.Revision = "A"
        self.Nomenclature = number
        self.Products = Collection(children)


class Selection:
    def __init__(self):
        self.selected = []
        self.show_calls = []
        self.VisProperties = SimpleNamespace(SetShow=self.set_show)

    @property
    def Count(self):
        return len(self.selected)

    def Clear(self):
        self.selected.clear()

    def Add(self, value):
        self.selected.append(value)

    def set_show(self, state):
        self.show_calls.append((tuple(self.selected), state))


class NoIterationList(list):
    def __iter__(self):
        raise AssertionError("Unexpected whole-list scan during a single screenshot")


class TreeCacheTests(unittest.TestCase):
    def test_healthy_same_number_instances_merge_but_broken_instance_survives(self):
        healthy_a, broken, healthy_b = [Product("SAME") for _ in range(3)]
        root = Product("ROOT", [healthy_a, broken, healthy_b])
        doc = SimpleNamespace(Name="SAME.CATPart", FullName="SAME.CATPart")
        results = {id(healthy_a): doc, id(healthy_b): doc, id(broken): None}
        with patch.object(core, "get_part_document", side_effect=lambda p, *_: results[id(p)]) as resolve:
            nodes = core.traverse_product_tree(SimpleNamespace(), root, Mock())

        leaves = [node for node in nodes if node.is_leaf]
        self.assertEqual(len(leaves), 2)
        self.assertEqual([node.tree_quantity for node in leaves], [2, 1])
        self.assertIsNone(leaves[1].part_doc)
        self.assertEqual([id(p) for p in nodes[0].descendant_leaf_products],
                         [id(healthy_a), id(healthy_b), id(broken)])
        items = core.create_bom_items(nodes)
        self.assertEqual(list(items), ["SAME"])
        self.assertEqual(items["SAME"].quantity, 2)
        self.assertEqual(Counter(id(call.args[0]) for call in resolve.call_args_list),
                         Counter({id(healthy_a): 1, id(broken): 1, id(healthy_b): 1}))

    def test_missing_documents_are_not_resolved_again_when_visiting(self):
        parts = [Product("SAME") for _ in range(3)]
        with patch.object(core, "get_part_document", return_value=None) as resolve:
            nodes = core.traverse_product_tree(SimpleNamespace(), Product("ROOT", parts), Mock())
        self.assertEqual(len([node for node in nodes if node.is_leaf]), 3)
        self.assertEqual(resolve.call_count, 3)

    def test_fallback_index_scans_documents_once_and_expires_between_trees(self):
        root = Product("ROOT", [Product("A"), Product("B")])

        def document(number):
            return SimpleNamespace(Name=number + ".CATPart", Product=SimpleNamespace(PartNumber=number))

        first_docs = [document("A"), document("B")]
        documents = Collection(first_docs)
        catia = SimpleNamespace(Documents=documents)
        first_nodes = core.traverse_product_tree(catia, root, Mock())
        self.assertEqual(documents.item_calls, 2)
        self.assertEqual(documents.count_calls, 1)
        self.assertIs(first_nodes[1].part_doc, first_docs[0])
        self.assertIs(first_nodes[2].part_doc, first_docs[1])

        second_docs = [document("A"), document("B")]
        documents.items = second_docs
        second_nodes = core.traverse_product_tree(catia, root, Mock())
        self.assertEqual(documents.item_calls, 4)
        self.assertEqual(documents.count_calls, 2)
        self.assertIs(second_nodes[1].part_doc, second_docs[0])
        self.assertIs(second_nodes[2].part_doc, second_docs[1])

    def test_direct_document_resolution_never_builds_fallback_index(self):
        part = Product("DIRECT")
        part.Parent = SimpleNamespace(Name="DIRECT.CATPart", Part=object())
        documents = Collection([])
        nodes = core.traverse_product_tree(SimpleNamespace(Documents=documents),
                                           Product("ROOT", [part]), Mock())
        self.assertIs(nodes[1].part_doc, part.Parent)
        self.assertEqual(documents.count_calls, 0)


class ScreenshotIndexTests(unittest.TestCase):
    def setUp(self):
        self.a, self.b = object(), object()
        self.doc = SimpleNamespace(Selection=Selection(), Activate=Mock())
        viewer = SimpleNamespace(Update=Mock())
        self.catia = SimpleNamespace(ActiveWindow=SimpleNamespace(Layout=0, ActiveViewer=viewer))
        root = SimpleNamespace(level=0, descendant_leaf_products=[self.a, self.b])
        self.state = core.ScreenshotDisplayState(self.catia, self.doc, [root], Mock())

    def test_owner_index_selects_relevant_and_global_artifacts_without_full_scan(self):
        a_only, b_only, global_item, shared = [object() for _ in range(4)]
        key_a, key_b = core._com_object_identity(self.a), core._com_object_identity(self.b)
        artifacts = [
            (self.doc, a_only, "a", frozenset([key_a])),
            (self.doc, b_only, "b", frozenset([key_b])),
            (self.doc, global_item, "global", frozenset()),
            (self.doc, shared, "shared", frozenset([key_a, key_b])),
        ]
        with patch.object(self.state, "_search_annotation_sets", return_value=[]), \
                patch.object(self.state, "_collect_capture_artifacts", return_value=artifacts):
            self.state.prepare()
        self.state.capture_artifacts = NoIterationList(artifacts)
        with patch.object(core, "_selected_show_state", return_value=0) as read_state:
            self.state.hide_capture_artifacts([self.a])
            self.assertEqual(self.doc.Selection.show_calls[-1], ((a_only, global_item, shared), 1))
            self.assertEqual(read_state.call_count, 3)
            self.state.hide_capture_artifacts([self.b])
            self.assertEqual(self.doc.Selection.show_calls[-1], ((b_only, global_item, shared), 1))
            self.assertEqual(read_state.call_count, 4)
            self.state.hide_capture_artifacts([self.b])
            self.assertEqual(self.doc.Selection.show_calls[-1], ((b_only, global_item, shared), 1))
            self.assertEqual(read_state.call_count, 4)
        self.assertEqual(len(self.state.artifact_states), 4)
        self.assertEqual(self.doc.Selection.selected, [])

    def test_delta_visibility_does_not_scan_all_leaves_after_initial_capture(self):
        self.state.all_leaf_products = [self.a, self.b]
        with patch.object(core, "_set_leaf_products_visibility") as update:
            self.state.show_only([self.a])
            update.reset_mock()
            self.state.all_leaf_products = NoIterationList([self.a, self.b])
            self.assertEqual(self.state.show_only([self.b]), [self.b])
            changes = [(call.args[1], call.kwargs["show"]) for call in update.call_args_list]
            self.assertEqual(changes, [([self.a], False), ([self.b], True)])
            update.reset_mock()
            self.state.show_only([self.b])
            self.assertTrue(all(not call.args[1] for call in update.call_args_list))


if __name__ == "__main__":
    unittest.main()
