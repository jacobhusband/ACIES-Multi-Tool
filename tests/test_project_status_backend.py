"""Exercise status normalization without starting the desktop app."""
import ast
from pathlib import Path
import unittest


class ProjectStatusBackendTests(unittest.TestCase):
    def test_status_roundtrip_preserves_workflow_and_defaults_legacy_blanks(self):
        source = (Path(__file__).resolve().parents[1] / "main.py").read_text(encoding="utf-8")
        constants = source[source.index("STATUS_CANON ="):source.index("# Deliverable PDF quick access")]
        tree = ast.parse(source)
        fn = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "sync_status_arrays")
        scope = {}
        exec(constants, scope)
        exec(compile(ast.Module(body=[fn], type_ignores=[]), "main.py", "exec"), scope)
        normalize = scope["sync_status_arrays"]
        for item in [{}, {"statuses": []}, {"status": "Working"}, {"statusTags": ["working"]}]:
            normalize(item)
            self.assertEqual(item["statuses"], ["In progress"])
            self.assertEqual(item["statusTags"], ["inProgress"])
            self.assertEqual(item["status"], "In progress")
        for status in scope["STATUS_CANON"]:
            with self.subTest(status=status):
                item = {"status": status}
                normalize(item)
                normalize(item)
                self.assertEqual(item["statuses"], [status])
                self.assertEqual(item["status"], status)


if __name__ == "__main__":
    unittest.main()
