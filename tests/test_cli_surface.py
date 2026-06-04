import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "parse_document_compare.py"


def load_module():
    spec = importlib.util.spec_from_file_location("parse_document_compare", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class CliSurfaceTest(unittest.TestCase):
    def test_skill_frontmatter_name(self):
        skill_md = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        self.assertIn("name: pdf-parse-skill", skill_md)
        self.assertIn("$pdf-parse-skill", skill_md)

    def test_cli_exposes_expected_commands(self):
        module = load_module()
        parser = module.build_parser()
        subparsers_action = next(
            action
            for action in parser._actions
            if getattr(action, "dest", None) == "command"
        )
        commands = set(subparsers_action.choices)
        expected = {
            "compare",
            "convert",
            "batch",
            "scan-dir",
            "tables",
            "doctor",
            "metadata",
            "chunk",
            "render-pages",
            "ocr",
            "auto",
            "extract-fields",
            "export-xlsx",
            "layout-json",
            "verify-fields",
            "classify",
            "knowledge-pack",
            "batch-knowledge",
            "qa",
            "diff-docs",
        }
        self.assertLessEqual(expected, commands)

    def test_qa_over_chunks_jsonl(self):
        module = load_module()
        with tempfile.TemporaryDirectory() as tmp:
            chunks = Path(tmp) / "chunks.jsonl"
            chunks.write_text(
                json.dumps(
                    {
                        "chunk_id": 1,
                        "text": "合同金额为100元。付款期限为30天。",
                        "source_file": "contract.pdf",
                        "page": 2,
                        "parser": "test",
                    },
                    ensure_ascii=False,
                )
                + "\n",
                encoding="utf-8",
            )

            payload = module.build_qa_payload(
                chunks,
                "合同金额是多少？",
                "auto",
                1,
                30,
                "page",
                2000,
                200,
                3,
                3,
            )

        self.assertEqual(payload["type"], "qa")
        self.assertEqual(payload["answer"]["status"], "found")
        self.assertTrue(payload["citations"])

    def test_diff_output_writer(self):
        module = load_module()
        payload = {
            "similarity": {"full_text_ratio": 0.8, "first_50000_chars": 0.8},
            "line_diff": {"changed_block_count": 1, "stats": {"replace": 1}},
            "field_changes": [{"field": "total", "left": "100", "right": "120"}],
            "unified_diff": ["--- old", "+++ new", "-100", "+120"],
            "left": {"file": "old.pdf"},
            "right": {"file": "new.pdf"},
        }
        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "diff"
            written = module.write_diff_output(payload, out_dir, "all")

            self.assertIn(out_dir / "diff_report.md", written)
            self.assertIn(out_dir / "diff_report.json", written)
            self.assertTrue((out_dir / "diff_report.md").exists())
            self.assertTrue((out_dir / "diff_report.json").exists())


if __name__ == "__main__":
    unittest.main()
