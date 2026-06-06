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
        self.assertNotIn("document-intelligence-skill", skill_md)

    def test_agent_metadata_uses_current_trigger(self):
        metadata = (ROOT / "agents" / "openai.yaml").read_text(encoding="utf-8")
        self.assertIn("$pdf-parse-skill", metadata)
        self.assertNotIn("document-intelligence-skill", metadata)

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
            "vote",
            "customer-pack",
            "batch-customer-pack",
            "probe",
            "convert",
            "batch",
            "scan-dir",
            "tables",
            "table-vote",
            "eval-golden",
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

    def test_vote_command_exposes_customer_profile_options(self):
        module = load_module()
        parser = module.build_parser()
        subparsers_action = next(
            action
            for action in parser._actions
            if getattr(action, "dest", None) == "command"
        )
        vote_parser = subparsers_action.choices["vote"]
        option_dests = {action.dest for action in vote_parser._actions}

        self.assertIn("profile", option_dests)
        self.assertIn("customer", option_dests)
        self.assertIn("no_field_fusion", option_dests)
        self.assertIn("field_layout", option_dests)
        self.assertIn("field_layout_max_pages", option_dests)
        self.assertIn("probe_before_vote", option_dests)
        self.assertIn("probe_max_pages", option_dests)
        self.assertIn("probe_min_quality", option_dests)
        self.assertIn("timeout", option_dests)
        self.assertIn("parser_health_cache", option_dests)

    def test_eval_golden_command_exposes_regression_options(self):
        module = load_module()
        parser = module.build_parser()
        subparsers_action = next(
            action
            for action in parser._actions
            if getattr(action, "dest", None) == "command"
        )
        eval_parser = subparsers_action.choices["eval-golden"]
        option_dests = {action.dest for action in eval_parser._actions}

        self.assertIn("path", option_dests)
        self.assertIn("recursive", option_dests)
        self.assertIn("max_pages", option_dests)
        self.assertIn("parsers", option_dests)
        self.assertIn("profile", option_dests)
        self.assertIn("timeout", option_dests)
        self.assertIn("parser_health_cache", option_dests)
        self.assertIn("health_cache", option_dests)

    def test_customer_pack_command_exposes_complex_pdf_options(self):
        module = load_module()
        parser = module.build_parser()
        subparsers_action = next(
            action
            for action in parser._actions
            if getattr(action, "dest", None) == "command"
        )
        pack_parser = subparsers_action.choices["customer-pack"]
        option_dests = {action.dest for action in pack_parser._actions}

        self.assertIn("table_pages", option_dests)
        self.assertIn("layout_max_pages", option_dests)
        self.assertIn("no_tables", option_dests)
        self.assertIn("no_layout", option_dests)
        self.assertIn("no_field_fusion", option_dests)
        self.assertIn("profile", option_dests)
        self.assertIn("timeout", option_dests)
        self.assertIn("parser_health_cache", option_dests)

    def test_table_vote_and_batch_customer_pack_commands_expose_options(self):
        module = load_module()
        parser = module.build_parser()
        subparsers_action = next(
            action
            for action in parser._actions
            if getattr(action, "dest", None) == "command"
        )
        table_parser = subparsers_action.choices["table-vote"]
        table_dests = {action.dest for action in table_parser._actions}
        self.assertIn("methods", table_dests)
        self.assertIn("pages", table_dests)
        self.assertIn("out_dir", table_dests)

        batch_parser = subparsers_action.choices["batch-customer-pack"]
        batch_dests = {action.dest for action in batch_parser._actions}
        self.assertIn("recursive", batch_dests)
        self.assertIn("table_pages", batch_dests)
        self.assertIn("timeout", batch_dests)
        self.assertIn("parser_health_cache", batch_dests)
        self.assertIn("no_field_fusion", batch_dests)

    def test_pdf_parser_registry_includes_markdown_parsers(self):
        module = load_module()
        pdf_parsers = set(module.get_parsers_for_format("pdf"))
        expected = {"pymupdf4llm", "docling", "pspdfkit", "ocr-tesseract"}
        self.assertLessEqual(expected, pdf_parsers)
        for parser_name in expected:
            self.assertIsNotNone(module.get_extractor("pdf", parser_name))

    def test_external_cli_parser_dependency_row(self):
        module = load_module()
        rows = module.parser_dependency_rows("pdf")
        pspdfkit_row = next(row for row in rows if row["parser"] == "pspdfkit")
        self.assertEqual(pspdfkit_row["kind"], "command")
        self.assertIn("pdf-to-markdown", pspdfkit_row["commands"])
        self.assertEqual(module.resolve_parser_name("pdf-to-markdown"), "pspdfkit")

        ocr_row = next(row for row in rows if row["parser"] == "ocr-tesseract")
        self.assertEqual(ocr_row["kind"], "ocr")
        self.assertIn("tesseract", ocr_row["commands"])

    def test_vote_payload_selects_best_parser(self):
        module = load_module()
        strong_text = "# 标题\n\n合同金额为100元。付款期限为30天。\n\n| 项目 | 金额 |\n|---|---|\n| 服务 | 100 |\n"
        weak_text = "(cid:123)(cid:456)"
        results = [
            module.result_from_text("strong", strong_text, 0.1),
            module.result_from_text("weak", weak_text, 0.1),
        ]
        normalized = {
            "strong": module.normalize_text(strong_text),
            "weak": module.normalize_text(weak_text),
        }
        payload = module.build_vote_payload(
            ROOT / "sample.pdf",
            ROOT / "sample.pdf",
            "pdf",
            results,
            normalized,
            50000,
            0.5,
            True,
        )
        self.assertEqual(payload["winner"]["parser"], "strong")
        self.assertGreater(payload["winner"]["vote_score"], 0)

    def test_probe_payload_marks_runtime_failures(self):
        module = load_module()
        ok_text = "发票号码:123456789012\n开票日期:2026年06月05日\n价税合计(小写)¥113.00"
        results = [
            module.result_from_text("ready", ok_text, 0.1),
            module.ParseResult(parser="broken", status="failed", seconds=0.2, error="model download failed"),
        ]
        normalized = {"ready": module.normalize_text(ok_text)}
        original_dependency_rows = module.parser_dependency_rows

        def fake_dependency_rows(fmt_filter=None):
            return [
                {
                    "parser": "ready",
                    "available": True,
                    "kind": "python",
                    "modules": ["ready"],
                    "commands": [],
                },
                {
                    "parser": "broken",
                    "available": True,
                    "kind": "python",
                    "modules": ["broken"],
                    "commands": [],
                },
            ]

        module.parser_dependency_rows = fake_dependency_rows
        try:
            payload = module.build_probe_payload(
                ROOT / "sample.pdf",
                ROOT / "sample.pdf",
                "pdf",
                ["ready", "broken"],
                results,
                normalized,
                0.5,
                True,
                1,
                1,
            )
        finally:
            module.parser_dependency_rows = original_dependency_rows

        rows = {row["parser"]: row for row in payload["rows"]}
        self.assertEqual(rows["ready"]["probe_status"], "ready")
        self.assertEqual(rows["broken"]["probe_status"], "runtime_failed")
        self.assertIn("broken", payload["diagnostics"][0])

    def test_vote_payload_penalizes_repeated_text(self):
        module = load_module()
        clean_text = "\n".join([
            "项目概述：系统提供合同台账、付款节点和审批记录。",
            "金额说明：合同金额为100元，付款期限为30天。",
            "交付要求：验收后生成归档文件并通知业务负责人。",
            "风险提示：逾期付款需要重新确认审批状态。",
            "联系人：测试负责人，电话字段不在本次解析范围。",
        ])
        repeated_text = "\n".join(["合同金额为100元，付款期限为30天。"] * 24)
        results = [
            module.result_from_text("repeated", repeated_text, 0.1),
            module.result_from_text("clean", clean_text, 0.1),
        ]
        normalized = {
            "repeated": module.normalize_text(repeated_text),
            "clean": module.normalize_text(clean_text),
        }

        payload = module.build_vote_payload(
            ROOT / "sample.pdf",
            ROOT / "sample.pdf",
            "pdf",
            results,
            normalized,
            50000,
            0.5,
            True,
            profile="none",
        )
        repeated_vote = next(item for item in payload["votes"] if item["parser"] == "repeated")

        self.assertEqual(payload["winner"]["parser"], "clean")
        self.assertGreater(repeated_vote["duplicate_penalty"], 0)

    def test_vote_payload_records_preflight_probe(self):
        module = load_module()
        text = "发票号码:123456789012\n开票日期:2026年06月05日\n价税合计(小写)¥113.00"
        preflight = {
            "summary": {
                "parser_count": 2,
                "ready_count": 1,
                "dependency_missing_count": 1,
                "runtime_failed_count": 0,
                "quality_failed_count": 0,
                "ready_parsers": ["ready"],
            }
        }
        payload = module.build_vote_payload(
            ROOT / "invoice.pdf",
            ROOT / "invoice.pdf",
            "pdf",
            [module.result_from_text("ready", text, 0.1)],
            {"ready": module.normalize_text(text)},
            50000,
            0.5,
            True,
            profile="invoice",
            preflight_probe=preflight,
        )
        markdown = module.vote_to_markdown(payload)

        self.assertEqual(module.ready_parsers_from_probe(preflight), ["ready"])
        self.assertEqual(payload["preflight_probe"], preflight)
        self.assertIn("Preflight probe enabled", "\n".join(payload["diagnostics"]))
        self.assertIn("## Preflight Probe", markdown)
        self.assertIn("Final vote parsers: `ready`", markdown)

    def test_table_quality_and_vote_selects_best_method(self):
        module = load_module()
        good = [
            {
                "page": 1,
                "table_index": 1,
                "headers": ["项目", "金额"],
                "rows": [["服务", "100.00"], ["税额", "13.00"]],
            }
        ]
        weak = [
            {
                "page": 1,
                "table_index": 1,
                "headers": ["项目"],
                "rows": [[]],
            }
        ]
        original = module.extract_tables_by_method

        def fake_extract(_path, _pages, method):
            if method == "good":
                return module.annotate_tables(good, method)
            if method == "weak":
                return module.annotate_tables(weak, method)
            return []

        module.extract_tables_by_method = fake_extract
        try:
            payload = module.build_table_vote_payload(ROOT / "sample.pdf", [0], ["weak", "good"])
        finally:
            module.extract_tables_by_method = original

        self.assertEqual(payload["winner"]["method"], "good")
        self.assertGreater(payload["winner"]["avg_quality"], 0.5)
        self.assertEqual(len(payload["best_tables"]), 1)

    def test_parser_health_cache_can_skip_recent_timeout(self):
        module = load_module()
        key = "pdf|docling|sample|1|1"
        cache = {"entries": {}}
        result = module.ParseResult(parser="docling", status="timeout", seconds=2.0, error="timeout_after_2s")
        module.update_parser_health_cache(cache, key, "docling", result, 2.0)
        skip, reason = module.should_skip_parser_from_health_cache(cache, key)

        self.assertTrue(skip)
        self.assertIn("timeout", reason)

    def test_invoice_vote_and_customer_delivery_prefer_valid_fields(self):
        module = load_module()
        invoice_text = "\n".join([
            "电子发票（普通发票）",
            "发票号码:123456789012",
            "开票日期:2026年06月05日",
            "购买方信息 名称:测试买方公司",
            "统一社会信用代码/纳税人识别号:91330100MA1234567A",
            "销售方信息 名称:测试销售公司",
            "统一社会信用代码/纳税人识别号:91440100MA7654321B",
            "项目名称 单位 数量 单价 金额 税率 税额",
            "*信息技术服务*软件服务 项 1 100.00 100.00 13% 13.00",
            "合计¥100.00¥13.00",
            "价税合计(大写)壹佰壹拾叁元整(小写)¥113.00",
            "开票人:张三",
        ])
        repeated_invoice_text = invoice_text.replace(
            "*信息技术服务*软件服务 项 1 100.00 100.00 13% 13.00",
            "\n".join(["*信息技术服务*软件服务 项 1 100.00 100.00 13% 13.00"] * 4),
        )
        results = [
            module.result_from_text("repeated_invoice", repeated_invoice_text, 0.1),
            module.result_from_text("clean_invoice", invoice_text, 0.1),
        ]
        normalized = {
            "repeated_invoice": module.normalize_text(repeated_invoice_text),
            "clean_invoice": module.normalize_text(invoice_text),
        }

        payload = module.build_vote_payload(
            ROOT / "invoice.pdf",
            ROOT / "invoice.pdf",
            "pdf",
            results,
            normalized,
            50000,
            0.5,
            True,
            profile="invoice",
        )
        repeated_vote = next(item for item in payload["votes"] if item["parser"] == "repeated_invoice")

        self.assertEqual(payload["winner"]["parser"], "clean_invoice")
        self.assertEqual(payload["profile"]["resolved"], "invoice")
        self.assertGreater(payload["winner"]["invoice_score"], repeated_vote["invoice_score"])

        with tempfile.TemporaryDirectory() as tmp:
            out_dir = Path(tmp) / "customer"
            written = module.write_customer_outputs(payload, normalized, out_dir, "all")
            self.assertEqual({path.name for path in written}, {"customer_best.md", "customer_best.txt", "customer_best.json"})
            customer_json = json.loads((out_dir / "customer_best.json").read_text(encoding="utf-8"))

        self.assertEqual(customer_json["type"], "customer_invoice_delivery")
        self.assertEqual(customer_json["invoice"]["invoice_number"], "123456789012")
        self.assertIn("field_confidence", customer_json)
        self.assertGreater(customer_json["field_confidence"]["invoice_number"]["confidence"], 0)
        self.assertEqual(customer_json["field_confidence"]["invoice_number"]["source_parser"], "clean_invoice")

    def test_field_confidence_uses_layout_bbox_when_available(self):
        module = load_module()
        payload = {
            "winner": {"parser": "clean_invoice"},
            "votes": [
                {
                    "parser": "clean_invoice",
                    "eligible": True,
                    "invoice_fields": {
                        "invoice_number": "123456789012",
                        "total_with_tax": 113.00,
                    },
                }
            ],
        }
        layout = {
            "pages": [
                {
                    "page": 1,
                    "blocks": [
                        {
                            "block_id": 1,
                            "lines": [
                                {
                                    "line_id": 1,
                                    "text": "发票号码:123456789012",
                                    "bbox": [10, 20, 150, 35],
                                    "spans": [
                                        {
                                            "span_id": 1,
                                            "text": "发票号码:123456789012",
                                            "bbox": [10, 20, 150, 35],
                                        }
                                    ],
                                },
                                {
                                    "line_id": 2,
                                    "text": "价税合计(小写)¥113.00",
                                    "bbox": [10, 80, 180, 95],
                                    "spans": [
                                        {
                                            "span_id": 1,
                                            "text": "价税合计(小写)¥113.00",
                                            "bbox": [10, 80, 180, 95],
                                        }
                                    ],
                                },
                            ],
                        }
                    ],
                }
            ]
        }

        confidence = module.build_field_confidence(payload, {"clean_invoice": ""}, layout)

        self.assertEqual(confidence["invoice_number"]["page"], 1)
        self.assertEqual(confidence["invoice_number"]["bbox"], [10.0, 20.0, 150.0, 35.0])
        self.assertEqual(confidence["invoice_number"]["location"]["match_type"], "span_contains")
        self.assertEqual(confidence["total_with_tax"]["bbox"], [10.0, 80.0, 180.0, 95.0])

    def test_invoice_customer_delivery_can_fuse_fields_across_parsers(self):
        module = load_module()
        winner_fields = {
            "invoice_type": "电子发票（普通发票）",
            "invoice_number": "123456789012",
            "invoice_date": "2026年06月05日",
            "buyer_name": "测试买方公司",
            "buyer_tax_id": "91330100MA1234567A",
            "seller_name": "测试销售公司",
            "seller_tax_id": "91440100MA7654321B",
            "total_amount": 100.0,
            "total_tax": 13.0,
            "total_with_tax": 113.0,
            "total_with_tax_cn": "壹佰壹拾叁元整",
            "drawer": "张三",
            "items": [
                {
                    "name": "*信息技术服务*软件服务",
                    "unit": "项",
                    "quantity": 1.0,
                    "unit_price": 100.0,
                    "amount": 100.0,
                    "tax_rate": "13%",
                    "tax_amount": 13.0,
                }
            ],
        }
        winner_fields["validation"] = module.validate_invoice_fields(winner_fields, strict=True)
        alternate_fields = dict(winner_fields)
        alternate_fields["buyer_name"] = "测试买方有限公司"
        alternate_fields["validation"] = module.validate_invoice_fields(alternate_fields, strict=True)
        payload = {
            "source_file": "invoice.pdf",
            "profile": {"resolved": "invoice"},
            "weights": {},
            "winner": {
                "parser": "winner_parser",
                "vote_score": 0.8,
                "invoice_score": 0.8,
                "quality_score": 0.8,
                "quality_label": "high",
                "invoice_fields": winner_fields,
            },
            "votes": [
                {
                    "parser": "winner_parser",
                    "eligible": True,
                    "vote_score": 0.8,
                    "invoice_score": 0.8,
                    "quality_score": 0.8,
                    "invoice_fields": winner_fields,
                },
                {
                    "parser": "alternate_parser",
                    "eligible": True,
                    "vote_score": 0.95,
                    "invoice_score": 0.95,
                    "quality_score": 0.95,
                    "invoice_fields": alternate_fields,
                },
            ],
        }

        delivery = module.build_customer_invoice_delivery(payload, {}, None, field_fusion=True)

        self.assertEqual(delivery["parser"], "field-fusion")
        self.assertEqual(delivery["base_parser"], "winner_parser")
        self.assertEqual(delivery["invoice"]["buyer_name"], "测试买方有限公司")
        self.assertTrue(delivery["audit"]["field_fusion"]["used"])
        self.assertIn("buyer_name", delivery["audit"]["field_fusion"]["changed_fields"])

    def test_vote_expectations_compare_invoice_fields(self):
        module = load_module()
        payload = {
            "winner": {
                "parser": "clean_invoice",
                "vote_score": 0.91,
                "invoice_fields": {
                    "invoice_number": "123456789012",
                    "buyer_name": "测试买方公司",
                    "validation": {"status": "ok"},
                },
            },
            "profile": {"resolved": "invoice"},
            "quality_gate": {"passed": True},
        }
        expected = {
            "winner_parser": "clean_invoice",
            "profile_resolved": "invoice",
            "min_vote_score": 0.9,
            "validation_status": "ok",
            "fields": {
                "invoice_number": "123456789012",
                "buyer_name": {"contains": "买方"},
            },
        }

        evaluation = module.evaluate_vote_expectations(payload, expected)

        self.assertTrue(evaluation["passed"])
        self.assertEqual(evaluation["check_count"], 6)

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
