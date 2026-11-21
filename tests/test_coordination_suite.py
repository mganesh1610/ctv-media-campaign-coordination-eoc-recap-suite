#!/usr/bin/env python3
"""Unit and integration tests for CTV & Digital Media Campaign Coordination & EOC Recap Suite."""

from __future__ import annotations

import json
import sqlite3
import unittest
import zipfile
from pathlib import Path

from src.campaign_coordination_engine import (
    DB_PATH,
    OUTPUT_DIR,
    evaluate_deliverable_spec_rules,
    run_pipeline,
)


class TestMediaCampaignCoordinationSuite(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.summary = run_pipeline()

    def test_all_12_validation_gates_pass(self) -> None:
        self.assertEqual(self.summary["validation_gates_passed"], 12)
        self.assertEqual(self.summary["validation_gates_total"], 12)
        for gate_name, passed in self.summary["gates"].items():
            self.assertTrue(passed, f"Gate failed: {gate_name}")

    def test_row_conservation_and_defect_interception(self) -> None:
        self.assertEqual(self.summary["total_source_records"], 240144)
        self.assertEqual(self.summary["creative_packages_audited"], 500)
        self.assertEqual(self.summary["seeded_defects"], 60)
        self.assertEqual(self.summary["defects_intercepted"], 60)
        self.assertEqual(self.summary["post_remediation_launch_ready"], 500)

    def test_zero_billing_variance_and_sql_integrity(self) -> None:
        self.assertEqual(self.summary["billing_variance_usd"], 0.0)
        conn = sqlite3.connect(DB_PATH)
        try:
            audit = conn.execute("SELECT total_orders, total_abs_billing_variance_usd FROM vw_billing_reconciliation_audit").fetchone()
            self.assertEqual(audit[0], 48)
            self.assertEqual(audit[1], 0.0)
            fk = conn.execute("PRAGMA foreign_key_check").fetchall()
            self.assertEqual(fk, [])
        finally:
            conn.close()

    def test_holdout_optimization_lifts(self) -> None:
        self.assertGreater(self.summary["shortfall_reduction_pct"], 50.0)
        self.assertGreater(self.summary["ctr_relative_lift_pct"], 0.0)
        self.assertGreater(self.summary["cpc_reduction_pct"], 0.0)
        self.assertGreater(self.summary["optimized_delivery_rate_pct"], self.summary["baseline_delivery_rate_pct"])

    def test_office_deliverables_structure(self) -> None:
        xlsx_path = OUTPUT_DIR / "media_campaign_coordination_tracker.xlsx"
        pptx_path = OUTPUT_DIR / "end_of_campaign_recap_deck.pptx"
        self.assertTrue(xlsx_path.exists())
        self.assertTrue(pptx_path.exists())
        with zipfile.ZipFile(xlsx_path) as zf:
            names = set(zf.namelist())
            self.assertIn("xl/workbook.xml", names)
            self.assertIn("xl/worksheets/sheet5.xml", names)
            sheet5 = zf.read("xl/worksheets/sheet5.xml").decode("utf-8")
            self.assertIn("<f>SUM(", sheet5)
        with zipfile.ZipFile(pptx_path) as zf:
            names = set(zf.namelist())
            self.assertIn("ppt/presentation.xml", names)
            self.assertIn("ppt/slides/slide8.xml", names)

    def test_rule_engine_catches_custom_unit_and_vast_errors(self) -> None:
        bad_pkg = {
            "asset_url": "http://bad.example/ad.mp4",
            "has_click_macro": False,
            "requires_vast_42": True,
            "vast_42_compliant": False,
            "bitrate_kbps": 20000,
            "max_bitrate_kbps": 10000,
            "resolution": "1280x720",
            "spec_resolution": "1920x1080",
            "io_signed": False,
            "lead_time_hours": 10.0,
            "min_lead_time_days": 5,
            "is_custom_unit": True,
            "custom_unit_qa_signoff": False,
            "brand_safety_verified": False,
            "billing_cost_center_valid": False,
        }
        ok, failures = evaluate_deliverable_spec_rules(bad_pkg)
        self.assertFalse(ok)
        self.assertEqual(len(failures), 10)


if __name__ == "__main__":
    unittest.main()
