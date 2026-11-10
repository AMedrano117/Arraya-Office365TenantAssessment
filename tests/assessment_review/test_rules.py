"""Review-rule tests on fictional packages."""
from __future__ import annotations

import unittest

from openpyxl import load_workbook

from tests.assessment_review.helpers import PackageTestCase
from tools.assessment_review.engine import (CUSTOMER, FACT, GAP, MSDOC, REVIEWER, SCRIPT,
                                            STATEMENT_TYPES)


class ReviewRuleTests(PackageTestCase):
    def setUp(self):
        super().setUp()
        self.pkg, self.items, self.readiness, self.ctx = self.review(self.build())

    def test_every_statement_is_typed(self):
        for item in self.items.values():
            self.assertTrue(item["Statements"], item["FindingId"])
            for s in item["Statements"]:
                self.assertIn(s["Type"], STATEMENT_TYPES)

    def test_registration_versus_enforcement(self):
        item = self.items["ENT-005"]
        self.assertTrue(any("not that MFA is enforced" in s for s in self.statements(item, REVIEWER)))
        self.assertIn("registration, not enforcement", item["CorrectedInterpretation"])
        self.assertEqual(item["ReadinessTreatment"], "ConfirmBeforePilot")

    def test_configuration_versus_operation(self):
        dlp = self.items["PUR-040"]
        self.assertEqual(dlp["ReviewDecision"], "ConfirmedWithQualification")
        self.assertEqual(dlp["FoundationalOrAISpecific"], "AI-specific")
        audit_gap = self.items["GAP-DATA-AUDIT"]
        self.assertEqual(audit_gap["ReviewDecision"], "ConfirmedWithQualification")
        self.assertTrue(self.statements(audit_gap, MSDOC))

    def test_inventory_versus_entitlement(self):
        lic = self.items["M365-129"]
        self.assertTrue(any("not user entitlement" in s for s in self.statements(lic, REVIEWER)))
        gap = self.items["GAP-LICENSE-ASSIGNMENT"]
        self.assertEqual(gap["ReviewDecision"], "CorrectFinding")
        self.assertEqual(gap["EvidenceSufficiency"], "Contradictory")

    def test_records_versus_distinct_entities(self):
        item = self.items["DEX-001"]
        check = [c for c in item["ConsistencyChecks"] if c["CheckId"] == "C-RECORDS"][0]
        self.assertEqual(check["Result"], "Pass")
        self.assertTrue(any("not distinct files or people" in s for s in self.statements(item, REVIEWER)))
        self.assertEqual(item["EvidenceSufficiency"], "Stale")

    def test_tenant_wide_gap_blocks_without_a_pilot_scope(self):
        # No pilot population is needed: a confirmed tenant-wide gap affects every Copilot user.
        self.assertEqual(self.items["ENT-020"]["ReadinessTreatment"], "BlocksPilot")
        self.assertEqual(self.readiness["ConfirmedBlockers"], ["ENT-020"])
        self.assertEqual(self.readiness["ReviewDecision"], "Not ready: fix the blocking items first")
        # Script blockers the evidence does not support are still re-rated, with a reason.
        for rid in ("ENT-005", "ENT-017", "DEF-013"):
            self.assertNotEqual(self.items[rid]["ReadinessTreatment"], "BlocksPilot", rid)
            self.assertTrue(self.items[rid]["ConclusionChanges"], rid)
            self.assertTrue(self.items[rid]["Plain"]["Change"], rid)

    def test_universal_gap_blocks(self):
        _, items, readiness, _ = self.review(self.build("nomfa", enforced_targeted_mfa=0))
        self.assertEqual(items["ENT-020"]["ReadinessTreatment"], "BlocksPilot")
        self.assertEqual(readiness["ReviewDecision"], "Not ready: fix the blocking items first")

    def test_no_pilot_scope_is_requested(self):
        for root in (self.build("plain"), self.build("scoped", pilot_scope=True)):
            _, items, readiness, _ = self.review(root)
            self.assertNotIn("RVW-PILOT-SCOPE", items)
            self.assertNotIn("PilotScopeRecorded", readiness)
            for item in items.values():
                for value in item["Plain"].values():
                    self.assertNotRegex(value, r"(?i)\bpilot\b", item["FindingId"])

    def test_foundation_before_ai_specific(self):
        ordered = list(self.items.values())
        steps = [i["DependencyStep"] for i in ordered]
        self.assertEqual(steps, sorted(steps))
        kinds = [i["FoundationalOrAISpecific"] for i in ordered]
        self.assertEqual(kinds, sorted(kinds, key=lambda k: k != "Foundational"))
        dspm = self.items["DEX-003"]
        self.assertEqual(dspm["FoundationalOrAISpecific"], "AI-specific")
        self.assertEqual(dspm["ReadinessTreatment"], "EvidenceRequest")
        self.assertTrue(any("does not replace" in s for s in self.statements(dspm, REVIEWER)))

    def test_microsoft_requirement_versus_suggested_practice(self):
        audit = self.items["PUR-029"]
        self.assertEqual(audit["ReviewDecision"], "CorrectRecommendation")
        self.assertIn("MS-PURVIEW-AUDIT-RET-01", audit["MicrosoftBenchmarkIds"])
        self.assertTrue(any("CopilotInteraction" in s for s in self.statements(audit, MSDOC)))
        retention = self.items["PUR-022"]
        self.assertEqual(retention["ReviewDecision"], "CorrectRecommendation")
        self.assertTrue(any("not used to train" in s for s in self.statements(retention, MSDOC)))
        labels = self.items["PUR-033"]
        self.assertTrue(self.statements(labels, CUSTOMER))
        for bench in self.ctx.benchmarks["Benchmarks"]:
            self.assertTrue(bench["Url"].startswith("https://learn.microsoft.com/"))

    def test_customer_decision_versus_technical_failure(self):
        sharing = self.items["M365-122"]
        self.assertEqual(sharing["ReviewDecision"], "CustomerDecisionRequired")
        self.assertTrue(sharing["CustomerDecisionRequired"])
        consent = self.items["ENT-017"]
        self.assertEqual(consent["ReviewDecision"], "CorrectFinding")
        self.assertEqual(consent["ReadinessTreatment"], "TenantHardening")
        _, items, _, _ = self.review(self.build("legacy", consent_policies=("microsoft-user-default-legacy",)))
        self.assertEqual(items["ENT-017"]["ReviewDecision"], "ConfirmedWithQualification")
        self.assertEqual(items["ENT-017"]["ReadinessTreatment"], "ConfirmBeforePilot")

    def test_missing_evidence_is_not_failure(self):
        for item in self.items.values():
            script = item.get("ScriptConclusion") or {}
            if script.get("Disposition") == "Coverage":
                self.assertNotEqual(item["ReadinessTreatment"], "BlocksPilot", item["FindingId"])
                self.assertNotIn(item["ReviewDecision"], ("CorrectFinding",) if item["EvidenceSufficiency"] != "Contradictory" else ())
        gap = self.items["GAP-DATA-AUDIT"]
        check = [c for c in gap["ConsistencyChecks"] if c["CheckId"] == "C-GAP-NOT-FAIL"][0]
        self.assertEqual(check["Result"], "Pass")
        self.assertTrue(self.statements(gap, GAP), "A confirmed gap is stated as an evidence gap.")

    def test_contradictory_workbook_and_json(self):
        root = self.build("wb")
        path = next((root / "Deliverables").glob("AI Readiness*.xlsx"))
        book = load_workbook(path)
        ws = book["Action Plan"]
        for row in ws.iter_rows(min_row=2):
            if row[2].value == "ENT-017":
                row[0].value = "Low"
        book.save(path)
        pkg, items, _, _ = self.review(root)
        self.assertTrue(any(c["ConflictId"].startswith("CNF-WB-ENT-017") for c in pkg.conflicts))
        check = [c for c in items["ENT-017"]["ConsistencyChecks"] if c["CheckId"] == "C-OUTPUTS-AGREE"][0]
        self.assertEqual(check["Result"], "Issue")

    def test_risky_user_count_corrected(self):
        item = self.items["ENT-013"]
        self.assertEqual(item["ReviewDecision"], "CorrectFinding")
        self.assertTrue(self.statements(item, FACT))

    def test_contradictory_device_sources(self):
        item = self.items["ENT-002"]
        self.assertEqual(item["EvidenceSufficiency"], "Contradictory")
        self.assertIn("RVW-ENDPOINT-PROTECTION", self.items)

    def test_original_script_conclusion_is_preserved(self):
        for rid, item in self.items.items():
            if item["ItemType"] == "ScriptFinding":
                self.assertEqual(item["ScriptConclusion"]["FindingId"], rid)
                self.assertTrue(self.statements(item, SCRIPT))
            else:
                self.assertEqual(item["ReviewDecision"], "AddMissingFinding")

    def test_closure_is_never_automatic(self):
        for item in self.items.values():
            closure = item["ClosureEvidence"]
            self.assertIn("never automatic", closure["ClosureRule"])
            self.assertTrue(closure["AcceptableExceptionDocumentation"])
            self.assertEqual(item["CustomerActionStatus"], "Not recorded")


if __name__ == "__main__":
    unittest.main()
