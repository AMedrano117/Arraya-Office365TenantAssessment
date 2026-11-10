"""Expert validation scenarios against a real completed package (opt-in).

Set ASSESSMENT_REVIEW_VALIDATION_PACKAGE to a completed package folder to run.
The package is read-only; no customer content is stored in this repository.
Expectations were set by comparing each scenario with the retained evidence and
the Microsoft Learn benchmarks verified on the registry date.
"""
from __future__ import annotations

from datetime import date
import os
import unittest

from tests.assessment_review.helpers import produce  # noqa: F401 (path setup)
from tools.assessment_review.outputs import produce as produce_outputs

PACKAGE = os.environ.get("ASSESSMENT_REVIEW_VALIDATION_PACKAGE")
REVIEW_DATE = date.fromisoformat(os.environ.get("ASSESSMENT_REVIEW_VALIDATION_DATE", "2026-10-28"))

# Scenario: finding id -> (expected decision, expected treatment)
EXPECTED = {
    "MFA": {"ENT-005": ("ConfirmedWithQualification", "ConfirmBeforePilot")},
    "Conditional Access": {"ENT-020": ("CorrectRecommendation", "BlocksPilot")},
    "Site ownership": {"GAP-CONTENT-OWNERSHIP": ("CorrectFinding", "ConfirmBeforePilot")},
    "Adoption (out of scope)": {"GAP-ADOPTION-BASELINE": ("NotApplicable", "NoAction")},
    "Privileged access": {"REC-FB1EB16E4C": ("ConfirmedWithQualification", "ConfirmBeforePilot"),
                          "RVW-GLOBAL-ADMIN-COUNT": ("AddMissingFinding", "TenantHardening")},
    "Intune and endpoint posture": {"ENT-002": ("ConfirmedWithQualification", "ConfirmBeforePilot"),
                                    "RVW-ENDPOINT-PROTECTION": ("AddMissingFinding", "ConfirmBeforePilot")},
    "Defender incidents": {"DEF-013": ("ConfirmedWithQualification", "ConfirmBeforePilot")},
    "SharePoint permissions": {"DEX-001": ("ConfirmedWithQualification", "ConfirmBeforePilot"),
                               "M365-127": ("Confirmed", "ConfirmBeforePilot")},
    "Sensitivity labels": {"GAP-DATA-LABELS": ("CorrectFinding", "ConfirmBeforePilot"),
                           "PUR-033": ("CorrectRecommendation", "LongerTermImprovement")},
    "DLP": {"PUR-040": ("ConfirmedWithQualification", "NoAction"),
            "GAP-DATA-DLP": ("ConfirmedWithQualification", "ConfirmBeforePilot")},
    "Audit and retention": {"PUR-029": ("CorrectRecommendation", "LongerTermImprovement"),
                            "PUR-022": ("CorrectRecommendation", "LongerTermImprovement"),
                            "GAP-DATA-AUDIT": ("ConfirmedWithQualification", "ConfirmBeforePilot")},
    "Application consent": {"ENT-017": ("CorrectFinding", "TenantHardening"),
                            "ENT-018": ("ConfirmedWithQualification", "TenantHardening")},
    "Copilot licensing": {"GAP-LICENSE-ASSIGNMENT": ("CorrectFinding", "ConfirmBeforePilot"),
                          "M365-129": ("ConfirmedWithQualification", "NoAction")},
    "Shadow AI": {"PUR-016": ("CorrectFinding", "LongerTermImprovement"),
                  "RVW-SHADOW-AI-VISIBILITY": ("AddMissingFinding", "EvidenceRequest")},
}


@unittest.skipUnless(PACKAGE, "Set ASSESSMENT_REVIEW_VALIDATION_PACKAGE to run expert validation.")
class ExpertValidation(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.outputs, cls.results, cls.ctx, cls.pkg, cls.problems = produce_outputs(PACKAGE, REVIEW_DATE)
        cls.items = {i["FindingId"]: i for i in cls.results["Items"]}

    def test_outputs_validate(self):
        self.assertEqual(self.problems, [])

    def test_scenarios(self):
        for scenario, expected in EXPECTED.items():
            for rid, (decision, treatment) in expected.items():
                with self.subTest(scenario=scenario, finding=rid):
                    item = self.items[rid]
                    self.assertEqual((item["ReviewDecision"], item["ReadinessTreatment"]), (decision, treatment))
                    self.assertTrue(item["MicrosoftBenchmarkIds"])
                    self.assertTrue(item["Statements"])

    def test_every_script_finding_reviewed(self):
        script = {f["RecommendationId"] for f in self.pkg.snapshot["customer_findings"]}
        self.assertTrue(script <= set(self.items))

    def test_no_conclusion_changes_without_decision(self):
        for item in self.results["Items"]:
            if item["ItemType"] == "ScriptFinding" and item["ReviewDecision"] not in ("Confirmed", "DuplicateOrOverlapping"):
                if item["ConclusionChanges"]:
                    self.assertTrue(item["DecisionRationale"])


if __name__ == "__main__":
    unittest.main()
