"""Package import and validation tests (fictional packages only)."""
from __future__ import annotations

import unittest

from tests.assessment_review.fixtures import edit_json
from tests.assessment_review.helpers import REVIEW_DATE, PackageTestCase, import_package
from tools.assessment_review.package import classify_source
from tools.assessment_review.util import ReviewError


class PackageTests(PackageTestCase):
    def test_valid_package(self):
        root = self.build()
        pkg = import_package(root, REVIEW_DATE)
        self.assertTrue(pkg.customer_ready, [c.detail for c in pkg.blocking])
        self.assertEqual(pkg.identity["AssessmentId"], "AST-" + "a1" * 16)
        self.assertEqual(pkg.identity["RunType"], "Initial")
        self.assertEqual(self.check(pkg, "PKG-ID-OK")[0].status, "Pass")
        snapshot = [f for f in pkg.files if f["Role"] == "assessment-snapshot"][0]
        self.assertEqual(snapshot["HashStatus"], "Verified")
        self.assertEqual(self.check(pkg, "PKG-INPUT-HASHES")[0].status, "Pass")

    def test_missing_required_file(self):
        root = self.build()
        (root / "collection.json").unlink()
        with self.assertRaises(ReviewError):
            import_package(root, REVIEW_DATE)

    def test_missing_required_deliverable_blocks_publication(self):
        root = self.build()
        for f in (root / "Deliverables").glob("Technical Evidence*"):
            f.unlink()
        pkg = import_package(root, REVIEW_DATE)
        self.assertFalse(pkg.customer_ready)
        self.assertTrue(self.check(pkg, "PKG-FILE-MISSING"))

    def test_hash_mismatch(self):
        root = self.build()
        (root / "Collected" / "1" / "inputs" / "1" / "Permissioned_users_count_SharePoint_report.csv").write_text(
            "tampered\n", encoding="utf-8")
        pkg = import_package(root, REVIEW_DATE)
        self.assertFalse(pkg.customer_ready)
        self.assertTrue(any("hash mismatch" in c.detail.lower() for c in pkg.blocking))

    def test_snapshot_hash_mismatch(self):
        root = self.build()
        snap = next((root / "assessment-results").glob("*.json"))
        snap.write_text(snap.read_text(encoding="utf-8") + " ", encoding="utf-8")
        pkg = import_package(root, REVIEW_DATE)
        self.assertFalse(pkg.customer_ready)
        self.assertTrue(self.check(pkg, "PKG-HASH-MISMATCH"))

    def test_inconsistent_assessment_and_run_ids(self):
        root = self.build()
        edit_json(root / "collection.json", lambda d: d["identity"].update({"RunId": "RUN-" + "9" * 32}))
        pkg = import_package(root, REVIEW_DATE)
        mismatch = self.check(pkg, "PKG-ID-MISMATCH")
        self.assertTrue(mismatch and mismatch[0].severity == "Blocking")
        root2 = self.build("pkg2")
        edit_json(root2 / "run-seed.json", lambda d: d.update({"AssessmentId": "AST-" + "8" * 32}))
        self.assertFalse(import_package(root2, REVIEW_DATE).customer_ready)

    def test_partial_collection(self):
        root = self.build()
        pkg = import_package(root, REVIEW_DATE)
        self.assertEqual(pkg.sources["m365_copilot_interaction_audit"]["Classification"], "Partial")
        self.assertTrue(self.check(pkg, "PKG-SOURCE-PARTIAL"))
        self.assertTrue(pkg.customer_ready, "Partial optional sources qualify, not block, publication.")

    def test_failed_required_source(self):
        root = self.build(ca_state="failed")
        pkg, items, readiness, _ = self.review(root)
        self.assertEqual(pkg.sources["ca_policies"]["Classification"], "Failed")
        self.assertTrue(self.check(pkg, "PKG-SOURCE-FAILED"))
        self.assertEqual(items["ENT-020"]["ReviewDecision"], "DowngradeToEvidenceGap")
        self.assertEqual(items["ENT-020"]["ReadinessTreatment"], "EvidenceRequest")

    def test_complete_empty_source(self):
        root = self.build(managed_devices_state="empty")
        pkg, items, _, ctx = self.review(root)
        self.assertEqual(pkg.sources["managed_devices"]["Classification"], "CompleteEmpty")
        self.assertEqual(ctx.facts["intune"]["ManagedDevices"], 0)
        self.assertNotEqual(items["ENT-002"]["ReviewDecision"], "DowngradeToEvidenceGap")
        self.assertTrue(any("complete but empty" in items["ENT-002"]["CorrectedInterpretation"] for _ in [0]))

    def test_truncated_source(self):
        root = self.build(truncated_signins=True)
        pkg = import_package(root, REVIEW_DATE)
        self.assertEqual(pkg.sources["signin_logs"]["Classification"], "Partial")
        self.assertTrue(self.check(pkg, "PKG-SOURCE-TRUNCATED"))

    def test_malformed_and_sentinel_timestamps(self):
        def mutate(findings, evidence):
            findings[0]["ObservationDate"] = "27/10/2026"
            findings[1]["ObservationDate"] = "0001-01-01"
            evidence[2]["observed_at"] = "2099-13-45"
        root = self.build(extra_finding_mutation=mutate)
        pkg, items, _, _ = self.review(root)
        self.assertTrue(self.check(pkg, "PKG-DATE-MALFORMED"))
        self.assertTrue(self.check(pkg, "PKG-DATE-SENTINEL"))
        self.assertEqual(items["ENT-005"]["Evidence"]["ObservationDate"], "Not recorded")
        self.assertIsNone(items["ENT-020"]["Evidence"]["AgeDays"], "Sentinel dates must not drive freshness.")

    def test_broken_evidence_links(self):
        def mutate(findings, evidence):
            findings[0]["EvidenceId"] = "EV-ffffffffffffffff"
            findings[0]["RelatedEvidenceIds"] = ["EV-ffffffffffffffff"]
        root = self.build(extra_finding_mutation=mutate)
        pkg = import_package(root, REVIEW_DATE)
        links = self.check(pkg, "PKG-EVIDENCE-LINKS")[0]
        self.assertEqual((links.status, links.severity), ("Fail", "Blocking"))
        self.assertFalse(pkg.customer_ready)

    def test_broken_detail_link_is_qualified(self):
        root = self.build()
        html = next((root / "Deliverables").glob("AI Readiness*.html"))
        html.write_text(html.read_text(encoding="utf-8").replace("#action-1'>a", "#missing-anchor'>a"), encoding="utf-8")
        pkg = import_package(root, REVIEW_DATE)
        self.assertTrue(any(c["ConflictId"] == "CNF-DETAIL-LINKS" for c in pkg.conflicts))

    def test_customer_boundary(self):
        root = self.build()
        edit_json(root / "collection.json", lambda d: d.update({"tenant_name": "Contoso Other"}))
        pkg = import_package(root, REVIEW_DATE)
        self.assertTrue(self.check(pkg, "PKG-BOUNDARY-CUSTOMER"))
        self.assertFalse(pkg.customer_ready)

    def test_source_classification_vocabulary(self):
        cases = {("available", True, 5): "Complete", ("available", True, 0): "CompleteEmpty",
                 ("empty", True, 0): "CompleteEmpty", ("partial", False, 0): "Partial",
                 ("failed", False, 0): "Failed", ("unavailable", False, 0): "Unavailable",
                 ("not_requested", False, 0): "NotRequested", ("not_applicable", None, None): "NotApplicable",
                 ("deferred", False, None): "Unknown"}
        for (state, complete, records), expected in cases.items():
            self.assertEqual(classify_source(state, complete, records), expected, state)

    def test_package_is_not_modified(self):
        root = self.build()
        before = {p: p.read_bytes() for p in root.rglob("*") if p.is_file()}
        self.review(root)
        after = {p: p.read_bytes() for p in root.rglob("*") if p.is_file()}
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main()
