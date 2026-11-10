"""Privacy and deliverable tests on fictional packages."""
from __future__ import annotations

import json
import re
import unittest
from io import BytesIO

from openpyxl import load_workbook

from tests.assessment_review.fixtures import CUSTOMERS, SECRET_VALUE
from tests.assessment_review.helpers import PackageTestCase, produce_package
from tools.assessment_review.outputs import (MANIFEST, RESULTS, SUMMARY_SHEETS, build_manifest, docx_text,
                                             write_outputs)
from tools.assessment_review.render import (ACTION_HEADERS, BRIEF, DETAIL_HEADERS, EXEC, PLAN, TECH, TRACE_SHEET,
                                            top_recommendations)

LISTS = "list sheets"


def sheet_rows(outputs, name):
    return list(load_workbook(BytesIO(outputs[PLAN]))[name].iter_rows(values_only=True))


def all_text(outputs):
    """Word documents, workbook summary sheets (PLAN), workbook list sheets (LISTS) and review-results.json."""
    texts = {name: docx_text(outputs[name]) for name in (EXEC, TECH, BRIEF)}
    book = load_workbook(BytesIO(outputs[PLAN]))

    def joined(sheets):
        return " ".join(str(c) for ws in sheets for r in ws.iter_rows(values_only=True) for c in r if c)
    texts[PLAN] = joined(ws for ws in book.worksheets if ws.title in SUMMARY_SHEETS)
    texts[LISTS] = joined(ws for ws in book.worksheets if ws.title not in SUMMARY_SHEETS)
    texts[RESULTS] = outputs[RESULTS].decode("utf-8")
    return texts


class PrivacyTests(PackageTestCase):
    def setUp(self):
        super().setUp()

        def inject(findings, evidence):
            c = CUSTOMERS["northwind"]
            findings[3]["Observation"] += (" Contact Jordan Avery at jordan.avery@northwindtraders.example or "
                                           "+1 (425) 555-0100. Payload {\"@odata.context\": \"x\"} "
                                           "client_secret=abc123XYZsecretvalue from 203.0.113.7 for "
                                           "11111111-2222-4333-8444-555555555555.")
            findings[10]["Recommendation"] += " Assigned to Priya Natarajan (Global Administrator)."
        self.root = self.build(extra_finding_mutation=inject)
        self.outputs, self.results, self.ctx, self.pkg, self.problems = produce_package(self.root)
        self.texts = all_text(self.outputs)

    def test_generation_passes_built_in_validation(self):
        self.assertEqual(self.problems, [])

    def test_no_personal_or_secret_values(self):
        c = CUSTOMERS["northwind"]
        forbidden = c["people"] + ["northwindtraders.example", "555-0100", "@odata", "abc123XYZsecretvalue",
                                   "203.0.113.7", c["tenant"], c["incident"], "fin-laptop-07"]
        for name, body in self.texts.items():
            if name == LISTS:
                continue
            for value in forbidden:
                self.assertNotIn(value.lower(), body.lower(), f"{value} leaked into {name}")
            self.assertIsNone(re.search(r"[\w.]+@[\w.]+\.\w+", body), name)

    def test_list_sheets_never_carry_secrets_or_payloads(self):
        lists = self.texts[LISTS]
        for value in (SECRET_VALUE, "abc123XYZsecretvalue", "@odata", "555-0100", "203.0.113.7"):
            self.assertNotIn(value, lists)
        for name, body in self.texts.items():
            self.assertNotIn(SECRET_VALUE, body, name)
        admins = sheet_rows(self.outputs, "Admin accounts")
        app_row = next(r for r in admins if "Directory Sync" in str(r[0]))
        self.assertIn("1 secret, latest expiry 2027-01-01", " ".join(map(str, app_row)))

    def test_names_appear_only_in_list_sheets(self):
        c = CUSTOMERS["northwind"]
        for name, body in self.texts.items():
            if name != LISTS:
                self.assertNotIn(c["incident"], body, name)
                self.assertNotIn("Priya", body, name)
        self.assertIn(c["incident"], self.texts[LISTS])
        self.assertIn("priya.natarajan@northwindtraders.example", self.texts[LISTS])

    def test_restricted_evidence_stays_restricted(self):
        restricted = {r["EvidenceId"] for i in self.results["Items"] for r in i["EvidenceReferences"]
                      if r["Classification"] == "Restricted"}
        self.assertIn("EV-0000000000000001", restricted)
        for name in (EXEC, BRIEF, PLAN, LISTS):
            for ev in restricted:
                self.assertNotIn(ev, self.texts[name], name)
        self.assertIn("EV-0000000000000001", self.texts[RESULTS])
        self.assertIn("[Restricted]", self.texts[TECH])


class DeliverableTests(PackageTestCase):
    def setUp(self):
        super().setUp()
        self.root = self.build()
        self.outputs, self.results, self.ctx, self.pkg, self.problems = produce_package(self.root)

    def test_all_documents_open_and_validate(self):
        self.assertEqual(self.problems, [])
        self.assertEqual(set(self.outputs), {EXEC, TECH, PLAN, BRIEF, RESULTS})
        texts = all_text(self.outputs)
        for name in (EXEC, TECH, BRIEF):
            self.assertIn(self.results["Readiness"]["ReviewDecision"], texts[name])

    def test_ids_and_counts_reconcile(self):
        book = load_workbook(BytesIO(self.outputs[PLAN]))
        details = list(book[TRACE_SHEET].iter_rows(values_only=True))
        ids = sorted(r[DETAIL_HEADERS.index("Ref")] for r in details[1:])
        self.assertEqual(ids, sorted(i["FindingId"] for i in self.results["Items"]))
        self.assertEqual(len(ids), self.results["Summary"]["ItemCount"])
        script_ids = {f["RecommendationId"] for f in self.pkg.snapshot["customer_findings"]}
        self.assertTrue(script_ids <= set(ids), "Every script finding is reviewed.")
        plan = list(book["Action Plan"].iter_rows(values_only=True))
        numbers = [r[0] for r in plan[1:]]
        self.assertEqual(numbers, list(range(1, self.results["Summary"]["ActionCount"] + 1)))
        quiet = {i["FindingId"] for i in self.results["Items"] if i["ReadinessTreatment"] == "NoAction"
                 or i["ReviewDecision"] == "DuplicateOrOverlapping"}
        self.assertFalse(quiet & {r[ACTION_HEADERS.index("Ref")] for r in plan[1:]}, "Plan lists only actions.")

    def test_documents_agree_on_rollout_items(self):
        texts = all_text(self.outputs)
        gated = [i for i in self.results["Items"] if i["ActionNumber"]
                 and i["ReadinessTreatment"] in ("BlocksPilot", "ConfirmBeforePilot")]
        self.assertTrue(gated)
        for item in gated:
            self.assertIn(item["FindingId"], texts[TECH])
            for name in (BRIEF, PLAN):
                self.assertIn(item["Plain"]["Headline"], texts[name], name)
        for item in top_recommendations(self.results):
            self.assertIn(item["Plain"]["Headline"], texts[EXEC])

    def test_every_action_has_plain_wording(self):
        for item in self.results["Items"]:
            plain = item["Plain"]
            for key in ("Headline", "Issue", "HowToFix", "DoneWhen"):
                self.assertTrue(plain[key], f"{item['FindingId']} lacks {key}")
            for value in plain.values():
                self.assertNotRegex(value, r"ConfirmBeforePilot|TenantHardening|EvidenceRequest|\bscript\b",
                                    item["FindingId"])

    def test_workbook_is_filterable(self):
        book = load_workbook(BytesIO(self.outputs[PLAN]))
        ws = book["Action Plan"]
        self.assertIn("ActionPlan", ws.tables)
        self.assertEqual(tuple(c.value for c in ws[1]), ACTION_HEADERS)
        self.assertLessEqual(len(ACTION_HEADERS), 11, "The action plan stays readable.")
        self.assertEqual(ws.tables["ActionPlan"].ref.split(":")[0], "A1")
        self.assertTrue(ws.data_validations.dataValidation, "Status offers a choice list.")
        self.assertEqual(book.sheetnames[:3], ["Start Here", "Action Plan", "Decisions"])
        self.assertEqual(book.sheetnames[-2:], [TRACE_SHEET, "Microsoft guidance"])
        for name in book.sheetnames[3:-2]:
            lists = book[name]
            self.assertTrue(lists.tables, f"{name} is a filterable table")
            headers = [c.value for c in lists[1]]
            self.assertIn(headers[-3], ("Action needed", "Options to fix"), name)
            self.assertEqual(headers[-2:], ["Actions", "Status"], name)

    def test_action_plan_points_to_named_lists(self):
        plan = sheet_rows(self.outputs, "Action Plan")
        col = ACTION_HEADERS.index("Affected items")
        by_ref = {r[ACTION_HEADERS.index("Ref")]: r for r in plan[1:]}
        expected = {"ENT-005": "Users without MFA sheet", "DEF-013": "Security incidents sheet",
                    "DEX-001": "Overshared sites sheet", "ENT-013": "Risky accounts sheet",
                    "RVW-GLOBAL-ADMIN-COUNT": "Admin accounts sheet", "ENT-002": "Devices sheet"}
        for ref, sheet in expected.items():
            self.assertIn(sheet, by_ref[ref][col], ref)
        incidents = sheet_rows(self.outputs, "Security incidents")
        self.assertEqual(len(incidents) - 1, 3)
        self.assertTrue(all(str(r[incidents[0].index("Link")]).startswith("https://security.microsoft.com/")
                            for r in incidents[1:]))
        sites = sheet_rows(self.outputs, "Overshared sites")
        where = sites[0].index("Where the broad access is")
        self.assertIn("Everyone except external users via Finance Members", sites[1][where])
        self.assertIn("Restricted Content Discovery", sites[1][sites[0].index("Options to fix")])
        apps = sheet_rows(self.outputs, "Apps with high access")
        flagged = {r[0]: r for r in apps[1:]}
        self.assertIn("Unverified Helper Tool", flagged)
        self.assertIn("Mail.ReadWrite", flagged["Northwind Traders Mail Archiver"][apps[0].index(
            "High-impact permissions")])
        inventory = {r[0]: r for r in sheet_rows(self.outputs, "Sites")[1:]}
        self.assertIn("archive", inventory["Northwind Traders Archive"][-3].lower())
        self.assertIn("Anyone links", inventory["Northwind Traders Archive"][-3])

    def test_engineering_qualification_preserved_when_blocked(self):
        def mutate(findings, evidence):
            findings[0]["EvidenceId"] = "EV-ffffffffffffffff"
        root = self.build("blocked", extra_finding_mutation=mutate)
        outputs, results, _, _, problems = produce_package(root)
        self.assertEqual(problems, [])
        self.assertFalse(results["Publication"]["CustomerReady"])
        for name, body in all_text(outputs).items():
            if name not in (RESULTS, LISTS):
                self.assertIn("ENGINEERING REVIEW ONLY", body, name)

    def test_manifest_hashes_match(self):
        manifest = build_manifest(self.outputs, self.results, self.pkg)
        out = write_outputs(self.outputs, manifest, self.tmp / "out")
        stored = json.loads((out / MANIFEST).read_text(encoding="utf-8"))
        self.assertEqual({d["File"] for d in stored["Deliverables"]}, {EXEC, TECH, PLAN, BRIEF, RESULTS})
        for d in stored["Deliverables"]:
            from hashlib import sha256
            self.assertEqual(sha256((out / d["File"]).read_bytes()).hexdigest(), d["SHA256"])

    def test_repeated_generation_is_deterministic(self):
        outputs2, _, _, _, _ = produce_package(self.root)
        for name in self.outputs:
            self.assertEqual(self.outputs[name], outputs2[name], name)

    def test_no_cross_customer_contamination(self):
        other = self.build("tailspin", customer="tailspin", pilot_scope=True,
                           consent_policies=("microsoft-user-default-legacy",))
        outputs_b, results_b, _, _, problems_b = produce_package(other)
        self.assertEqual(problems_b, [], "Second customer runs without code changes.")
        a, b = CUSTOMERS["northwind"], CUSTOMERS["tailspin"]
        texts_a, texts_b = all_text(self.outputs), all_text(outputs_b)
        for name in texts_b:
            for value in (a["name"], a["assessment"], a["run"], a["policy"]):
                self.assertNotIn(value, texts_b[name], f"{value} in tailspin {name}")
            for value in (b["name"], b["assessment"], b["run"]):
                self.assertNotIn(value, texts_a[name], f"{value} in northwind {name}")
        self.assertEqual(results_b["AssessmentId"], b["assessment"])

    def test_no_implementation_commitment(self):
        for name, body in all_text(self.outputs).items():
            self.assertIsNone(re.search(r"(?i)arraya (will|shall) (implement|remediate|deliver)", body), name)
            if name not in (RESULTS, LISTS):
                self.assertIn("does not commit Arraya", body, name)

    def test_review_results_required_fields(self):
        data = json.loads(self.outputs[RESULTS])
        required = ("AssessmentId", "RunId", "ReviewVersion", "BenchmarkVersion", "ControlId", "FindingId",
                    "ScriptConclusion", "ReviewDecision", "CorrectedInterpretation", "MicrosoftBenchmarkIds",
                    "EvidenceReferences", "EvidenceSufficiency", "ReadinessTreatment", "FoundationalOrAISpecific",
                    "RecommendedAction", "ClosureEvidence", "Limitations")
        for item in data["Items"]:
            for key in required:
                self.assertIn(key, item, f"{item['FindingId']} lacks {key}")


if __name__ == "__main__":
    unittest.main()
