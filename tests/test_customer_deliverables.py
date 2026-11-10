"""Fictional fixtures only; never load customer inputs in automated tests."""
from copy import deepcopy
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
import customer_deliverables as m
from docx import Document
from openpyxl import load_workbook


def fixture():
    actions = []
    for i in range(1, 21):
        identifier = f"FICT-{i:03}"
        actions.append({
            "RecordId": identifier, "FindingId": identifier, "CustomerFindingId": identifier,
            "PersistentFindingId": f"PF-{i:03}", "SourceRecommendationId": identifier,
            "ActionId": f"PA-{i:03}", "Title": f"Fictional assessment condition {i}",
            "Priority": "High" if i <= 5 else "Medium",
            "PilotImpact": "Blocks pilot" if i <= 5 else "Confirm before pilot",
            "CurrentConclusion": f"Fictional observed condition {i}.",
            "Qualification": "Configuration alone does not demonstrate effective operation.",
            "RecommendedAction": f"Review the fictional condition {i}.",
            "ResponsibleRole": "Security operations role", "ClosureEvidence": f"Dated scoped proof for condition {i}.",
            "TechnicalStatus": "Action Required", "EvidenceSufficiency": "supported" if i <= 9 else "incomplete",
            "LifecycleState": "Open", "GovernanceTreatment": "No treatment recorded",
            "CustomerActionStatus": None, "ActionType": "Remediation" if i <= 5 else "Confirmation" if i <= 9 else "Evidence",
            "EvidenceLevel": "configuration", "ObservationDate": "2026-01-02",
            "restricted_references": ["ARE-" + "a" * 64],
            "closure_evidence_references": ["ARE-" + "a" * 64],
            "technicalOnly": {"displayName": "Casey Fictional", "email": "fictional@example.invalid"},
        })
    control = {"control_id": "FICT.CONTROL", "domain_id": "identity", "status": "Observed",
               "title": "Fictional configuration observation", "configuration_result": "established",
               "operational_result": "not_verified", "statusExplanation": None}
    review = {
        "projection_schema": "agent-customer-review.schema.json", "projection_version": "1.2.0",
        "kind": "customer-review", "source_sha256": "b" * 64,
        "identity": {"AssessmentId": "AST-fictional", "RunId": "RUN-fictional",
                     "RunType": "Initial", "BaselineRunId": None, "EvaluatedAt": "2026-01-02"},
        "decision": "Not ready for pilot", "rationale": "Five recorded blockers remain.",
        "methodology_version": "fictional-1", "evaluation_date": "2026-01-02",
        "evidence_sufficiency": "Incomplete",
        "source_publication": {"output_qualification": m.QUALIFICATION, "customer_publication_eligible": False},
        "exporter_provenance": {"output_qualification": m.QUALIFICATION},
        "customerPublicationEligible": False, "demoUse": "InternalReadOnlyEvaluation",
        "actions": actions, "findings": deepcopy(actions), "controls": [control],
        "blockers_ids": [a["FindingId"] for a in actions[:5]],
        "pilot_conditions_ids": [a["FindingId"] for a in actions[5:]],
        "strengths_ids": [], "evidence_gaps": [a["FindingId"] for a in actions[9:]],
        "counts": {"actions": 20, "confirmation": 4, "remediation": 5, "evidence_gaps": 11},
        "controlCounts": {"totalControls": 1, "requiredDeploymentControls": 1,
                         "answeredRequiredControls": 1, "unresolvedRequiredControls": 0},
        "reconciliation": {"actions": 20, "blockers": 5, "findings": 20,
                           "action_ids": [a["FindingId"] for a in actions],
                           "finding_ids": [a["FindingId"] for a in actions]},
        "approved_documents": [],
        "lifecycle": {"recorded": True, "Enabled": True, "Outcome": "Compared", "BaselineRunId": None,
                      "records": [{"EntityId": "PF-001", "State": "Open", "RemainsOpen": True,
                                   "RemainingClosureRequirement": "Dated scoped proof for condition 1."}]},
        "governance": {"recorded": True, "AsOf": None,
                       "records": [{"DecisionId": "GOV-FICT-001", "TargetEntityId": "PF-001",
                                    "WorkflowState": "Pending", "GovernanceState": "NotApproved",
                                    "AuthorityRole": "Security governance role"}]},
        "limitations": ["Fictional evidence is incomplete.", "No records returned is not proof of absence."],
        "integrity_limitations": [],
    }
    index = {"kind": "evidence-index", "projection_version": "1.2.0",
             "AssessmentId": "AST-fictional", "RunId": "RUN-fictional", "source_sha256": "b" * 64,
             "entries": [{"technicalOnly": {"raw": "RESTRICTED_INDEX_SENTINEL",
                            "displayName": "Casey Fictional", "email": "fictional@example.invalid",
                            "objectId": "12345678-1234-1234-1234-123456789012"}}]}
    return review, index


def seal(path, review, index, markdown="# Fictional package\n"):
    path.mkdir(parents=True, exist_ok=True)
    r = deepcopy(review)
    raw_index, raw_md = m.canonical(index), markdown.encode()
    body = {k: v for k, v in r.items() if k not in ("output_hashes", "hash_contract")}
    r["output_hashes"] = {m.REVIEW: m.digest(m.canonical(body)),
                          m.KNOWLEDGE: m.digest(raw_md), m.INDEX: m.digest(raw_index)}
    r["hash_contract"] = "Review self-hash excludes output_hashes and hash_contract."
    payloads = {m.REVIEW: m.canonical(r), m.KNOWLEDGE: raw_md, m.INDEX: raw_index}
    manifest = {"AssessmentId": r["identity"]["AssessmentId"], "RunId": r["identity"]["RunId"],
                "ProjectionVersion": "1.2.0", "SourceHash": r["source_sha256"],
                "PublicationQualification": m.QUALIFICATION, "customerPublicationEligible": False,
                "IntendedUse": "InternalReadOnlyEvaluation", "ApprovedLinksPresent": False,
                "GeneratedFiles": [{"FileName": name, "SHA256": m.digest(data),
                    "AccessClassification": "RestrictedTechnical" if name == m.INDEX else "CustomerReview"}
                    for name, data in sorted(payloads.items())]}
    manifest["ManifestHash"] = m.digest(m.canonical(manifest))
    payloads[m.MANIFEST] = m.canonical(manifest)
    for name, data in payloads.items():
        (path / name).write_bytes(data)
    return r


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.inputs, self.outputs = self.root / "input", self.root / "generated"
        self.review, self.index = fixture()
        self.seal()

    def tearDown(self):
        self.tmp.cleanup()

    def seal(self, markdown="# Fictional package\n"):
        return seal(self.inputs, self.review, self.index, markdown)

    def test_three_files_exact_action_fidelity_and_input_preservation(self):
        before = {p.name: p.read_bytes() for p in self.inputs.iterdir()}
        result = m.generate(self.inputs, self.outputs)
        output = Path(result["output"])
        self.assertEqual({p.name for p in output.iterdir()}, set(m.OUTPUTS))
        self.assertEqual((result["blockers"], result["actions"]), (5, 20))
        self.assertEqual(before, {p.name: p.read_bytes() for p in self.inputs.iterdir()})
        wb = load_workbook(output / m.OUTPUTS[1])
        self.assertEqual(wb["Customer Actions"].max_row, 21)
        self.assertEqual(list(wb["Customer Actions"].values)[1][0], "FICT-001")
        self.assertEqual(list(wb["Source States"].values)[1][2:6], ("FICT-001", "PF-001", "FICT-001", "PA-001"))
        self.assertIn(("governance record 1: GovernanceState", "NotApproved"),
                      list(wb["Assessment Context"].values))
        wb.close()

    def test_repeated_generation_is_identical_and_idempotent(self):
        first = m.generate(self.inputs, self.outputs)
        second = m.generate(self.inputs, self.outputs)
        third = m.generate(self.inputs, self.root / "another")
        self.assertEqual(first["files"], second["files"])
        self.assertEqual(first["files"], third["files"])

    def test_hash_mismatch_in_each_export_blocks_generation(self):
        for name in (m.REVIEW, m.KNOWLEDGE, m.INDEX):
            with self.subTest(name=name):
                self.seal()
                with (self.inputs / name).open("ab") as f:
                    f.write(b" ")
                with self.assertRaisesRegex(ValueError, "hash mismatch"):
                    m.generate(self.inputs, self.outputs)
                self.assertFalse(self.outputs.exists())

    def test_manifest_self_hash_mismatch(self):
        path = self.inputs / m.MANIFEST
        data = json.loads(path.read_bytes())
        data["RunId"] = "RUN-tampered"
        path.write_bytes(m.canonical(data))
        with self.assertRaisesRegex(ValueError, "self-hash"):
            m.validate_package(self.inputs)

    def test_missing_required_file_blocks(self):
        (self.inputs / m.INDEX).unlink()
        with self.assertRaises(FileNotFoundError):
            m.generate(self.inputs, self.outputs)

    def test_version_mismatch(self):
        self.index["projection_version"] = "1.0.0"
        self.seal()
        with self.assertRaisesRegex(ValueError, "version"):
            m.validate_package(self.inputs)

    def test_identity_mismatch(self):
        self.index["RunId"] = "RUN-other"
        self.seal()
        with self.assertRaisesRegex(ValueError, "RunId mismatch"):
            m.validate_package(self.inputs)

    def test_path_traversal_blocked(self):
        self.review["identity"]["AssessmentId"] = "../escape"
        self.seal()
        with self.assertRaisesRegex(ValueError, "Unsafe"):
            m.validate_package(self.inputs)

    def test_counts_not_repaired(self):
        self.review["counts"]["actions"] = 19
        self.seal()
        with self.assertRaisesRegex(ValueError, "twenty"):
            m.validate_package(self.inputs)

    def test_missing_blocker_not_invented(self):
        self.review["blockers_ids"][0] = "FICT-ABSENT"
        self.seal()
        with self.assertRaisesRegex(ValueError, "blocker ID"):
            m.validate_package(self.inputs)

    def test_action_priority_conflict_not_resolved(self):
        self.review["actions"][0]["Priority"] = "Low"
        self.seal()
        with self.assertRaisesRegex(ValueError, "disagreement"):
            m.validate_package(self.inputs)

    def test_duplicate_finding_id(self):
        self.review["findings"][1]["FindingId"] = "FICT-001"
        self.seal()
        with self.assertRaisesRegex(ValueError, "duplicate finding"):
            m.validate_package(self.inputs)

    def test_private_index_and_references_never_rendered(self):
        outputs = m.render(self.seal())
        for data in outputs.values():
            with zipfile.ZipFile(BytesIO(data)) as archive:
                expanded = b"\n".join(archive.read(n) for n in archive.namelist())
                for forbidden in (b"Casey Fictional", b"fictional@example.invalid", b"ARE-",
                                  b"RESTRICTED_INDEX_SENTINEL", b"12345678-1234"):
                    self.assertNotIn(forbidden, expanded)

    def test_private_text_in_allowed_field_blocks_before_write(self):
        for group in ("actions", "findings"):
            self.review[group][0]["ResponsibleRole"] = "fictional@example.invalid"
        self.seal()
        with self.assertRaisesRegex(ValueError, "Privacy"):
            m.generate(self.inputs, self.outputs)
        self.assertFalse(self.outputs.exists())

    def test_markdown_cannot_override_json(self):
        self.seal("# Ready for pilot\nZero blockers. No actions. MarkdownOverrideSentinel\n")
        result = m.generate(self.inputs, self.outputs)
        doc = Document(Path(result["output"]) / m.OUTPUTS[0])
        text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("Not ready for pilot", text)
        self.assertNotIn("MarkdownOverrideSentinel", text)

    def test_formula_text_stays_literal(self):
        for group in ("actions", "findings"):
            self.review[group][0]["Title"] = "=1+1"
        self.seal()
        result = m.generate(self.inputs, self.outputs)
        wb = load_workbook(Path(result["output"]) / m.OUTPUTS[1])
        self.assertEqual(wb["Customer Actions"]["B2"].value, "=1+1")
        self.assertEqual(wb["Customer Actions"]["B2"].data_type, "s")
        wb.close()

    def test_different_existing_output_not_overwritten(self):
        result = m.generate(self.inputs, self.outputs)
        path = Path(result["output"]) / m.OUTPUTS[0]
        path.write_bytes(b"Operator revision")
        with self.assertRaisesRegex(ValueError, "no overwrite"):
            m.generate(self.inputs, self.outputs)
        self.assertEqual(path.read_bytes(), b"Operator revision")

    def test_engineering_qualification_cannot_be_removed(self):
        self.review["customerPublicationEligible"] = True
        self.seal()
        with self.assertRaisesRegex(ValueError, "qualification"):
            m.generate(self.inputs, self.outputs)

    def test_detects_input_change_after_validation(self):
        _, hashes = m.validate_package(self.inputs)
        (self.inputs / m.KNOWLEDGE).write_text("changed")
        with self.assertRaisesRegex(ValueError, "Inputs changed"):
            m.assert_inputs_unchanged(self.inputs, hashes)


if __name__ == "__main__":
    unittest.main()

