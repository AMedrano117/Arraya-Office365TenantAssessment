"""Shared test helpers."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import shutil
import sys
import tempfile
import unittest

REPO = Path(__file__).resolve().parents[2]
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

from tools.assessment_review.engine import run_review          # noqa: E402
from tools.assessment_review.outputs import produce             # noqa: E402
from tools.assessment_review.package import import_package      # noqa: E402

from tests.assessment_review.fixtures import build_package      # noqa: E402

REVIEW_DATE = date(2026, 3, 3)


class PackageTestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="arv-"))

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def build(self, name="pkg", **kwargs):
        return build_package(self.tmp / name, **kwargs)

    def review(self, root):
        pkg = import_package(root, REVIEW_DATE)
        ctx, items, controls, readiness = run_review(pkg, REVIEW_DATE)
        return pkg, {i["FindingId"]: i for i in items}, readiness, ctx

    @staticmethod
    def check(pkg, check_id):
        return [c for c in pkg.checks if c.check_id == check_id]

    @staticmethod
    def statements(item, kind):
        return [s["Text"] for s in item["Statements"] if s["Type"] == kind]


def produce_package(root):
    return produce(root, REVIEW_DATE)
