"""Command line entry point.

    .venv\\Scripts\\python.exe -B -m tools.assessment_review --package <dir> --review-date YYYY-MM-DD
"""
from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path
import sys

from .outputs import MANIFEST, build_manifest, produce, verify_input_unchanged, write_outputs
from .util import ReviewError

REPO_ROOT = Path(__file__).resolve().parents[2]


def main(argv=None):
    parser = argparse.ArgumentParser(description="Review one completed assessment package.")
    parser.add_argument("--package", required=True, help="Completed assessment package folder (read-only).")
    parser.add_argument("--review-date", help="Review date (YYYY-MM-DD). Defaults to the package evaluation date.")
    parser.add_argument("--output-root", default=str(REPO_ROOT / "generated" / "assessment-review"))
    parser.add_argument("--allow-engineering-only", action="store_true",
                        help="Write outputs even when customer-ready publication is stopped.")
    args = parser.parse_args(argv)
    review_date = date.fromisoformat(args.review_date) if args.review_date else None
    try:
        if review_date is None:
            from .package import import_package
            review_date = date.fromisoformat(import_package(args.package).identity["EvaluationDate"])
        outputs, results, ctx, pkg, problems = produce(args.package, review_date)
        if problems:
            print("VALIDATION FAILED")
            for p in problems:
                print(" -", p)
            return 2
        if not results["Publication"]["CustomerReady"] and not args.allow_engineering_only:
            print("Customer-ready publication is stopped:")
            for reason in results["Publication"]["Reasons"]:
                print(" -", reason)
            print("Rerun with --allow-engineering-only to write qualified engineering outputs.")
            return 3
        manifest = build_manifest(outputs, results, pkg)
        out_dir = Path(args.output_root) / results["AssessmentId"] / results["RunId"]
        write_outputs(outputs, manifest, out_dir)
        verify_input_unchanged(pkg)
    except ReviewError as exc:
        print(f"REVIEW STOPPED: {exc}")
        return 1
    print(f"Readiness decision: {results['Readiness']['ReviewDecision']}")
    print(f"Customer-ready: {results['Publication']['CustomerReady']}")
    print(f"Items: {results['Summary']['ItemCount']} ({results['Summary']['AddedFindings']} added)")
    print(f"Output: {out_dir}")
    for d in manifest["Deliverables"]:
        print(f"  {d['SHA256']}  {d['File']}")
    print(f"  manifest: {MANIFEST}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
