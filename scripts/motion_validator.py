#!/usr/bin/env python3
"""Compatibility interface. QA inspects; only PMGOrchestrator repairs/certifies.

Legacy public names and tuple results are retained. Static inspection never
certifies delivery and this module contains no independent repair loop.
"""
import argparse
import os
import sys
from typing import Dict, List, Optional, Tuple

if __package__ in (None, ""):
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class DeckReviewReport:
    """Legacy view of the canonical structured audit, not a second validator."""
    def __init__(self, pptx_path: str, audit: Optional[dict] = None):
        self.pptx_path = pptx_path
        self.audit = audit or {}
        self.total_slides = self.audit.get("total_slides", 0)
        self.slide_transitions = {s["slide_number"]: s["transition_type"] for s in self.audit.get("slides", [])}
        self.missing_transitions = self.audit.get("missing_transitions", [])
        self.traveling_anchors = self.audit.get("traveling_anchors", {})
        self.discontinuous_anchors = self.audit.get("discontinuous_anchors", {})
        self.word_counts = {s["slide_number"]: s["word_count"] for s in self.audit.get("slides", [])}
        self.overflow_slides = [s for s, words in self.word_counts.items() if words > 35]
        self.durations = {s["slide_number"]: s["transition_duration_ms"] for s in self.audit.get("slides", [])}
        self.is_valid = self.audit.get("is_valid", False)
        self.errors = [d["description"] for d in self.audit.get("defects", []) if d["severity"] in ("critical", "major")]
        self.warnings = [d["description"] for d in self.audit.get("defects", []) if d["severity"] not in ("critical", "major")]
        self.status = self.audit.get("status", "NOT_VERIFIED")
        self.certified = False

    def print_summary(self):
        print(f"PMG audit: {self.total_slides} slides; static_valid={self.is_valid}; status={self.status}")
        for message in self.errors + self.warnings:
            print(message)
        if not self.certified:
            print("Read-only inspection; real PowerPoint playback certification is NOT_VERIFIED.")


def validate_deck(pptx_path: str) -> DeckReviewReport:
    from pmg.validator import StaticValidator
    return DeckReviewReport(pptx_path, StaticValidator().validate(pptx_path))


def repair_xml_transitions(pptx_path: str, target_slides: List[int], trans_type: str = "morph", dur_ms: int = 1600):
    from pmg.self_healing import repair_xml_transitions as repair
    return repair(pptx_path, target_slides, trans_type, dur_ms)


def repair_traveling_anchor_gap(pptx_path: str, anchor_tag: str, gap_slides: List[int]):
    from pmg.self_healing import repair_traveling_anchor_gap as repair
    return repair(pptx_path, anchor_tag, gap_slides)


def substitute_with_fallback_skill(pptx_path: str, failed_slides: List[int], reason: str) -> str:
    from pmg.self_healing import substitute_with_fallback_skill as substitute
    return substitute(pptx_path, failed_slides, reason)


def review_and_heal_deck(pptx_path: str, max_repair_loops: int = 3, allow_substitution: bool = True) -> Tuple[bool, DeckReviewReport, List[str]]:
    from pmg.orchestrator import PMGOrchestrator
    from pmg.models import ValidationStatus
    manifest = PMGOrchestrator().run_pipeline(
        output_pptx_path=pptx_path, existing_pptx_path=pptx_path,
        max_repair_loops=max_repair_loops, allow_substitution=allow_substitution)
    report = DeckReviewReport(pptx_path, manifest.static_validation)
    report.status = manifest.status.value
    report.certified = manifest.status == ValidationStatus.PASS and bool(manifest.delivered_pptx)
    actions = [action for repair in manifest.repair_history for action in repair.get("actions_taken", [])]
    return report.certified, report, actions


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--auto-heal", action="store_true")
    parser.add_argument("--max-attempts", type=int, default=3)
    parser.add_argument("--no-substitution", action="store_true")
    args = parser.parse_args()
    if args.auto_heal:
        success, report, _ = review_and_heal_deck(args.pptx, args.max_attempts, not args.no_substitution)
    else:
        report = validate_deck(args.pptx)
        success = report.is_valid
    report.print_summary()
    sys.exit(0 if success else 1)
