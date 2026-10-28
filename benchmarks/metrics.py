"""
benchmarks.metrics
==================
Quantitative benchmark metrics for the 24-case PMG Animation Benchmark Suite (Upgrade B).
"""

from dataclasses import dataclass, field, asdict
from typing import Dict, Any, List, Optional


@dataclass
class CaseBenchmarkMetrics:
    case_id: int
    case_slug: str
    group: str
    title: str
    is_negative_or_repair_test: bool
    detected_broken_pre_heal: Optional[bool]
    final_status: str
    passed: bool
    object_detection_rate: float
    motion_classification_accuracy: float
    start_time_error_ms: float
    duration_error_ms: float
    trajectory_rmse_pct: float
    final_position_error_pct: float
    visibility_correctness: float
    text_preservation_rate: float
    editability_score: float
    repair_loops_used: int
    elapsed_sec: float
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def aggregate_benchmark_metrics(results: List[CaseBenchmarkMetrics]) -> Dict[str, Any]:
    if not results:
        return {"total_cases": 0, "passed_cases": 0, "pass_rate": 0.0}

    n = len(results)
    passed = sum(1 for r in results if r.passed)
    groups: Dict[str, Dict[str, Any]] = {}
    for r in results:
        g = groups.setdefault(r.group, {"total": 0, "passed": 0})
        g["total"] += 1
        if r.passed:
            g["passed"] += 1

    return {
        "total_cases": n,
        "passed_cases": passed,
        "failed_cases": n - passed,
        "pass_rate": round((passed / n) * 100.0, 2),
        "avg_object_detection_rate": round(sum(r.object_detection_rate for r in results) / n, 4),
        "avg_motion_classification_accuracy": round(sum(r.motion_classification_accuracy for r in results) / n, 4),
        "avg_start_time_error_ms": round(sum(r.start_time_error_ms for r in results) / n, 2),
        "avg_duration_error_ms": round(sum(r.duration_error_ms for r in results) / n, 2),
        "avg_trajectory_rmse_pct": round(sum(r.trajectory_rmse_pct for r in results) / n, 2),
        "avg_final_position_error_pct": round(sum(r.final_position_error_pct for r in results) / n, 2),
        "avg_visibility_correctness": round(sum(r.visibility_correctness for r in results) / n, 4),
        "avg_text_preservation_rate": round(sum(r.text_preservation_rate for r in results) / n, 4),
        "avg_editability_score": round(sum(r.editability_score for r in results) / n, 4),
        "group_summary": groups,
    }

