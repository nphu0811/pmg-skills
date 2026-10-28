"""
benchmarks
==========
24-Case Animation Benchmark Suite for PMG Skills (v1.5.1 Upgrade B).
"""

from benchmarks.cases import BENCHMARK_CASES, BenchmarkCaseSpec, get_benchmark_case
from benchmarks.metrics import CaseBenchmarkMetrics, aggregate_benchmark_metrics
from benchmarks.runner import BenchmarkRunner

__all__ = [
    "BENCHMARK_CASES",
    "BenchmarkCaseSpec",
    "get_benchmark_case",
    "CaseBenchmarkMetrics",
    "aggregate_benchmark_metrics",
    "BenchmarkRunner",
]

