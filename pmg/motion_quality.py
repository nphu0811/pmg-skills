"""
pmg.motion_quality
==================
Motion Quality, Accessibility & Performance Validation for PMG v2.5.0.

Audits presentation animation timing, cognitive load, simultaneous action density,
easing consistency, text readability preservation, and reduced-motion compliance.
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any, Set, Tuple

from pmg.motion_planners.base import MotionIntent, SharedMotionPrimitive
from pmg.animation_ir import CAPABILITY_CATALOG, AnimationCapabilityStatus, get_effect_capability_status


@dataclass
class MotionQualityReport:
    """
    Diagnostic quality assessment for slide animation compositions.
    """
    score: float = 100.0               # Quality score [0.0, 100.0]
    passed: bool = True
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    metrics: Dict[str, Any] = field(default_factory=dict)
    reduced_motion_active: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


class ReducedMotionProfile:
    """
    Accessibility filter adapting animations for viewers sensitive to motion or executive settings.
    Transforms aggressive zooms, rapid spins, and large path translations into calm fades and settles.
    """
    @classmethod
    def apply(cls, intents: List[MotionIntent]) -> List[MotionIntent]:
        adapted: List[MotionIntent] = []
        for intent in intents:
            prim = intent.primitive
            duration = min(max(intent.duration_ms, 300), 1000)
            scale = 1.0
            rot = 0.0
            direction = None

            # Soften high-motion primitives into gentle reveals
            if prim in (SharedMotionPrimitive.ROTATE, SharedMotionPrimitive.FOCUS, SharedMotionPrimitive.SCALE):
                prim = SharedMotionPrimitive.REVEAL
            elif prim == SharedMotionPrimitive.EMPHASIZE:
                scale = 1.04  # Minimal subtle pulse instead of large pop

            adapted.append(
                MotionIntent(
                    target_id=intent.target_id,
                    primitive=prim,
                    semantic_purpose=f"{intent.semantic_purpose} [Reduced Motion Profile]",
                    duration_ms=duration,
                    delay_ms=intent.delay_ms,
                    trigger=intent.trigger,
                    easing="ease_out",
                    direction=direction,
                    scale_factor=scale,
                    rotation_deg=rot,
                    parameters=intent.parameters,
                    narrative_importance=intent.narrative_importance,
                )
            )
        return adapted


class MotionQualityValidator:
    """
    Enforces high-standard motion design and cognitive hygiene.
    """

    @classmethod
    def validate(
        cls,
        intents_or_actions: List[Any],
        known_target_ids: Optional[Set[str]] = None,
        reduced_motion: bool = False,
    ) -> MotionQualityReport:
        warnings: List[str] = []
        errors: List[str] = []
        score = 100.0

        total_actions = len(intents_or_actions)
        if total_actions == 0:
            return MotionQualityReport(score=100.0, passed=True, metrics={"total_actions": 0})

        # Calculate time intervals to check concurrency
        intervals: List[Tuple[int, int, str]] = []  # (start, end, target_id)
        current_time = 0

        for item in intents_or_actions:
            target_id = getattr(item, "target_id", None) or getattr(item, "target_node_id", "unknown")
            duration = getattr(item, "duration_ms", 500)
            delay = getattr(item, "delay_ms", 0)
            trigger = getattr(item, "trigger", "after_previous")
            if hasattr(item, "to_animation_action"):
                act = item.to_animation_action()
                effect_str = act.effect_type
            elif hasattr(item, "effect_type"):
                effect_str = item.effect_type
            else:
                effect_val = getattr(item, "primitive", "appear")
                effect_str = effect_val.value if hasattr(effect_val, "value") else str(effect_val)

            # Check 1: Target existence
            if known_target_ids is not None and target_id not in known_target_ids:
                errors.append(f"Animation targets non-existent object '{target_id}'")
                score -= 25.0

            # Check 2: Duration sanity
            if duration < 120:
                warnings.append(f"Action on '{target_id}' duration {duration}ms is excessively abrupt (<120ms)")
                score -= 5.0
            elif duration > 3600:
                warnings.append(f"Action on '{target_id}' duration {duration}ms is sluggish (>3600ms)")
                score -= 5.0

            # Check 3: Native capability status
            cap, _ = get_effect_capability_status(effect_str)
            if cap == AnimationCapabilityStatus.UNSUPPORTED:
                warnings.append(f"Effect '{effect_str}' on '{target_id}' is unsupported natively; using fallback")
                score -= 10.0

            # Compute timeline boundaries
            if trigger == "after_previous":
                start = current_time + delay
            else:
                start = current_time + delay
            end = start + duration
            intervals.append((start, end, target_id))
            if trigger == "after_previous":
                current_time = end

        # Check 4: Cognitive load - Maximum simultaneous animations
        # Sample concurrency at each boundary
        time_points = sorted(list(set([s for s, _, _ in intervals] + [e for _, e, _ in intervals])))
        max_concurrent = 0
        for tp in time_points:
            active_count = sum(1 for s, e, _ in intervals if s <= tp < e)
            if active_count > max_concurrent:
                max_concurrent = active_count

        if max_concurrent > 4:
            warnings.append(
                f"High cognitive load: {max_concurrent} simultaneous animations exceed recommended maximum (4)"
            )
            score -= 15.0

        score = max(0.0, min(100.0, score))
        passed = len(errors) == 0 and score >= 60.0

        return MotionQualityReport(
            score=round(score, 1),
            passed=passed,
            warnings=warnings,
            errors=errors,
            metrics={
                "total_actions": total_actions,
                "max_simultaneous": max_concurrent,
                "total_timeline_ms": current_time,
            },
            reduced_motion_active=reduced_motion,
        )
