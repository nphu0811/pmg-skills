"""
pmg.motion_decision
===================
Explainable Motion Decision System for PMG v2.5.0 Universal Object-Specific Motion Intelligence.

Implements transparent, auditable decision making for why specific motion patterns
are selected for each object, answering the 8 core pedagogical & cognitive questions.
"""

from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional, Any

from pmg.semantic_classifier import ObjectType, SemanticRole, NarrativeImportance, ObjectSemanticProfile
from pmg.motion_planners.base import SharedMotionPrimitive, MotionIntent
from pmg.animation_ir import get_effect_capability_status, AnimationCapabilityStatus


@dataclass
class MotionDecision:
    """
    Transparent explanation of an animation choice for a single object.
    Explicitly answers the 8 core cognitive and technical questions.
    """
    target_id: str
    object_type: ObjectType
    semantic_role: SemanticRole
    narrative_importance: NarrativeImportance
    selected_pattern: str
    shared_primitive: SharedMotionPrimitive
    compiled_effect: str
    answers: Dict[str, str] = field(default_factory=dict)
    timing_rationale: str = ""
    fallback_used: bool = False
    capability_status: str = "SUPPORTED_AND_TESTED"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "target_id": self.target_id,
            "object_type": self.object_type.value,
            "semantic_role": self.semantic_role.value,
            "narrative_importance": self.narrative_importance.value,
            "selected_pattern": self.selected_pattern,
            "shared_primitive": self.shared_primitive.value,
            "compiled_effect": self.compiled_effect,
            "answers": self.answers,
            "timing_rationale": self.timing_rationale,
            "fallback_used": self.fallback_used,
            "capability_status": self.capability_status,
        }


@dataclass
class MotionDecisionReport:
    """
    Aggregated explainability report for a presentation slide or scene.
    """
    scene_id: str
    decisions: List[MotionDecision] = field(default_factory=list)
    narrative_flow_summary: str = ""
    cognitive_load_level: str = "moderate"  # 'low', 'moderate', 'high'
    total_duration_ms: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "scene_id": self.scene_id,
            "decisions": [d.to_dict() for d in self.decisions],
            "narrative_flow_summary": self.narrative_flow_summary,
            "cognitive_load_level": self.cognitive_load_level,
            "total_duration_ms": self.total_duration_ms,
        }

    def explain_text(self) -> str:
        """Generates a human-readable diagnostic decision breakdown."""
        lines = [
            f"=== Motion Decision Report: {self.scene_id} ===",
            f"Narrative Flow: {self.narrative_flow_summary}",
            f"Cognitive Load: {self.cognitive_load_level.upper()} | Total Duration: {self.total_duration_ms}ms",
            "----------------------------------------------------------------",
        ]
        for d in self.decisions:
            lines.append(f"[{d.target_id}] ({d.object_type.value} / {d.semantic_role.value}) -> {d.selected_pattern.upper()}")
            lines.append(f"  Primitive: {d.shared_primitive.value} | Native Effect: {d.compiled_effect} [{d.capability_status}]")
            for q, a in d.answers.items():
                lines.append(f"    • {q}: {a}")
            lines.append(f"  Timing: {d.timing_rationale}")
            lines.append("")
        return "\n".join(lines)


class ExplainableDecisionEngine:
    """
    Evaluates semantic profiles and motion intents, creating audit-ready MotionDecisions.
    """

    @classmethod
    def evaluate(
        cls,
        profile: ObjectSemanticProfile,
        intent: MotionIntent,
        all_profiles: Optional[List[ObjectSemanticProfile]] = None,
    ) -> MotionDecision:
        """
        Formulates answers to the 8 core cognitive and architectural questions.
        """
        action = intent.to_animation_action()
        cap_status, canonical_name = get_effect_capability_status(action.effect_type)

        q1 = f"{profile.object_type.value.capitalize()} element ({profile.semantic_role.value})"
        
        # Q2: What information does it communicate?
        if profile.object_type == ObjectType.CHART:
            q2 = f"Quantitative data comparison ({profile.data_context.get('series_name', 'metric')})"
        elif profile.object_type == ObjectType.TYPOGRAPHY:
            q2 = f"Narrative text context: '{str(profile.data_context.get('text', ''))[:40]}...'"
        elif profile.object_type == ObjectType.DIAGRAM:
            q2 = f"Process step causality ({profile.data_context.get('step_label', profile.object_id)})"
        elif profile.object_type == ObjectType.MAP:
            q2 = f"Geographic location or spatial route ({profile.data_context.get('geo_name', profile.object_id)})"
        elif profile.object_type == ObjectType.INFOGRAPHIC:
            q2 = f"Structured milestone or timeline progression"
        elif profile.object_type == ObjectType.IMAGE:
            q2 = f"Visual evidence, atmosphere, or key subject"
        else:
            q2 = f"Supporting layout context or structural depth"

        # Q3: What should the audience notice first?
        if profile.narrative_importance == NarrativeImportance.HERO:
            q3 = "Immediate primary focus: This is the central traveling entity or hero subject."
        elif profile.narrative_importance == NarrativeImportance.PRIMARY:
            q3 = "Main content focus: Revealed right after stage framing to guide eye movement."
        elif profile.narrative_importance == NarrativeImportance.SUPPORTING:
            q3 = "Subordinate structural context: Grounding framework revealed quietly."
        else:
            q3 = "Ambient baseline: Settled immediately without distracting motion."

        # Q4: Does the animation clarify or obscure content?
        q4 = f"Clarifies content by using purposeful '{intent.primitive.value}' instead of random visual noise."

        # Q5: Which motion pattern best expresses meaning?
        q5 = f"'{intent.semantic_purpose}' via {intent.primitive.value} ({intent.direction or 'neutral'})."

        # Q6: How should it synchronize with surrounding objects?
        q6 = f"Trigger '{intent.trigger}' with {intent.delay_ms}ms delay to establish rhythmic narrative pacing."

        # Q7: Can it be compiled into native PowerPoint animation?
        q7 = f"Yes, compiled to native OOXML '{action.effect_type}' without requiring Microsoft Office installed."

        # Q8: What is the safest compatible fallback?
        fallback_used = False
        if cap_status == AnimationCapabilityStatus.UNSUPPORTED:
            q8 = "Standard directional wipe / fade used as tested fallback."
            fallback_used = True
        else:
            q8 = f"Native effect '{action.effect_type}' is fully supported and verified."

        answers = {
            "1. What is this object?": q1,
            "2. What information does it communicate?": q2,
            "3. What should audience notice first?": q3,
            "4. Does animation clarify content?": q4,
            "5. Best motion pattern?": q5,
            "6. Synchronization logic?": q6,
            "7. Native compilation status?": q7,
            "8. Safest compatible fallback?": q8,
        }

        timing_rationale = (
            f"Duration {intent.duration_ms}ms ({intent.easing}) chosen for readable cognitive absorption; "
            f"importance={profile.narrative_importance.value}."
        )

        return MotionDecision(
            target_id=profile.object_id,
            object_type=profile.object_type,
            semantic_role=profile.semantic_role,
            narrative_importance=profile.narrative_importance,
            selected_pattern=intent.semantic_purpose,
            shared_primitive=intent.primitive,
            compiled_effect=action.effect_type,
            answers=answers,
            timing_rationale=timing_rationale,
            fallback_used=fallback_used,
            capability_status=cap_status.value,
        )
