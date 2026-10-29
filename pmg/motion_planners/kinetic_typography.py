"""
pmg.motion_planners.kinetic_typography
======================================
Specialized Kinetic Typography Motion Planner for PMG v2.5.0 (`pmg-kinetic-typography`).

Designs purposeful typographic animation patterns (title unveiling, word-by-word cadence,
keyword punch, typewriter rhythm, and cross-slide kinetic headline morphs).
"""

from __future__ import annotations
from typing import Dict, List, Optional, Any, Set

from pmg.motion_planners.base import BaseMotionPlanner, MotionIntent, SharedMotionPrimitive


class KineticTypographyPlanner(BaseMotionPlanner):
    """
    Intelligent motion planner for text, titles, subtitles, and keywords.
    Respects reading cadence, cognitive absorption limits, and typographic hierarchy.
    """
    @property
    def planner_name(self) -> str:
        return "pmg-kinetic-typography"

    @property
    def supported_object_types(self) -> Set[str]:
        return {"typography", "title", "headline", "keyword", "body_text"}

    def plan_motion(self, context: Dict[str, Any]) -> List[MotionIntent]:
        """
        Plans typographic animation based on style and hierarchy.
        Context keys:
          - style: 'title_unveiling', 'word_stagger', 'keyword_punch', 'typewriter', 'paragraph_sequence'
          - target_id: ID of the text element or list of span IDs
          - subtitle_id: Optional ID for accompanying subtitle
          - keyword_ids: Optional list of keyword spans to emphasize
          - words: Optional list of individual word IDs for word-level stagger
        """
        intents: List[MotionIntent] = []
        style = context.get("style", "title_unveiling").lower()
        target_id = context.get("target_id", "title")
        subtitle_id = context.get("subtitle_id")
        keyword_ids = context.get("keyword_ids", [])
        words = context.get("words", [])

        # 1. Authoritative Title Unveiling
        if style == "title_unveiling":
            intents.append(
                MotionIntent(
                    target_id=target_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose="Authoritative headline unveiling establishing slide topic",
                    duration_ms=650,
                    delay_ms=50,
                    trigger="after_previous",
                    easing="ease_out",
                    direction="right",
                    narrative_importance="primary",
                )
            )
            if subtitle_id:
                intents.append(
                    MotionIntent(
                        target_id=subtitle_id,
                        primitive=SharedMotionPrimitive.REVEAL,
                        semantic_purpose="Subtitle settles gently beneath headline to provide context",
                        duration_ms=500,
                        delay_ms=120,
                        trigger="after_previous",
                        direction="up",
                        easing="ease_out",
                        narrative_importance="secondary",
                    )
                )

        # 2. Word-by-Word Reading Cadence
        elif style == "word_stagger":
            if words:
                for idx, word_id in enumerate(words):
                    intents.append(
                        MotionIntent(
                            target_id=word_id,
                            primitive=SharedMotionPrimitive.STAGGER,
                            semantic_purpose=f"Word {idx+1} revealed in natural reading cadence",
                            duration_ms=250,
                            delay_ms=idx * 60,
                            trigger="after_previous" if idx == 0 else "with_previous",
                            easing="linear",
                            narrative_importance="primary",
                        )
                    )
            else:
                # Fallback to rhythmic line reveal
                intents.append(
                    MotionIntent(
                        target_id=target_id,
                        primitive=SharedMotionPrimitive.REVEAL,
                        semantic_purpose="Staggered typographic reveal across text block",
                        duration_ms=700,
                        delay_ms=60,
                        trigger="after_previous",
                    )
                )

        # 3. Keyword Punch / Emphasis
        elif style == "keyword_punch":
            # Primary phrase reveals first
            intents.append(
                MotionIntent(
                    target_id=target_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose="Framing statement revealed to audience",
                    duration_ms=550,
                    delay_ms=50,
                    trigger="after_previous",
                )
            )
            # Essential keywords scale pop into focus
            for kidx, kid in enumerate(keyword_ids):
                intents.append(
                    MotionIntent(
                        target_id=kid,
                        primitive=SharedMotionPrimitive.EMPHASIZE,
                        semantic_purpose=f"High-impact keyword punch commanding focal attention",
                        duration_ms=380,
                        delay_ms=100 + kidx * 90,
                        trigger="after_previous",
                        scale_factor=1.14,
                        narrative_importance="hero",
                    )
                )

        # 4. Typewriter Cadence for Technical Specs / Quotes
        elif style == "typewriter":
            intents.append(
                MotionIntent(
                    target_id=target_id,
                    primitive=SharedMotionPrimitive.REVEAL,
                    semantic_purpose="Stepped typewriter sequence evoking real-time computation",
                    duration_ms=1200,
                    delay_ms=80,
                    trigger="after_previous",
                    easing="linear",
                    parameters={"typewriter_steps": context.get("character_count", 30)},
                    narrative_importance="primary",
                )
            )

        # 5. Reading-Order Paragraph Sequence
        elif style == "paragraph_sequence":
            paragraphs = context.get("paragraphs", [target_id])
            for idx, pid in enumerate(paragraphs):
                intents.append(
                    MotionIntent(
                        target_id=pid,
                        primitive=SharedMotionPrimitive.SEQUENCE,
                        semantic_purpose=f"Paragraph {idx+1} disclosed sequentially to avoid cognitive overload",
                        duration_ms=500,
                        delay_ms=idx * 160,
                        trigger="after_previous",
                        direction="up",
                        easing="ease_out",
                        narrative_importance="secondary",
                    )
                )

        # 6. Kinetic Cross-Slide Headline Morph
        elif style == "kinetic_morph":
            # Enforces persistent '!!' cross-slide Morph transition
            intents.append(
                MotionIntent(
                    target_id=target_id,
                    primitive=SharedMotionPrimitive.MORPH,
                    semantic_purpose="Kinetic headline glides across slides preserving visual continuity",
                    duration_ms=1000,
                    delay_ms=0,
                    trigger="after_previous",
                    narrative_importance="hero",
                )
            )

        return intents

    def explain_selection(self, intent: MotionIntent, context: Dict[str, Any]) -> str:
        return (
            f"KineticTypographyPlanner selected '{intent.primitive.value}' for '{intent.target_id}' "
            f"to respect reading cadence and establish strong typographic hierarchy."
        )
