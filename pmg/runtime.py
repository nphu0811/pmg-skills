"""
pmg.runtime
===========
Unified Public API and Runtime Orchestration for PMG Skills (v2.1.0).

Provides the canonical high-level interface for AI agents and automated scripts:
```python
from pmg import PMGRuntime

runtime = PMGRuntime.auto()
result = runtime.build(
    project="presentation_spec.json",
    output="presentation.pptx",
    validate=True,
)
print(result.output_path)
print(result.validation_report)
```
"""

import os
import json
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Union, Callable

from pmg.version import VERSION, get_engine_provenance
from pmg.models import ValidationStatus
from pmg.capabilities import (
    EnvironmentCapabilities,
    ExecutionMode,
    detect_environment,
)
from pmg.compiler import NativePPTXCompiler, CompilerReport
from pmg.validation import (
    ComprehensiveValidationReport,
    validate_presentation,
)


@dataclass
class PMGBuildResult:
    """Result object returned by runtime.build()."""
    output_path: str
    validation_report: Optional[ComprehensiveValidationReport] = None
    compiler_report: Optional[CompilerReport] = None
    is_valid: bool = False
    can_deliver: bool = False
    execution_mode: ExecutionMode = ExecutionMode.LOCAL_DESKTOP
    summary: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "output_path": self.output_path,
            "is_valid": self.is_valid,
            "can_deliver": self.can_deliver,
            "execution_mode": self.execution_mode.value,
            "summary": self.summary,
            "validation_report": self.validation_report.to_dict() if self.validation_report else None,
            "compiler_report": self.compiler_report.to_dict() if self.compiler_report else None,
        }


class PMGRuntime:
    """
    Universal, capability-aware PMG Runtime.
    Adapts automatically across Mode A (Web/Sandbox), Mode B (Local Desktop),
    and Mode C (Remote Worker).
    """

    def __init__(
        self,
        capabilities: Optional[EnvironmentCapabilities] = None,
        worker_url: Optional[str] = None,
        worker_token: Optional[str] = None,
        offline_cache_dir: Optional[str] = None,
    ):
        self.capabilities = capabilities or detect_environment()
        self.worker_url = worker_url or self.capabilities.remote_worker_url or os.environ.get("PMG_QA_WORKER_URL")
        self.worker_token = worker_token or os.environ.get("PMG_QA_WORKER_TOKEN")
        self.offline_cache_dir = offline_cache_dir or os.environ.get("PMG_OFFLINE_CACHE")
        self.compiler = NativePPTXCompiler()

    @classmethod
    def auto(
        cls,
        probes_override: Optional[Dict[str, Any]] = None,
        worker_url: Optional[str] = None,
        worker_token: Optional[str] = None,
        **kwargs,
    ) -> "PMGRuntime":
        """Automatically detects the host environment and returns a tailored runtime instance."""
        caps = detect_environment(probes_override=probes_override, worker_url=worker_url, worker_token=worker_token)
        return cls(capabilities=caps, worker_url=worker_url, worker_token=worker_token, **kwargs)

    def doctor(self) -> str:
        """Returns human-readable diagnostic report for current environment."""
        return self.capabilities.format_diagnostic_report()

    def build(
        self,
        project: Optional[Union[str, Dict[str, Any], Callable[[str], str]]] = None,
        output: str = "presentation.pptx",
        validate: bool = True,
        validate_levels: Optional[List[int]] = None,
        preset: Optional[str] = None,
        title: Optional[str] = None,
        slides: int = 4,
        **kwargs,
    ) -> PMGBuildResult:
        """
        Builds a native PowerPoint presentation from a JSON spec, preset, or builder function.
        PowerPoint Desktop is NOT required for compilation!
        """
        abs_output = os.path.abspath(output)
        os.makedirs(os.path.dirname(abs_output), exist_ok=True)

        comp_report: Optional[CompilerReport] = None

        # 1. Project Specification (JSON file path or spec dictionary)
        if isinstance(project, (str, dict)) and (isinstance(project, dict) or (isinstance(project, str) and (os.path.isfile(project) or project.strip().startswith("{")))):
            comp_report = self.compiler.compile_from_spec(project, abs_output)

        # 2. Programmatic builder callback
        elif callable(project):
            project(abs_output)
            comp_report = CompilerReport(output_path=abs_output, total_slides=slides)

        # 3. Built-in Preset
        elif preset:
            comp_report = self._build_preset(preset, abs_output, title=title, slides=slides, **kwargs)

        # 4. Default project specification if no project provided
        else:
            default_spec = {
                "theme": "LIGHT_ACADEMIC_CYAN",
                "slides": [
                    {
                        "title": title or "PMG Presentation",
                        "subtitle": "Native Cross-Platform Presentation Graphics",
                        "elements": [
                            {
                                "type": "card",
                                "name": "!!HeroCard",
                                "left": 1.0,
                                "top": 2.5,
                                "width": 5.5,
                                "height": 3.0,
                                "title": "Core Architecture",
                                "body": "Universal compilation across web sandboxes, local agents, and desktop environments.",
                            }
                        ],
                        "animations": [
                            {"target": "!!HeroCard", "effect": "fade", "duration_ms": 800, "delay_ms": 200}
                        ],
                    }
                ],
            }
            comp_report = self.compiler.compile_from_spec(default_spec, abs_output)

        # Validation phase
        val_report: Optional[ComprehensiveValidationReport] = None
        is_valid = True
        can_deliver = True

        if validate:
            val_report = validate_presentation(
                pptx_path=abs_output,
                levels=validate_levels,
                capabilities=self.capabilities,
                worker_url=self.worker_url,
                worker_token=self.worker_token,
            )
            is_valid = (val_report.overall_status == ValidationStatus.PASS)
            can_deliver = val_report.can_deliver

            # Companion manifest JSON
            report_out = f"{abs_output}.validation_report.json"
            val_report.save(report_out)

        summary = f"Built native presentation: {abs_output}. "
        if val_report:
            summary += val_report.summary
        else:
            summary += "Validation skipped."

        return PMGBuildResult(
            output_path=abs_output,
            validation_report=val_report,
            compiler_report=comp_report,
            is_valid=is_valid,
            can_deliver=can_deliver,
            execution_mode=self.capabilities.execution_mode,
            summary=summary,
        )

    def validate(
        self,
        pptx_path: str,
        levels: Optional[List[int]] = None,
        output_report_path: Optional[str] = None,
    ) -> ComprehensiveValidationReport:
        """Validates an existing presentation across Level 1, Level 2, and Level 3."""
        abs_path = os.path.abspath(pptx_path)
        report = validate_presentation(
            pptx_path=abs_path,
            levels=levels,
            capabilities=self.capabilities,
            worker_url=self.worker_url,
            worker_token=self.worker_token,
        )
        if output_report_path:
            report.save(output_report_path)
        return report

    def _build_preset(
        self,
        preset: str,
        output_path: str,
        title: Optional[str] = None,
        slides: int = 4,
        **kwargs,
    ) -> CompilerReport:
        p = preset.lower()
        if p in ("spotify", "app_shell", "music"):
            from pmg.cinematic_showcase import SpotifyAppShellBuilder
            b = SpotifyAppShellBuilder(topic_title=title or "SPOTIFY")
            b.build_deck(output_path, num_slides=slides)
        elif p in ("earth_space", "space_cinematic"):
            from pmg.cinematic_showcase import EarthSpaceMotionBuilder
            b = EarthSpaceMotionBuilder(topic_title=title or "EARTH AND SPACE")
            b.build_deck(output_path, num_slides=slides)
        elif p in ("bloom", "university_bloom", "university"):
            from pmg.cinematic_showcase import UniversityBloomBuilder
            b = UniversityBloomBuilder(presentation_title=title or "BÁO CÁO ĐỀ TÀI")
            b.build_deck(output_path)
        elif p in ("cinematic", "v1_3_cinematic"):
            from pmg.cinematic_builder import build_v1_3_cinematic_environmental_deck
            build_v1_3_cinematic_environmental_deck(output_path, num_slides=slides)
        elif p in ("environment", "eco"):
            from pmg.engine import build_environmental_protection_deck
            build_environmental_protection_deck(output_path, num_slides=slides)
        else:
            from pmg.engine import LightThemeCraftDeck
            deck = LightThemeCraftDeck(title=title or "PMG Presentation")
            for i in range(1, slides + 1):
                deck.add_craft_slide(
                    title=f"{title or 'PMG STAGE'} — PART {i:02d}",
                    category="PMG ENGINE",
                    body="Native PowerPoint presentation built via PMG Universal Runtime.",
                )
            deck.save(output_path)

        return CompilerReport(output_path=output_path, total_slides=slides)

