"""
pmg.cli
=======
Canonical Command-Line Interface for PMG Skills:
- pmg build
- pmg validate
- pmg render
- pmg record
- pmg review
- pmg inspect
- pmg reconstruct
"""

import os
import sys
import json
import shutil
import argparse
from typing import Optional, List

from pmg.version import VERSION, get_engine_provenance
from pmg.models import ValidationStatus
from pmg.engine import (
    build_environmental_protection_deck,
    build_v1_2_single_slide_animation_showcase,
    LightThemeCraftDeck,
    MotionDeck,
    Themes,
)
from pmg.validator import StaticValidator, extract_expected_timeline
from pmg.video_pipeline import (
    export_video_with_powerpoint,
    record_actual_slideshow,
    probe_video_metadata,
)
from pmg.visual_reviewer import VisualAnimationReviewer
from pmg.orchestrator import PMGOrchestrator


def cmd_build(args: argparse.Namespace) -> int:
    orchestrator = PMGOrchestrator()

    def _builder(target_path: str) -> str:
        preset = (args.preset or "environment").lower()
        if preset in ("v1_2_showcase", "timeline", "animation"):
            return build_v1_2_single_slide_animation_showcase(target_path)
        elif preset in ("environment", "eco", "kite", "earth"):
            return build_environmental_protection_deck(target_path, num_slides=args.slides)
        elif preset == "camau":
            from scripts.create_camau_deck import build_camau_presentation
            return build_camau_presentation(target_path)
        else:
            deck = LightThemeCraftDeck(title=args.title or "PMG Presentation")
            for i in range(1, args.slides + 1):
                deck.add_craft_slide(
                    title=f"{args.title or 'PMG STAGE'} — PART {i:02d}",
                    category="PMG ENGINE",
                    badge=f"SLIDE {i:02d}",
                    subtitle="Generated via canonical PMG Engine with 60 FPS Morph choreography.",
                )
            deck.save(target_path)
            return target_path

    if args.skip_playback:
        manifest = orchestrator.run_pipeline(
            output_pptx_path=args.output,
            builder_fn=_builder,
            hold_duration_ms=args.hold_ms,
            require_video_playback=False,
        )
        shutil.copyfile(manifest.source_pptx, os.path.abspath(args.output))
        print(json.dumps(manifest.to_dict(), indent=2, ensure_ascii=False))
        return 0 if manifest.static_validation.get("is_valid") else 1

    manifest = orchestrator.run_pipeline(
        output_pptx_path=args.output,
        builder_fn=_builder,
        hold_duration_ms=args.hold_ms,
        max_repair_loops=args.max_repairs,
        require_video_playback=True,
    )
    print(f"[*] PMG Build & Playback Validation Complete: Status = {manifest.status.value}")
    if manifest.delivered_pptx:
        print(f"[+] Final Validated Deliverable: {manifest.delivered_pptx}")
        return 0
    else:
        print(f"[!] Delivery Blocked (Status={manifest.status.value}). Inspect QA logs at: {manifest.qa_dir}")
        return 1


def cmd_validate(args: argparse.Namespace) -> int:
    if args.static_only:
        validator = StaticValidator()
        res = validator.validate(args.pptx)
        print(json.dumps(res, indent=2, ensure_ascii=False))
        return 0 if res.get("is_valid") else 1

    orchestrator = PMGOrchestrator()
    out_path = args.output or args.pptx
    manifest = orchestrator.run_pipeline(
        output_pptx_path=out_path,
        existing_pptx_path=args.pptx,
        hold_duration_ms=args.hold_ms,
        max_repair_loops=args.max_repairs,
        require_video_playback=True,
    )
    print(json.dumps(manifest.to_dict(), indent=2, ensure_ascii=False))
    return 0 if manifest.status == ValidationStatus.PASS else 1


def cmd_render(args: argparse.Namespace) -> int:
    meta = export_video_with_powerpoint(
        pptx_path=args.pptx,
        output_video_path=args.output,
        slide_duration=args.slide_duration,
        requested_fps=args.fps,
        resolution=args.resolution,
    )
    print(json.dumps(meta.to_dict(), indent=2, ensure_ascii=False))
    return 0 if meta.decodable else 1


def cmd_record(args: argparse.Namespace) -> int:
    timeline = extract_expected_timeline(args.pptx, hold_duration_ms=args.hold_ms)
    meta, telemetry = record_actual_slideshow(
        pptx_path=args.pptx,
        output_video_path=args.output,
        expected_timeline=timeline,
        target_fps=args.fps,
    )
    print(json.dumps({"metadata": meta.to_dict(), "telemetry": telemetry}, indent=2, ensure_ascii=False))
    return 0 if meta.decodable else 1


def cmd_review(args: argparse.Namespace) -> int:
    validator = StaticValidator()
    static_val = validator.validate(args.pptx)
    timeline = extract_expected_timeline(args.pptx, hold_duration_ms=args.hold_ms)
    exp_meta = probe_video_metadata(args.export_video, requested_fps=60)
    rec_meta = probe_video_metadata(args.record_video, requested_fps=30)

    qa_dir = args.qa_dir or os.path.dirname(os.path.abspath(args.record_video))
    reviewer = VisualAnimationReviewer()
    report, analysis = reviewer.review(
        export_meta=exp_meta,
        recording_meta=rec_meta,
        recording_telemetry={"slide_events": [{"slide_number": s.slide_number, "actual_start_ms": s.absolute_start_ms, "actual_end_ms": s.absolute_end_ms} for s in timeline.slides]},
        expected_timeline=timeline,
        static_validation=static_val,
        qa_dir=qa_dir,
    )
    print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False))
    return 0 if report.status == ValidationStatus.PASS else 1


def cmd_inspect(args: argparse.Namespace) -> int:
    validator = StaticValidator()
    static_val = validator.validate(args.pptx)
    timeline = extract_expected_timeline(args.pptx, hold_duration_ms=args.hold_ms)
    if args.timeline_out:
        timeline.save(args.timeline_out)
    output = {
        "provenance": get_engine_provenance(entrypoint="pmg.cli.inspect"),
        "static_validation": static_val,
        "expected_timeline_summary": {
            "deck_title": timeline.deck_title,
            "total_slides": timeline.total_slides,
            "total_duration_ms": timeline.total_duration_ms,
        },
    }
    print(json.dumps(output, indent=2, ensure_ascii=False))
    return 0 if static_val.get("is_valid") else 1


def cmd_reconstruct(args: argparse.Namespace) -> int:
    """
    Base reconstruction entrypoint in v1.1-v1.3 (extracts timeline & reconstructs editable deck
    from a reference PPTX; full motion transfer in v1.4 and video reconstruction in v1.5).
    """
    source = args.source
    output = args.output
    if not os.path.exists(source):
        print(f"[!] Source file not found: {source}", file=sys.stderr)
        return 1

    timeline = extract_expected_timeline(source)
    deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)
    for s_spec in timeline.slides:
        slide = deck.add_slide(
            transition=s_spec.transition_type if s_spec.slide_number > 1 else "none",
            duration_ms=s_spec.transition_duration_ms or 1600,
        )
        for obj in s_spec.objects:
            x, y = obj.final_position
            w, h = max(0.4, obj.final_size[0]), max(0.3, obj.final_size[1])
            rot = obj.rotation_params.get("end_deg", 0.0)
            if obj.object_type == "textbox":
                tb = slide.shapes.add_textbox(
                    deck.prs.slide_width.__class__(int(x * 914400)),
                    deck.prs.slide_width.__class__(int(y * 914400)),
                    deck.prs.slide_width.__class__(int(w * 914400)),
                    deck.prs.slide_width.__class__(int(h * 914400)),
                )
                tb.name = obj.object_id
                tb.text_frame.text = obj.object_id.replace("!!", "")
            else:
                deck.add_thematic_traveling_entity(
                    slide, left=x, top=y, width=w, height=h, rot_deg=rot, name=obj.object_id
                )
    deck.save(output)
    print(json.dumps({"reconstructed_pptx": os.path.abspath(output), "slides": len(timeline.slides)}, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pmg",
        description=f"PMG Skills Canonical Engine CLI (v{VERSION})",
    )
    parser.add_argument("--version", action="version", version=f"pmg-skills {VERSION}")
    sub = parser.add_subparsers(dest="command", required=True)

    # 1. pmg build
    p_build = sub.add_parser("build", help="Build and validate a presentation through the canonical PMG engine")
    p_build.add_argument("-o", "--output", required=True, help="Output .pptx file path")
    p_build.add_argument(
        "--preset",
        default="environment",
        choices=["environment", "eco", "kite", "earth", "camau", "craft", "v1_2_showcase", "timeline", "animation"],
        help="Presentation preset",
    )
    p_build.add_argument("--title", default="Protecting Our Living Earth", help="Presentation title")
    p_build.add_argument("--slides", type=int, default=10, help="Number of slides (default: 10)")
    p_build.add_argument("--hold-ms", type=int, default=1200, help="Per-slide hold duration in ms (default: 1200)")
    p_build.add_argument("--max-repairs", type=int, default=3, help="Max self-healing loops (default: 3)")
    p_build.add_argument("--skip-playback", action="store_true", help="Skip live PowerPoint playback verification (marks NOT_VERIFIED)")
    p_build.set_defaults(func=cmd_build)

    # 2. pmg validate
    p_val = sub.add_parser("validate", help="Run static and/or full playback validation with 3-loop self-healing")
    p_val.add_argument("pptx", help="Path to input .pptx")
    p_val.add_argument("-o", "--output", help="Path to output validated .pptx")
    p_val.add_argument("--hold-ms", type=int, default=1200, help="Hold duration per slide in ms")
    p_val.add_argument("--max-repairs", type=int, default=3, help="Max self-healing loops (default: 3)")
    p_val.add_argument("--static-only", action="store_true", help="Run OpenXML static checks only")
    p_val.set_defaults(func=cmd_validate)

    # 3. pmg render
    p_ren = sub.add_parser("render", help="Export native MP4 video via Microsoft PowerPoint and inspect via FFprobe")
    p_ren.add_argument("pptx", help="Path to input .pptx")
    p_ren.add_argument("-o", "--output", required=True, help="Output .mp4 path")
    p_ren.add_argument("--fps", type=int, default=60, help="Requested frame rate (default: 60)")
    p_ren.add_argument("--resolution", type=int, default=1080, help="Vertical resolution (default: 1080)")
    p_ren.add_argument("--slide-duration", type=float, default=2.0, help="Default slide duration in seconds")
    p_ren.set_defaults(func=cmd_render)

    # 4. pmg record
    p_rec = sub.add_parser("record", help="Play presentation in PowerPoint Slide Show mode and record actual playback")
    p_rec.add_argument("pptx", help="Path to input .pptx")
    p_rec.add_argument("-o", "--output", required=True, help="Output recording .mp4 path")
    p_rec.add_argument("--fps", type=int, default=30, help="Target capture frame rate (default: 30)")
    p_rec.add_argument("--hold-ms", type=int, default=1200, help="Slide hold duration in ms")
    p_rec.set_defaults(func=cmd_record)

    # 5. pmg review
    p_rev = sub.add_parser("review", help="Run AI & Computer Vision animation review on exported and recorded videos")
    p_rev.add_argument("pptx", help="Path to .pptx presentation")
    p_rev.add_argument("--export-video", required=True, help="Path to PowerPoint exported MP4")
    p_rev.add_argument("--record-video", required=True, help="Path to live slideshow recorded MP4")
    p_rev.add_argument("--qa-dir", help="Directory to store extracted frame samples and reports")
    p_rev.add_argument("--hold-ms", type=int, default=1200, help="Slide hold duration in ms")
    p_rev.set_defaults(func=cmd_review)

    # 6. pmg inspect
    p_ins = sub.add_parser("inspect", help="Inspect OpenXML structure, Morph anchors, transitions, and timeline")
    p_ins.add_argument("pptx", help="Path to .pptx file")
    p_ins.add_argument("--timeline-out", help="Optional path to save expected_timeline.json")
    p_ins.add_argument("--hold-ms", type=int, default=1200, help="Slide hold duration in ms")
    p_ins.set_defaults(func=cmd_inspect)

    # 7. pmg reconstruct
    p_recon = sub.add_parser("reconstruct", help="Reconstruct editable PowerPoint slides from reference source")
    p_recon.add_argument("source", help="Reference .pptx or .mp4 file")
    p_recon.add_argument("-o", "--output", required=True, help="Output reconstructed .pptx path")
    p_recon.set_defaults(func=cmd_reconstruct)

    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
