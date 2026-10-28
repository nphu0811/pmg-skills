"""
pmg.cli
=======
Canonical Command-Line Interface for PMG Skills (v1.5.1):
- pmg build
- pmg validate (supports --worker auto|local|remote|static and --no-video)
- pmg render
- pmg record
- pmg review
- pmg inspect
- pmg reconstruct
- pmg worker (doctor, serve, status)
- pmg benchmark (run --all, run --case <id>, report)
"""

import os
import sys
import json
import shutil
import argparse
from typing import Optional, List

from pmg.version import VERSION, get_engine_provenance, verify_runtime_dependencies
from pmg.models import ValidationStatus, ReviewReport
from pmg.engine import (
    build_environmental_protection_deck,
    build_v1_2_single_slide_animation_showcase,
    LightThemeCraftDeck,
    MotionDeck,
    Themes,
)
from pmg.cinematic_builder import build_v1_3_cinematic_environmental_deck
from pmg.reverse_engineer import (
    PPTXMotionReverseEngineer,
    build_v1_4_reference_smartphone_slide,
)
from pmg.video_reconstructor import VideoMotionReconstructor
from pmg.validator import StaticValidator, extract_expected_timeline
from pmg.video_pipeline import (
    export_video_with_powerpoint,
    record_actual_slideshow,
    probe_video_metadata,
)
from pmg.visual_reviewer import VisualAnimationReviewer
from pmg.orchestrator import PMGOrchestrator
from pmg.worker import (
    check_worker_health,
    PMGValidationWorkerServer,
    PMGWorkerClient,
)
from benchmarks import BenchmarkRunner


def cmd_build(args: argparse.Namespace) -> int:
    orchestrator = PMGOrchestrator()

    def _builder(target_path: str) -> str:
        preset = (args.preset or "v1_3_cinematic").lower()
        if preset in ("spotify", "app_shell", "music", "mac_lenin"):
            from pmg.cinematic_showcase import SpotifyAppShellBuilder
            b = SpotifyAppShellBuilder(
                topic_title=args.title or "SPOTIFY",
                topic_subtitle=getattr(args, "subtitle", "") or "TEAM 8 PRESENTS",
                author_name=getattr(args, "author", "Presenter") or "Presenter",
            )
            return b.build_deck(target_path, num_slides=args.slides)
        elif preset in ("earth_space", "space_cinematic", "reticle", "radar_wheel"):
            from pmg.cinematic_showcase import EarthSpaceMotionBuilder
            b = EarthSpaceMotionBuilder(
                topic_title=args.title or "EARTH AND SPACE",
                topic_subtitle=getattr(args, "subtitle", "") or "A CINEMATIC EXPLORATION",
            )
            return b.build_deck(target_path, num_slides=args.slides)
        elif preset in ("university_bloom", "bloom", "petal", "toc_8", "university"):
            from pmg.cinematic_showcase import UniversityBloomBuilder
            b = UniversityBloomBuilder(
                presentation_title=args.title or "BÁO CÁO ĐỀ TÀI",
                subtitle=getattr(args, "subtitle", "") or "KHẢO SÁT Ý KIẾN VÀ ĐỀ XUẤT",
            )
            return b.build_deck(target_path)
        elif preset in ("space", "v1_5_1_space", "beyond_the_horizon"):
            from pmg.space_builder import build_v1_5_1_space_exploration_20_slide_deck
            return build_v1_5_1_space_exploration_20_slide_deck(target_path, num_slides=args.slides)
        elif preset in ("v1_3_cinematic", "cinematic"):
            return build_v1_3_cinematic_environmental_deck(target_path, num_slides=args.slides)
        elif preset in ("v1_4_reference", "smartphone"):
            return build_v1_4_reference_smartphone_slide(target_path)
        elif preset in ("v1_2_showcase", "timeline", "animation"):
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
            worker_mode="static",
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
        worker_mode=getattr(args, "worker", "auto"),
    )
    print(f"[*] PMG Build & Playback Validation Complete: Status = {manifest.status.value}")
    if manifest.delivered_pptx:
        print(f"[+] Final Validated Deliverable: {manifest.delivered_pptx}")
        return 0
    else:
        print(f"[!] Delivery Blocked (Status={manifest.status.value}). Inspect QA logs at: {manifest.qa_dir}")
        return 1


def cmd_validate(args: argparse.Namespace) -> int:
    worker_mode = getattr(args, "worker", "auto") or "auto"
    if args.static_only or getattr(args, "no_video", False) or worker_mode == "static":
        orchestrator = PMGOrchestrator()
        out_path = args.output or args.pptx
        manifest = orchestrator.run_pipeline(
            output_pptx_path=out_path,
            existing_pptx_path=args.pptx,
            hold_duration_ms=args.hold_ms,
            max_repair_loops=args.max_repairs,
            require_video_playback=False,
            worker_mode="static",
        )
        print(json.dumps(manifest.to_dict(), indent=2, ensure_ascii=False))
        return 0 if manifest.static_validation.get("is_valid") else 1

    orchestrator = PMGOrchestrator()
    out_path = args.output or args.pptx
    manifest = orchestrator.run_pipeline(
        output_pptx_path=out_path,
        existing_pptx_path=args.pptx,
        hold_duration_ms=args.hold_ms,
        max_repair_loops=args.max_repairs,
        require_video_playback=True,
        worker_mode=worker_mode,
        worker_url=getattr(args, "worker_url", None),
        worker_token=getattr(args, "worker_token", None),
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
    telemetry_path = getattr(args, "telemetry", None) or os.path.join(qa_dir, "recording_telemetry.json")
    if not os.path.isfile(telemetry_path):
        print(json.dumps(ReviewReport(status=ValidationStatus.NOT_VERIFIED,
                                     summary="Actual recording telemetry is required; run pmg record or provide --telemetry.").to_dict()))
        return 1
    with open(telemetry_path, encoding="utf-8") as stream:
        telemetry = json.load(stream)
    reviewer = VisualAnimationReviewer()
    report, analysis = reviewer.review(
        export_meta=exp_meta,
        recording_meta=rec_meta,
        recording_telemetry=telemetry,
        expected_timeline=timeline,
        static_validation=static_val,
        qa_dir=qa_dir,
    )
    print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False))
    return 0 if report.status == ValidationStatus.PASS else 1


def cmd_inspect(args: argparse.Namespace) -> int:
    validator = StaticValidator()
    static_val = validator.validate(args.pptx)
    if not static_val.get("is_valid"):
        print(json.dumps({"static_validation": static_val}, indent=2, ensure_ascii=False))
        return 1
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
    source = args.source
    output = args.output
    if not os.path.exists(source):
        print(f"[!] Source file not found: {source}", file=sys.stderr)
        return 1

    if source.lower().endswith(".pptx"):
        re_engine = PPTXMotionReverseEngineer()
        spec = re_engine.extract_motion_spec(source)
        if getattr(args, "spec_out", None):
            spec.save(args.spec_out)

        content_transfer = {}
        if getattr(args, "transfer_title", None):
            content_transfer["title"] = args.transfer_title

        re_engine.reconstruct_and_transfer(
            motion_spec=spec,
            output_pptx_path=output,
            content_transfer=content_transfer,
        )
        comp = re_engine.compare_reference_and_reconstructed(source, output)
        if getattr(args, "compare_out", None):
            os.makedirs(os.path.dirname(os.path.abspath(args.compare_out)), exist_ok=True)
            with open(args.compare_out, "w", encoding="utf-8") as f:
                json.dump(comp, f, indent=2, ensure_ascii=False)

        print(json.dumps({"reconstructed_pptx": os.path.abspath(output), "comparison": comp}, indent=2, ensure_ascii=False))
        return 0

    if source.lower().endswith((".mp4", ".mov", ".avi", ".mkv")):
        vid_recon = VideoMotionReconstructor(
            max_duration_sec=getattr(args, "max_duration_sec", 8.0) or 8.0,
            max_scenes=getattr(args, "max_scenes", 2) or 2,
        )
        rep = vid_recon.reconstruct_from_video(source, output)
        if getattr(args, "compare_out", None):
            rep.save(args.compare_out)
        print(
            json.dumps(
                {
                    "reconstructed_pptx": rep.reconstructed_pptx,
                    "scenes_detected": rep.scenes_detected,
                    "total_editable_objects_reconstructed": rep.total_editable_objects_reconstructed,
                    "total_native_animations_reconstructed": rep.total_native_animations_reconstructed,
                    "com_animation_pane_verified": rep.com_animation_pane_verified,
                    "com_animation_pane_count": rep.com_animation_pane_count,
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        return 0

    print(f"[!] Unsupported source format for reconstruct: {source}", file=sys.stderr)
    return 1


def cmd_worker(args: argparse.Namespace) -> int:
    action = args.worker_action
    if action == "doctor":
        health = check_worker_health()
        deps = verify_runtime_dependencies(require_windows_com=True)
        payload = {"worker_health": health.to_dict(), "runtime_dependencies": deps}
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0 if health.status == "healthy" else 1

    if action == "status":
        client = PMGWorkerClient(worker_url=args.url, worker_token=args.token)
        status = client.get_remote_status()
        print(json.dumps(status, indent=2, ensure_ascii=False))
        return 0 if status.get("reachable") and status.get("status") == "healthy" else 1

    if action == "serve":
        srv = PMGValidationWorkerServer(
            host=args.host,
            port=args.port,
            auth_token=args.token,
        )
        print(f"[*] PMG Windows PowerPoint Validation Worker listening on {srv.url} ...")
        try:
            srv.serve_forever()
        except KeyboardInterrupt:
            srv.shutdown()
        return 0

    return 1


def cmd_benchmark(args: argparse.Namespace) -> int:
    runner = BenchmarkRunner()
    action = args.bench_action
    if action == "report":
        rep = runner.load_latest_report()
        print(json.dumps(rep, indent=2, ensure_ascii=False))
        return 0 if "error" not in rep else 1

    if action == "run":
        w_mode = args.worker or "auto"
        if args.case is not None:
            res = runner.run_all(case_ids=[args.case], worker_mode=w_mode, hold_duration_ms=args.hold_ms)
            print(json.dumps(res, indent=2, ensure_ascii=False))
            return 0 if res["summary"]["passed_cases"] == len(res["cases"]) else 1
        else:
            res = runner.run_all(case_ids=None, worker_mode=w_mode, hold_duration_ms=args.hold_ms)
            print(json.dumps(res["summary"], indent=2, ensure_ascii=False))
            return 0 if res["summary"]["passed_cases"] == res["summary"]["total_cases"] else 1

    return 1


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
        default="v1_3_cinematic",
        choices=[
            "space",
            "v1_5_1_space",
            "beyond_the_horizon",
            "environment",
            "eco",
            "kite",
            "earth",
            "camau",
            "craft",
            "v1_2_showcase",
            "timeline",
            "animation",
            "v1_3_cinematic",
            "cinematic",
            "v1_4_reference",
            "smartphone",
            "spotify",
            "app_shell",
            "earth_space",
            "radar_wheel",
            "university_bloom",
            "bloom",
        ],
        help="Presentation preset",
    )
    p_build.add_argument("--title", default="Protecting Our Living Earth", help="Presentation title")
    p_build.add_argument("--slides", type=int, default=10, help="Number of slides (default: 10)")
    p_build.add_argument("--subtitle", default="", help="Presentation subtitle")
    p_build.add_argument("--author", default="Presenter", help="Author/Presenter name")
    p_build.add_argument("--hold-ms", type=int, default=1200, help="Per-slide hold duration in ms (default: 1200)")
    p_build.add_argument("--max-repairs", type=int, default=3, help="Max self-healing loops (default: 3)")
    p_build.add_argument(
        "--worker",
        default="auto",
        choices=["auto", "local", "remote", "static"],
        help="Validation worker mode (default: auto)",
    )
    p_build.add_argument(
        "--skip-playback",
        action="store_true",
        help="Skip live PowerPoint playback verification (marks NOT_VERIFIED)",
    )
    p_build.set_defaults(func=cmd_build)

    # 2. pmg validate
    p_val = sub.add_parser("validate", help="Run static and/or full playback validation with 3-loop self-healing")
    p_val.add_argument("pptx", help="Path to input .pptx")
    p_val.add_argument("-o", "--output", help="Path to output validated .pptx")
    p_val.add_argument("--hold-ms", type=int, default=1200, help="Hold duration per slide in ms")
    p_val.add_argument("--max-repairs", type=int, default=3, help="Max self-healing loops (default: 3)")
    p_val.add_argument(
        "--worker",
        default="auto",
        choices=["auto", "local", "remote", "static"],
        help="Validation worker mode (auto, local, remote, static)",
    )
    p_val.add_argument("--worker-url", help="Remote Windows worker URL (overrides PMG_QA_WORKER_URL)")
    p_val.add_argument("--worker-token", help="Remote Windows worker token (overrides PMG_QA_WORKER_TOKEN)")
    p_val.add_argument("--static-only", action="store_true", help="Run OpenXML static checks only (NOT_VERIFIED)")
    p_val.add_argument("--no-video", action="store_true", help="Alias for --static-only (returns NOT_VERIFIED)")
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
    p_rev.add_argument("--telemetry", help="Actual recording_telemetry.json produced by pmg record")
    p_rev.add_argument("--hold-ms", type=int, default=1200, help="Slide hold duration in ms")
    p_rev.set_defaults(func=cmd_review)

    # 6. pmg inspect
    p_ins = sub.add_parser("inspect", help="Inspect OpenXML structure, Morph anchors, transitions, and timeline")
    p_ins.add_argument("pptx", help="Path to .pptx file")
    p_ins.add_argument("--timeline-out", help="Optional path to save expected_timeline.json")
    p_ins.add_argument("--hold-ms", type=int, default=1200, help="Slide hold duration in ms")
    p_ins.set_defaults(func=cmd_inspect)

    # 7. pmg reconstruct
    p_recon = sub.add_parser("reconstruct", help="Reconstruct editable PowerPoint slides from reference source (.pptx or .mp4)")
    p_recon.add_argument("source", help="Reference .pptx or .mp4 file")
    p_recon.add_argument("-o", "--output", required=True, help="Output reconstructed .pptx path")
    p_recon.add_argument("--spec-out", help="Optional path to save extracted PresentationMotionSpec JSON")
    p_recon.add_argument("--compare-out", help="Optional path to save comparison/reconstruction report JSON")
    p_recon.add_argument("--transfer-title", help="Optional new title text for style transfer")
    p_recon.add_argument("--max-duration-sec", type=float, default=8.0, help="Max video duration in sec to reconstruct (default: 8.0)")
    p_recon.add_argument("--max-scenes", type=int, default=2, help="Max scenes to reconstruct from video (default: 2)")
    p_recon.set_defaults(func=cmd_reconstruct)

    # 8. pmg worker
    p_wrk = sub.add_parser("worker", help="Manage the PMG Windows PowerPoint Validation Worker (doctor, serve, status)")
    w_sub = p_wrk.add_subparsers(dest="worker_action", required=True)
    w_sub.add_parser("doctor", help="Check Windows PowerPoint COM, FFmpeg, OpenCV, and runtime readiness")
    w_srv = w_sub.add_parser("serve", help="Start the PMG Windows PowerPoint Validation Worker HTTP/HTTPS server")
    w_srv.add_argument("--host", default="127.0.0.1", help="Bind host (default: 127.0.0.1)")
    w_srv.add_argument("--port", type=int, default=8765, help="Bind port (default: 8765)")
    w_srv.add_argument("--token", help="Bearer authentication token (default: PMG_QA_WORKER_TOKEN)")
    w_st = w_sub.add_parser("status", help="Check remote Windows worker status via GET /health")
    w_st.add_argument("--url", help="Remote worker URL (default: PMG_QA_WORKER_URL)")
    w_st.add_argument("--token", help="Remote worker Bearer token (default: PMG_QA_WORKER_TOKEN)")
    p_wrk.set_defaults(func=cmd_worker)

    # 9. pmg benchmark
    p_bnc = sub.add_parser("benchmark", help="Run the 24-case PMG Animation Benchmark Suite or print latest report")
    b_sub = p_bnc.add_subparsers(dest="bench_action", required=True)
    b_run = b_sub.add_parser("run", help="Execute benchmark suite cases")
    b_run.add_argument("--all", action="store_true", help="Run all 24 benchmark cases")
    b_run.add_argument("--case", type=int, help="Run a specific benchmark case ID (1..24)")
    b_run.add_argument(
        "--worker",
        default="auto",
        choices=["auto", "local", "remote", "static"],
        help="Validation worker mode (default: auto)",
    )
    b_run.add_argument("--hold-ms", type=int, default=850, help="Hold duration per slide in ms (default: 850)")
    b_sub.add_parser("report", help="Print the latest benchmark report JSON")
    p_bnc.set_defaults(func=cmd_benchmark)

    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
