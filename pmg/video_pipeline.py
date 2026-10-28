"""
pmg.video_pipeline
==================
Mandatory Step 3 (Native PowerPoint Video Export) and Step 4 (Actual Slideshow Recording)
with FFprobe verification, process timeout guards, and safe COM process cleanup.
"""

import os
import sys
import json
import time
import ctypes
import threading
import subprocess
from typing import Tuple, Dict, Any, Optional, List

import cv2
import numpy as np

from pmg.models import VideoMetadata, ExpectedTimeline, ValidationStatus


def is_powerpoint_available() -> bool:
    """Checks if Windows + pywin32 + Microsoft PowerPoint Desktop COM are installed and available."""
    if sys.platform != "win32":
        return False
    try:
        import winreg
        import win32com.client  # noqa: F401
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID"):
            return True
    except Exception:
        return False


def _eval_fraction(frac_str: str) -> float:
    try:
        if "/" in frac_str:
            num, den = frac_str.split("/", 1)
            d = float(den)
            return float(num) / d if d != 0 else 0.0
        return float(frac_str)
    except Exception:
        return 0.0


def probe_video_metadata(video_path: str, requested_fps: int = 60, dropped_frames: int = 0) -> VideoMetadata:
    """
    Inspects a video file using FFprobe and verifies actual decodability via OpenCV.
    Never claims 60 FPS unless measured in the actual output stream.
    """
    abs_path = os.path.abspath(video_path)
    meta = VideoMetadata(path=abs_path, requested_fps=requested_fps, dropped_frames=dropped_frames)

    if not os.path.exists(abs_path) or os.path.getsize(abs_path) == 0:
        meta.error = "Video file does not exist or is empty."
        return meta

    meta.exists = True
    meta.size_bytes = os.path.getsize(abs_path)

    # 1. Inspect with FFprobe
    try:
        cmd = [
            "ffprobe",
            "-v", "error",
            "-select_streams", "v:0",
            "-show_entries", "format=duration,size:stream=width,height,r_frame_rate,avg_frame_rate,nb_frames,codec_name",
            "-of", "json",
            abs_path,
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        if res.returncode == 0 and res.stdout.strip():
            info = json.loads(res.stdout)
            streams = info.get("streams", [])
            fmt = info.get("format", {})
            if streams:
                s0 = streams[0]
                meta.codec_name = s0.get("codec_name", "")
                meta.width = int(s0.get("width", 0))
                meta.height = int(s0.get("height", 0))
                meta.r_frame_rate = s0.get("r_frame_rate", "0/1")
                meta.avg_frame_rate = s0.get("avg_frame_rate", "0/1")
                meta.nb_frames = int(s0.get("nb_frames", 0) or 0)
            if fmt:
                meta.duration_sec = round(float(fmt.get("duration", 0.0) or 0.0), 3)
            avg_fps = _eval_fraction(meta.avg_frame_rate)
            r_fps = _eval_fraction(meta.r_frame_rate)
            meta.actual_fps = round(avg_fps if avg_fps > 0 else r_fps, 2)
    except Exception as e:
        meta.error = f"ffprobe error: {str(e)}"

    # 2. Verify actual frame decoding with OpenCV
    try:
        cap = cv2.VideoCapture(abs_path)
        if cap.isOpened():
            ret, frame = cap.read()
            if ret and frame is not None and frame.size > 0:
                meta.decodable = True
                if meta.width == 0:
                    meta.width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                    meta.height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                if meta.nb_frames == 0:
                    meta.nb_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                if meta.actual_fps == 0.0:
                    meta.actual_fps = round(float(cap.get(cv2.CAP_PROP_FPS)), 2)
        cap.release()
    except Exception as e:
        meta.decodable = False
        meta.error = f"OpenCV decode error: {str(e)}"

    return meta


def export_video_with_powerpoint(
    pptx_path: str,
    output_video_path: str,
    slide_duration: float = 2.0,
    requested_fps: int = 60,
    resolution: int = 1080,
    timeout_sec: int = 180,
) -> VideoMetadata:
    """
    STEP 3 — Native PowerPoint Video Export.
    Opens the PPTX in Microsoft PowerPoint Desktop and exports an MP4 video,
    then verifies and records actual FPS, resolution, and duration via FFprobe.
    """
    abs_pptx = os.path.abspath(pptx_path)
    abs_video = os.path.abspath(output_video_path)
    os.makedirs(os.path.dirname(abs_video), exist_ok=True)

    if os.path.exists(abs_video):
        try:
            os.remove(abs_video)
        except Exception:
            pass

    if sys.platform != "win32":
        meta = VideoMetadata(path=abs_video, requested_fps=requested_fps)
        meta.error = "BLOCKED: Native PowerPoint export requires Windows + Microsoft PowerPoint."
        return meta

    import win32com.client

    ppt = None
    deck = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        try:
            ppt.Visible = True
        except Exception:
            pass

        deck = ppt.Presentations.Open(abs_pptx, ReadOnly=True, WithWindow=True)
        default_dur = max(1, int(round(slide_duration)))
        deck.CreateVideo(abs_video, -1, default_dur, resolution, requested_fps, 90)

        time.sleep(1.0)
        start_time = time.time()
        while True:
            status = deck.CreateVideoStatus
            # 1 = ppMediaTaskStatusInProgress, 2 = ppMediaTaskStatusQueued, 3 = Done, 4 = Failed
            if status not in (1, 2):
                break
            if time.time() - start_time > timeout_sec:
                raise TimeoutError(f"PowerPoint CreateVideo timed out after {timeout_sec}s.")
            time.sleep(0.5)

        time.sleep(1.0)
    except Exception as e:
        meta = VideoMetadata(path=abs_video, requested_fps=requested_fps)
        meta.error = f"PowerPoint export failed: {str(e)}"
        return meta
    finally:
        if deck is not None:
            try:
                deck.Close()
            except Exception:
                pass
        if ppt is not None:
            try:
                if ppt.Presentations.Count == 0 and ppt.SlideShowWindows.Count == 0:
                    ppt.Quit()
            except Exception:
                pass

    return probe_video_metadata(abs_video, requested_fps=requested_fps)


def _find_slideshow_hwnd() -> int:
    """Finds the visible PowerPoint full-screen Slide Show window ('screenClass')."""
    import win32gui

    found_hwnd = 0

    def _cb(h, _):
        nonlocal found_hwnd
        if win32gui.IsWindowVisible(h):
            cls = win32gui.GetClassName(h)
            if cls == "screenClass":
                found_hwnd = h

    win32gui.EnumWindows(_cb, None)
    return found_hwnd


def record_actual_slideshow(
    pptx_path: str,
    output_video_path: str,
    expected_timeline: ExpectedTimeline,
    target_fps: int = 30,
    timeout_sec: int = 240,
) -> Tuple[VideoMetadata, Dict[str, Any]]:
    """
    STEP 4 — Record Actual Slideshow Playback.
    Launches Microsoft PowerPoint Desktop in full-screen Slide Show mode (no editing UI or toolbars),
    steps through the presentation according to ExpectedTimeline's playback script,
    and records the live slideshow window into an MP4 video.
    """
    abs_pptx = os.path.abspath(pptx_path)
    abs_video = os.path.abspath(output_video_path)
    os.makedirs(os.path.dirname(abs_video), exist_ok=True)

    telemetry: Dict[str, Any] = {
        "capture_method": "win32_printwindow_fullscreen_slideshow_ffmpeg",
        "target_fps": target_fps,
        "captured_unique_frames": 0,
        "duplicated_frames": 0,
        "dropped_frames": 0,
        "slide_events": [],
        "click_events": [],
    }

    if sys.platform != "win32":
        meta = VideoMetadata(path=abs_video, requested_fps=target_fps)
        meta.error = "BLOCKED: Actual PowerPoint slideshow recording requires Windows + Microsoft PowerPoint."
        return meta, telemetry

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

    import win32com.client
    import win32gui
    import win32ui

    if os.path.exists(abs_video):
        try:
            os.remove(abs_video)
        except Exception:
            pass

    ppt = None
    pres = None
    ss_win = None
    ffmpeg_proc = None
    stop_event = threading.Event()

    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = True
        pres = ppt.Presentations.Open(abs_pptx, ReadOnly=True, WithWindow=True)
        # ShowType = 1 (ppShowTypeSpeaker full screen, zero toolbars/editing panes)
        pres.SlideShowSettings.ShowType = 1
        pres.SlideShowSettings.AdvanceMode = 1  # ppSlideShowManualAdvance
        ss_win = pres.SlideShowSettings.Run()
        try:
            ss_win.Activate()
        except Exception:
            pass

        # Wait for screenClass window to appear
        hwnd = 0
        for _ in range(30):
            hwnd = _find_slideshow_hwnd()
            if hwnd:
                break
            time.sleep(0.1)

        if not hwnd:
            raise RuntimeError("Could not locate PowerPoint Slide Show window (screenClass).")

        time.sleep(0.35)

        l, t, r, b = win32gui.GetWindowRect(hwnd)
        w = (r - l) - ((r - l) % 2)
        h = (b - t) - ((b - t) % 2)
        telemetry["capture_width"] = w
        telemetry["capture_height"] = h

        ffmpeg_cmd = [
            "ffmpeg",
            "-y",
            "-f", "rawvideo",
            "-pix_fmt", "bgr24",
            "-s", f"{w}x{h}",
            "-r", str(target_fps),
            "-i", "-",
            "-c:v", "libx264",
            "-preset", "ultrafast",
            "-crf", "18",
            "-pix_fmt", "yuv420p",
            abs_video,
        ]
        ffmpeg_proc = subprocess.Popen(
            ffmpeg_cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

        # Background frame grabber thread synchronized to wall-clock time
        rec_start_perf = time.perf_counter()

        def _capture_loop():
            hwndDC = win32gui.GetWindowDC(hwnd)
            mfcDC = win32ui.CreateDCFromHandle(hwndDC)
            saveDC = mfcDC.CreateCompatibleDC()
            bmp = win32ui.CreateBitmap()
            bmp.CreateCompatibleBitmap(mfcDC, w, h)
            saveDC.SelectObject(bmp)

            frames_written = 0
            unique_frames = 0
            dup_frames = 0
            last_bgr_bytes = None

            try:
                while not stop_event.is_set():
                    t_now = time.perf_counter() - rec_start_perf
                    target_frame_idx = int(t_now * target_fps)

                    if frames_written <= target_frame_idx:
                        res = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)
                        if res == 1:
                            bmp_bytes = bmp.GetBitmapBits(True)
                            arr = np.frombuffer(bmp_bytes, dtype=np.uint8).reshape((h, w, 4))
                            bgr = np.ascontiguousarray(arr[:, :, :3])
                            last_bgr_bytes = bgr.tobytes()
                            unique_frames += 1

                        if last_bgr_bytes is not None and ffmpeg_proc.stdin:
                            catchup = max(1, target_frame_idx - frames_written + 1)
                            for c_i in range(catchup):
                                ffmpeg_proc.stdin.write(last_bgr_bytes)
                                frames_written += 1
                                if c_i > 0:
                                    dup_frames += 1
                    else:
                        time.sleep(0.004)
            except Exception:
                pass
            finally:
                telemetry["captured_unique_frames"] = unique_frames
                telemetry["duplicated_frames"] = dup_frames
                telemetry["dropped_frames"] = dup_frames
                telemetry["total_frames_written"] = frames_written
                try:
                    win32gui.DeleteObject(bmp.GetHandle())
                    saveDC.DeleteDC()
                    mfcDC.DeleteDC()
                    win32gui.ReleaseDC(hwnd, hwndDC)
                except Exception:
                    pass

        grab_thread = threading.Thread(target=_capture_loop, daemon=True)
        grab_thread.start()

        # Execute the expected timeline playback script
        for s_idx, slide_spec in enumerate(expected_timeline.slides, 1):
            ev_start = round((time.perf_counter() - rec_start_perf) * 1000.0, 1)
            if s_idx == 1:
                has_intra = any(o.trigger_type not in ("static", "slide_transition") for o in slide_spec.objects)
                if has_intra:
                    try:
                        # Reset Slide 1 animations right after grab_thread starts so t=0 is captured
                        ss_win.View.GotoSlide(1, -1)
                    except Exception:
                        pass
            else:
                ss_win.View.Next()

            # Wait for transition duration
            trans_wait_sec = (slide_spec.transition_duration_ms / 1000.0) if s_idx > 1 else 0.0
            if trans_wait_sec > 0:
                time.sleep(trans_wait_sec)

            elapsed_on_slide_ms = slide_spec.transition_duration_ms if s_idx > 1 else 0

            # Execute any intra-slide click triggers according to expected slide_relative_ms
            for click_ev in slide_spec.clicks:
                target_rel_ms = click_ev.get("slide_relative_ms", elapsed_on_slide_ms + 500)
                wait_ms = max(100, target_rel_ms - elapsed_on_slide_ms)
                time.sleep(wait_ms / 1000.0)
                click_actual_ms = round((time.perf_counter() - rec_start_perf) * 1000.0, 1)
                ss_win.View.Next()
                telemetry["click_events"].append({
                    "slide_number": s_idx,
                    "sequence_index": click_ev.get("sequence_index", 0),
                    "target_object": click_ev.get("target_object", ""),
                    "expected_slide_relative_ms": target_rel_ms,
                    "actual_ms": click_actual_ms,
                })
                elapsed_on_slide_ms = target_rel_ms

            # Wait for any remaining automatic or click-triggered animations on the slide
            expected_slide_span_ms = max(
                0, slide_spec.absolute_end_ms - slide_spec.absolute_start_ms - slide_spec.hold_duration_ms
            )
            rem_anim_ms = max(0, expected_slide_span_ms - elapsed_on_slide_ms)
            if rem_anim_ms > 0:
                time.sleep(rem_anim_ms / 1000.0)

            # Hold steady state on slide
            hold_sec = max(0.8, slide_spec.hold_duration_ms / 1000.0)
            time.sleep(hold_sec)

            ev_end = round((time.perf_counter() - rec_start_perf) * 1000.0, 1)
            telemetry["slide_events"].append({
                "slide_number": s_idx,
                "transition_type": slide_spec.transition_type,
                "transition_duration_ms": slide_spec.transition_duration_ms,
                "actual_start_ms": ev_start,
                "actual_end_ms": ev_end,
            })

            if (time.perf_counter() - rec_start_perf) > timeout_sec:
                break

        stop_event.set()
        grab_thread.join(timeout=5.0)

        if ffmpeg_proc and ffmpeg_proc.stdin:
            try:
                ffmpeg_proc.stdin.close()
            except Exception:
                pass
            ffmpeg_proc.wait(timeout=15.0)

    except Exception as e:
        stop_event.set()
        meta = VideoMetadata(path=abs_video, requested_fps=target_fps)
        meta.error = f"Slideshow recording failed: {str(e)}"
        return meta, telemetry
    finally:
        stop_event.set()
        if ffmpeg_proc and ffmpeg_proc.poll() is None:
            try:
                ffmpeg_proc.kill()
            except Exception:
                pass
        if ss_win is not None:
            try:
                ss_win.View.Exit()
            except Exception:
                pass
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if ppt is not None:
            try:
                if ppt.Presentations.Count == 0 and ppt.SlideShowWindows.Count == 0:
                    ppt.Quit()
            except Exception:
                pass

    meta = probe_video_metadata(
        abs_video,
        requested_fps=target_fps,
        dropped_frames=telemetry.get("dropped_frames", 0),
    )
    return meta, telemetry
