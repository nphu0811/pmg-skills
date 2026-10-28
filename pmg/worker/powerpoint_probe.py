"""Bounded, read-only PowerPoint compatibility probe. Never requests repair.

The subprocess isolates a potentially modal Presentations.Open call. Alerts stay
enabled; dialog/Protected View observations, animation counts and source hashes
are evidence, not an independent delivery certificate.
"""
import csv
import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Dict

PP_ALERTS_ALL = 2  # PpAlertLevel.ppAlertsAll; 1 is ppAlertsNone.


def file_sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, 'rb') as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _powerpoint_pids():
    result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq POWERPNT.EXE', '/FO', 'CSV', '/NH'],
                            capture_output=True, text=True, timeout=5,
                            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    return {int(row[1]) for row in csv.reader(io.StringIO(result.stdout))
            if len(row) > 1 and row[0].lower() == 'powerpnt.exe' and row[1].isdigit()}


def _probe_in_process(pptx_path: str, state_path: str) -> Dict[str, Any]:
    import pythoncom
    import win32com.client
    import win32gui
    import win32process

    previous_pids = _powerpoint_pids()
    state: Dict[str, Any] = {'is_clean': False, 'skipped': False, 'display_alerts': PP_ALERTS_ALL,
                             'interruptions': [], 'owned_process': False}
    state_lock = threading.Lock()
    stop = threading.Event()

    def save_state():
        with state_lock:
            temporary = state_path + '.tmp'
            Path(temporary).write_text(json.dumps(state), encoding='utf-8')
            os.replace(temporary, state_path)

    def watch_dialogs():
        while not stop.wait(0.05):
            def check(hwnd, _):
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if win32process.GetWindowThreadProcessId(hwnd)[1] not in state.get('powerpoint_pids', []):
                    return
                title = win32gui.GetWindowText(hwnd)
                cls = win32gui.GetClassName(hwnd)
                if cls == '#32770' or (cls != 'PPTFrameClass' and any(
                    word in title.lower() for word in ('repair', 'protected view', 'compatibility', 'problem with content')
                )):
                    text_parts = [title]
                    win32gui.EnumChildWindows(hwnd, lambda h, _: text_parts.append(win32gui.GetWindowText(h)), None)
                    observation = {'class': cls, 'text': ' '.join(t for t in text_parts if t)}
                    if observation not in state['interruptions']:
                        state['interruptions'].append(observation)
                        save_state()
            try:
                win32gui.EnumWindows(check, None)
            except Exception:
                pass

    pythoncom.CoInitialize()
    app = presentation = None
    try:
        app = win32com.client.DispatchEx('PowerPoint.Application')
        app.DisplayAlerts = PP_ALERTS_ALL
        current_pids = _powerpoint_pids()
        created_pids = current_pids - previous_pids
        state['powerpoint_pids'] = sorted(current_pids)
        state['owned_process'] = len(created_pids) == 1
        if state['owned_process']:
            state['powerpoint_pid'] = next(iter(created_pids))
        save_state()
        monitor = threading.Thread(target=watch_dialogs, daemon=True)
        monitor.start()
        # Open2007 explicitly disables Microsoft's automatic repair procedure.
        presentation = app.Presentations.Open2007(
            os.path.abspath(pptx_path), ReadOnly=True, Untitled=False,
            WithWindow=False, OpenAndRepair=False)
        state['slide_count'] = int(presentation.Slides.Count)
        state['saved'] = bool(presentation.Saved)
        state['protected_view_count'] = int(app.ProtectedViewWindows.Count)
        state['animation_counts'] = {}
        for index in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(index)
            sequences = slide.TimeLine
            count = int(sequences.MainSequence.Count)
            for sequence in sequences.InteractiveSequences:
                count += int(sequence.Count)
            state['animation_counts'][str(index)] = count
        state['is_clean'] = (state['saved'] and not state['interruptions']
                             and state['protected_view_count'] == 0)
        if not state['is_clean']:
            state['error'] = 'PowerPoint required repair, changed the in-memory deck, or interrupted opening.'
    except Exception as exc:
        state['error'] = f'PowerPoint clean-open failed: {exc}'
        state['traceback'] = traceback.format_exc()
    finally:
        if presentation is not None:
            presentation.Close()  # Read-only; never Save/SaveAs.
        if app is not None and state['owned_process']:
            if app.Presentations.Count == 0:
                app.Quit()
        stop.set()
        pythoncom.CoUninitialize()
    save_state()
    return state


def verify_clean_powerpoint_open(pptx_path: str, timeout_sec: float = 30.0) -> Dict[str, Any]:
    """Return honest compatibility evidence; unavailable/timeout never means clean."""
    if not os.path.isfile(pptx_path):
        return {'is_clean': False, 'skipped': False, 'error': f'Presentation file does not exist: {pptx_path}'}
    if sys.platform != 'win32':
        return {'is_clean': False, 'skipped': True, 'status': 'BLOCKED', 'error': 'Windows PowerPoint required.'}
    before = file_sha256(pptx_path)
    with tempfile.TemporaryDirectory(prefix='pmg_open_probe_') as temp:
        state_path = os.path.join(temp, 'result.json')
        command = [sys.executable, '-m', 'pmg.worker.powerpoint_probe', os.path.abspath(pptx_path), state_path]
        try:
            process = subprocess.run(command, capture_output=True, text=True, timeout=timeout_sec,
                                     creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            state = json.loads(Path(state_path).read_text(encoding='utf-8')) if os.path.exists(state_path) else {}
            if process.returncode != 0 or not state:
                state.update(is_clean=False, skipped=False, error=process.stderr.strip() or 'PowerPoint probe failed.')
        except subprocess.TimeoutExpired:
            state = json.loads(Path(state_path).read_text(encoding='utf-8')) if os.path.exists(state_path) else {}
            state.update(is_clean=False, skipped=False, timed_out=True,
                         error=f'PowerPoint clean-open timed out after {timeout_sec}s; possible modal dialog.')
            # Only terminate a newly created probe-owned PowerPoint process.
            if state.get('owned_process') and state.get('powerpoint_pid') in _powerpoint_pids():
                import win32api
                handle = win32api.OpenProcess(1, False, state['powerpoint_pid'])
                win32api.TerminateProcess(handle, 1)
                handle.Close()
        after = file_sha256(pptx_path)
        state.update(sha256=after, source_unchanged=(before == after))
        if before != after:
            state.update(is_clean=False, error='PowerPoint probe changed the source PPTX.')
        return state


if __name__ == '__main__':
    _probe_in_process(sys.argv[1], sys.argv[2])
