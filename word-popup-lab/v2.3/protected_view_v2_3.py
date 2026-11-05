"""Real Word Protected View selection/copy controls; hosted disposable runner only.

Uses the immutable V2.1 UI helpers and synthetic/redacted fixtures. No account,
original document, global security setting, or activation state is changed.
"""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
import stat
import subprocess
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("wordlab_v21", ROOT / "v2.1" / "reproduce_v2_1.py")
lab = importlib.util.module_from_spec(spec)
spec.loader.exec_module(lab)
OUT = Path(sys.argv[1]).resolve()
OUT.mkdir(parents=True, exist_ok=True)
lab.OUT = OUT


def run_protected(word: Path, label: str, method: str = "mouse", structured: bool = False,
                  start: int = 0, length: int = 12, caret: bool = False,
                  reopen_normal: bool = False) -> dict:
    import win32clipboard
    import win32com.client
    from pywinauto import Desktop, keyboard, mouse
    path = OUT / ("synthetic-" + label + ".docx")
    (lab.make_structure_docx if structured else lab.make_docx)(path)
    os.chmod(path, stat.S_IREAD)
    proc = subprocess.Popen([str(word), "/n", str(path)])
    result = {"scenario": label, "method": method, "caretBrowsing": caret,
              "selectedSyntheticTextConfirmed": False, "dialogAfterSelection": False,
              "reopenNormal": reopen_normal}
    app = None
    try:
        print("START " + label, flush=True)
        time.sleep(20)
        keyboard.send_keys("{ESC}{ESC}")
        windows = Desktop(backend="uia").windows(process=proc.pid, class_name="OpusApp")
        if not windows:
            raise RuntimeError("No trial Word window")
        main = windows[0]
        main.set_focus()
        onboarding = lab.finish_privacy_onboarding(main, proc.pid, label)
        result["onboarding"] = onboarding
        if not onboarding["ready"]:
            raise RuntimeError("Startup onboarding blocked")
        main.maximize()
        main.minimize()
        time.sleep(1)
        app = win32com.client.GetActiveObject("Word.Application")
        if int(app.ActiveWindow.Hwnd) != int(main.handle):
            raise RuntimeError("COM instance differs from disposable test")
        original = app.ActiveDocument
        blank = app.Documents.Add()
        original.Close(0)
        pv = app.ProtectedViewWindows.Open(str(path), False, "", True)
        pv.Activate()
        pv.WindowState = 1
        blank.Close(0)
        time.sleep(3)
        result["protectedViewCount"] = int(app.ProtectedViewWindows.Count)
        result["sourceName"] = str(pv.SourceName)
        result["sourcePath"] = str(pv.SourcePath)
        result["sourcePathMatches"] = os.path.normcase(str(Path(pv.SourcePath) / pv.SourceName)) == os.path.normcase(str(path))
        result["fileReadonlyObserved"] = bool(path.stat().st_file_attributes & stat.FILE_ATTRIBUTE_READONLY)
        for key, getter in [
            ("documentReadOnly", lambda: bool(pv.Document.ReadOnly)),
            ("documentFullNameMatches", lambda: os.path.normcase(str(pv.Document.FullName)) == os.path.normcase(str(path))),
            ("viewType", lambda: int(pv.Document.ActiveWindow.View.Type)),
        ]:
            try:
                result[key] = getter()
            except Exception as exc:
                result[key + "Error"] = type(exc).__name__ + ": " + str(exc)[:200]
        if reopen_normal:
            placeholder = app.Documents.Add()
            pv.Close()
            app.Documents.Open(str(path), ReadOnly=True, AddToRecentFiles=False)
            placeholder.Close(0)
            time.sleep(3)
            result["protectedViewCountAfterReopen"] = int(app.ProtectedViewWindows.Count)
            result["documentReadOnly"] = bool(app.ActiveDocument.ReadOnly)
            result["viewType"] = int(app.ActiveWindow.View.Type)
        candidates = [w for w in Desktop(backend="uia").windows(class_name="OpusApp")
                      if path.stem in w.window_text() and w.is_visible()]
        if len(candidates) != 1:
            raise RuntimeError("Expected one document window after transition, got " + str(len(candidates)))
        main = candidates[0]
        main.restore()
        main.maximize()
        main.set_focus()
        pid = main.process_id()
        result["windowTitle"] = main.window_text()
        result["protectedBannerObserved"] = "受保护" in main.window_text() or any(
            "受保护" in e.window_text() for e in main.descendants() if e.is_visible())
        lab.emit(OUT / (label + "-opened-ui.json"), lab.inspect_windows(pid))
        lab.screenshot(label + "-opened")
        body = [e for e in main.descendants(control_type="Edit") if e.element_info.automation_id == "Body"]
        if not body:
            body = [e for e in main.descendants(control_type="Document") if e.is_visible() and e.is_enabled()]
        if not body:
            raise RuntimeError("No document text control")
        body[0].click_input()
        if caret:
            keyboard.send_keys("{F7}")
            time.sleep(1)
        before = lab.inspect_windows(pid)
        result["dialogBeforeSelection"] = lab.connected_dialog(before)
        if result["dialogBeforeSelection"]:
            raise RuntimeError("Target warning preceded selection")
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
        finally:
            win32clipboard.CloseClipboard()
        if method == "mouse":
            span = body[0].iface_text.DocumentRange.Clone()
            span.MoveEndpointByRange(1, span, 0)
            if start:
                span.MoveEndpointByUnit(0, 0, start)
            span.MoveEndpointByUnit(1, 0, length)
            bounds = list(span.GetBoundingRectangles())
            if len(bounds) < 4 or bounds[2] <= 1:
                raise RuntimeError("No nonempty real text geometry")
            x, y, width, height = [int(v) for v in bounds[:4]]
            result["selectionRect"] = [x, y, width, height]
            mouse.press(button="left", coords=(x+1, y+max(1, height//2)))
            mouse.move(coords=(x+max(2, width-1), y+max(1, height//2)))
            time.sleep(.5)
            mouse.release(button="left", coords=(x+max(2, width-1), y+max(1, height//2)))
        else:
            keyboard.send_keys("^{HOME}")
            if start:
                keyboard.send_keys("{RIGHT " + str(start) + "}")
            keyboard.send_keys("+{RIGHT " + str(length) + "}")
        time.sleep(4)
        after = lab.inspect_windows(pid)
        lab.emit(OUT / (label + "-after-selection-ui.json"), after)
        lab.screenshot(label + "-after-selection")
        result["dialogAfterSelection"] = lab.connected_dialog(after)
        if not result["dialogAfterSelection"]:
            keyboard.send_keys("^c")
            time.sleep(1)
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(13):
                    text = win32clipboard.GetClipboardData(13)
                    result["copiedSyntheticSample"] = text[:80]
                    result["selectedSyntheticTextConfirmed"] = bool(text.strip()) and lab.SYNTHETIC[start:].startswith(text.strip())
            finally:
                win32clipboard.CloseClipboard()
        result["result"] = "COMPLETED"
    except Exception as exc:
        result["result"] = "BLOCKED_AUTOMATION"
        result["error"] = type(exc).__name__ + ": " + str(exc)[:500]
        lab.screenshot(label + "-error")
    finally:
        keyboard.send_keys("{ESC}{ESC}")
        if app is not None:
            try:
                app.Quit(0)
            except Exception:
                pass
        subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True, check=False)
        time.sleep(2)
    lab.emit(OUT / (label + ".json"), result)
    print(json.dumps(result, ensure_ascii=False), flush=True)
    return result


def valid(result: dict) -> bool:
    return (result.get("result") == "COMPLETED" and
            result.get("selectedSyntheticTextConfirmed") is True and
            result.get("dialogAfterSelection") is False and
            result.get("fileReadonlyObserved") is True and
            result.get("documentReadOnly") is True and
            result.get("protectedViewCount") == 1 and
            result.get("sourcePathMatches") is True and
            result.get("protectedBannerObserved") is True and
            result.get("viewType") == 3)


def main() -> int:
    if os.name != "nt" or os.environ.get("GITHUB_ACTIONS") != "true" or os.environ.get("RUNNER_ENVIRONMENT") != "github-hosted":
        raise SystemExit("Disposable hosted Windows only")
    word = lab.find_word()
    privacy = lab.privacy_snapshot()
    baseline = {"writable": lab.run_scenario(word, False, "writable"),
                "readonly": lab.run_scenario(word, True, "readonly")}
    lab.emit(OUT / "baseline.json", baseline)
    trials = [run_protected(word, "protected-mouse-B")]
    if trials[0].get("selectedSyntheticTextConfirmed") and not trials[0].get("dialogAfterSelection"):
        trials.extend([
            run_protected(word, "protected-mouse-B2"),
            run_protected(word, "protected-structure-mouse", structured=True),
            run_protected(word, "protected-chinese-mouse", start=lab.SYNTHETIC.index("这"), length=8),
            run_protected(word, "protected-keyboard", method="keyboard"),
            run_protected(word, "protected-caret-keyboard", method="keyboard", caret=True),
            run_protected(word, "protected-caret-chinese-keyboard", method="keyboard", caret=True, start=lab.SYNTHETIC.index("这"), length=8),
        ])
    reversal = run_protected(word, "protected-then-normal-A2", reopen_normal=True)
    unchanged = privacy == lab.privacy_snapshot() and all(v == 2 for v in privacy.values())
    result = {"version": "2.3", "baseline": baseline, "trials": trials, "reversal": reversal,
              "privacyUnchangedAndDisabled": unchanged,
              "verifiedMousePageViewFallback": baseline["writable"]["selectedSyntheticTextConfirmed"] and baseline["readonly"]["dialogAfterSelection"] and
              len(trials) > 1 and all(valid(t) for t in trials[:4]) and reversal["dialogAfterSelection"] and unchanged,
              "rootCauseFixed": False}
    lab.emit(OUT / "attempt.json", result)
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
