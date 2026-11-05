"""End-to-end portable PowerShell launcher controls; hosted runner only.

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


def run_launcher(word: Path, label: str, method: str = "mouse", structured: bool = False,
                  start: int = 0, length: int = 12) -> dict:
    import win32clipboard
    import win32com.client
    from pywinauto import Desktop, keyboard, mouse
    path = OUT / ("synthetic-" + label + ".docx")
    (lab.make_structure_docx if structured else lab.make_docx)(path)
    os.chmod(path, stat.S_IREAD)
    launcher = OUT / "launcher_package" / "Word_ReadOnly_Open_V1_1.ps1"
    metadata_path = OUT / (label + "-launcher.json")
    proc = None
    word_pid = None
    result = {"scenario": label, "method": method, "selectedSyntheticTextConfirmed": False,
              "dialogAfterSelection": False, "launcherSuccess": False}
    app = None
    try:
        print("START " + label, flush=True)
        proc = subprocess.Popen(["powershell.exe", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass",
                                 "-File", str(launcher), "-FilePath", str(path), "-ResultJson", str(metadata_path)],
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = proc.communicate(timeout=90)
        result["launcherExitCode"] = proc.returncode
        result["launcherExitedBeforeSelection"] = proc.poll() is not None
        if not metadata_path.exists():
            raise RuntimeError("No launcher result: " + stderr.decode("utf-8", errors="replace")[-500:])
        metadata = json.loads(metadata_path.read_text(encoding="utf-8-sig"))
        result["launcherMetadata"] = metadata
        result["launcherSuccess"] = metadata.get("success") is True
        if proc.returncode != 0 or not result["launcherSuccess"]:
            raise RuntimeError("Launcher failed: " + str(metadata.get("error", "")))
        word_pid = int(metadata["wordProcessId"])
        time.sleep(5)
        windows = Desktop(backend="uia").windows(process=word_pid, class_name="OpusApp")
        if not windows:
            raise RuntimeError("Word exited after launcher process completed")
        main = windows[0]
        main.maximize()
        main.minimize()
        time.sleep(1)
        app = win32com.client.GetActiveObject("Word.Application")
        if int(app.ActiveWindow.Hwnd) != int(metadata["wordWindowHandle"]):
            raise RuntimeError("Word COM window differs from launcher result")
        result["protectedViewCount"] = int(app.ProtectedViewWindows.Count)
        result["fileReadonlyObserved"] = bool(path.stat().st_file_attributes & stat.FILE_ATTRIBUTE_READONLY)
        result["documentReadOnly"] = bool(app.ActiveDocument.ReadOnly)
        result["documentFullNameMatches"] = os.path.normcase(str(app.ActiveDocument.FullName)) == os.path.normcase(str(path))
        result["viewType"] = int(app.ActiveWindow.View.Type)
        main.restore()
        main.maximize()
        main.set_focus()
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
            time.sleep(.25)
            for step in range(1, 9):
                mouse.move(coords=(x+1+max(1, width-2)*step//8, y+max(1, height//2)))
                time.sleep(.06)
            time.sleep(.25)
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
        if proc is not None and proc.poll() is None:
            proc.kill()
        if word_pid is not None:
            subprocess.run(["taskkill", "/PID", str(word_pid), "/T", "/F"], capture_output=True, check=False)
        time.sleep(2)
    lab.emit(OUT / (label + ".json"), result)
    print(json.dumps(result, ensure_ascii=False), flush=True)
    return result


def valid(result: dict) -> bool:
    return (result.get("result") == "COMPLETED" and
            result.get("launcherSuccess") is True and result.get("launcherExitedBeforeSelection") is True and
            result.get("selectedSyntheticTextConfirmed") is True and
            result.get("dialogAfterSelection") is False and
            result.get("fileReadonlyObserved") is True and
            result.get("documentReadOnly") is True and
            result.get("protectedViewCount") == 0 and
            result.get("documentFullNameMatches") is True and
            result.get("viewType") == 3)


def main() -> int:
    if os.name != "nt" or os.environ.get("GITHUB_ACTIONS") != "true" or os.environ.get("RUNNER_ENVIRONMENT") != "github-hosted":
        raise SystemExit("Disposable hosted Windows only")
    import zipfile
    with zipfile.ZipFile(ROOT / "solution" / "v1.1" / "Word_ReadOnly_Open_V1_1.zip") as archive:
        archive.extractall(OUT / "launcher_package")
    word = lab.find_word()
    privacy = lab.privacy_snapshot()
    baseline = {"writable": lab.run_scenario(word, False, "writable"),
                "readonly": lab.run_scenario(word, True, "readonly")}
    lab.emit(OUT / "baseline.json", baseline)
    chinese = lab.SYNTHETIC.index("这")
    trials = [run_launcher(word, "launcher-keyboard-B", method="keyboard")]
    if valid(trials[0]):
        trials.extend([
            run_launcher(word, "launcher-keyboard-B2", method="keyboard"),
            run_launcher(word, "launcher-mouse"),
            run_launcher(word, "launcher-structured-mouse", structured=True),
            run_launcher(word, "launcher-chinese-keyboard", method="keyboard", start=chinese, length=8),
            run_launcher(word, "launcher-chinese-mouse", start=chinese, length=8),
        ])
    reversal = lab.run_scenario(word, True, "normal-launch-A2")
    unchanged = privacy == lab.privacy_snapshot() and all(v == 2 for v in privacy.values())
    positive = baseline["writable"]["selectedSyntheticTextConfirmed"] and baseline["readonly"]["dialogAfterSelection"]
    result = {"version": "2.8", "baseline": baseline, "trials": trials, "reversal": reversal,
              "privacyUnchangedAndDisabled": unchanged,
              "verifiedPortableLauncher": positive and len(trials) == 6 and all(valid(t) for t in trials) and reversal["dialogAfterSelection"] and unchanged,
              "rootCauseFixed": False}
    lab.emit(OUT / "attempt.json", result)
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
