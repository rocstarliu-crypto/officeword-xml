"""Validate the actual native Open split-button menu, not a launcher prototype."""
from __future__ import annotations
import importlib.util
import json
import os
from pathlib import Path
import re
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


def copy_sample() -> str:
    import win32clipboard
    from pywinauto import keyboard
    keyboard.send_keys("^c")
    time.sleep(1)
    win32clipboard.OpenClipboard()
    try:
        return win32clipboard.GetClipboardData(13) if win32clipboard.IsClipboardFormatAvailable(13) else ""
    finally:
        win32clipboard.CloseClipboard()


def run_ui(word: Path, label: str, mode: str, method: str = "keyboard", structured: bool = False, chinese: bool = False) -> dict:
    import win32clipboard
    import win32com.client
    from pywinauto import Desktop, keyboard, mouse
    target = OUT / ("synthetic-" + label + ".docx")
    startup = OUT / ("startup-writable-" + label + ".docx")
    (lab.make_structure_docx if structured else lab.make_docx)(target)
    lab.make_docx(startup)
    os.chmod(target, stat.S_IREAD)
    proc = subprocess.Popen([str(word), "/n", str(startup)])
    result = {"scenario": label, "requestedMenu": mode, "method": method,
              "selectedSyntheticTextConfirmed": False, "dialogAfterSelection": False}
    app = None
    try:
        print("START " + label, flush=True)
        time.sleep(20)
        keyboard.send_keys("{ESC}{ESC}")
        main = Desktop(backend="uia").windows(process=proc.pid, class_name="OpusApp")[0]
        main.maximize(); main.set_focus()
        if not lab.finish_privacy_onboarding(main, proc.pid, label)["ready"]:
            raise RuntimeError("Startup onboarding blocks native UI")
        body = [e for e in main.descendants(control_type="Edit") if e.element_info.automation_id == "Body"][0]
        body.click_input()
        keyboard.send_keys("^{HOME}+{RIGHT 12}")
        result["startupWritableCopy"] = copy_sample().strip() == lab.SYNTHETIC[:12]
        if not result["startupWritableCopy"]:
            raise RuntimeError("Startup window is not usable before Open dialog")
        main.minimize(); time.sleep(1)
        app = win32com.client.GetActiveObject("Word.Application")
        if int(app.ActiveWindow.Hwnd) != int(main.handle):
            raise RuntimeError("Wrong Word COM instance")
        main.restore(); main.maximize(); main.set_focus()
        keyboard.send_keys("^{F12}")
        time.sleep(2)
        dialogs = []
        for host in Desktop(backend="uia").windows(process=proc.pid):
            dialogs.extend(w for w in [host] + host.descendants(control_type="Window")
                           if w.is_visible() and re.fullmatch(r"Open|打开", w.window_text()))
        lab.emit(OUT / (label + "-file-dialog-ui.json"), lab.inspect_windows(proc.pid))
        if not dialogs:
            raise RuntimeError("Classic Open dialog was not found")
        dialog = dialogs[0]
        dialog.set_focus()
        edits = [e for e in dialog.descendants(control_type="Edit")
                 if e.element_info.automation_id in ("1148", "FileNameControlHost") or re.search(r"File name|文件名", e.window_text(), re.I)]
        if edits:
            try:
                edits[0].set_edit_text(str(target))
            except Exception:
                dialog.set_focus(); keyboard.send_keys("%n^a")
                keyboard.send_keys(str(target), with_spaces=True)
        else:
            keyboard.send_keys("%n^a")
            keyboard.send_keys(str(target), with_spaces=True)
        time.sleep(1)
        buttons = [e for e in dialog.descendants() if e.element_info.control_type in ("Button", "SplitButton")
                   and e.is_visible() and e.is_enabled() and re.fullmatch(r"(?:Open|打开)(?:\(&?[A-Za-z]\))?", e.window_text())]
        if not buttons:
            raise RuntimeError("Open split button was not found")
        button = sorted(buttons, key=lambda e: e.rectangle().bottom)[-1]
        rect = button.rectangle()
        mouse.click(button="left", coords=(rect.right-7, (rect.top+rect.bottom)//2))
        time.sleep(1)
        lab.emit(OUT / (label + "-open-menu-ui.json"), lab.inspect_windows(proc.pid))
        lab.screenshot(label + "-open-menu")
        regex = r"受保护.*视图.*打开|Open in Protected View" if mode == "protected" else r"只读.*打开|Open Read.?Only"
        choices = []
        names = []
        for w in Desktop(backend="uia").windows(process=proc.pid):
            for e in w.descendants():
                if e.is_visible() and e.is_enabled():
                    name = e.window_text()
                    if name:
                        names.append(name[:100])
                    if re.search(regex, name, re.I):
                        choices.append(e)
        result["menuNames"] = list(dict.fromkeys(names))[-100:]
        if not choices:
            raise RuntimeError("Requested native menu option was not found")
        result["chosenMenuText"] = choices[0].window_text()
        choices[0].click_input()
        time.sleep(5)
        result["protectedViewCount"] = int(app.ProtectedViewWindows.Count)
        if mode == "protected":
            pv = app.ProtectedViewWindows.Item(app.ProtectedViewWindows.Count)
            document = pv.Document
            result["sourcePathMatches"] = os.path.normcase(str(Path(pv.SourcePath) / pv.SourceName)) == os.path.normcase(str(target))
        else:
            document = app.ActiveDocument
            result["sourcePathMatches"] = os.path.normcase(str(document.FullName)) == os.path.normcase(str(target))
        result["fileReadonlyObserved"] = bool(target.stat().st_file_attributes & stat.FILE_ATTRIBUTE_READONLY)
        result["documentReadOnly"] = bool(document.ReadOnly)
        result["documentFullNameMatches"] = os.path.normcase(str(document.FullName)) == os.path.normcase(str(target))
        result["viewType"] = int(document.ActiveWindow.View.Type)
        windows = [w for w in Desktop(backend="uia").windows(process=proc.pid, class_name="OpusApp") if target.stem in w.window_text()]
        if len(windows) != 1:
            raise RuntimeError("Expected the exact target Word window")
        main = windows[0]
        main.restore(); main.maximize(); main.set_focus()
        body = [e for e in main.descendants(control_type="Edit") if e.element_info.automation_id == "Body"][0]
        body.click_input()
        before = lab.inspect_windows(proc.pid)
        result["dialogBeforeSelection"] = lab.connected_dialog(before)
        if result["dialogBeforeSelection"]:
            raise RuntimeError("Target warning preceded actual test selection")
        win32clipboard.OpenClipboard()
        try: win32clipboard.EmptyClipboard()
        finally: win32clipboard.CloseClipboard()
        start = lab.SYNTHETIC.index("这") if chinese else 0
        length = 8 if chinese else 12
        if method == "mouse":
            span = body.iface_text.DocumentRange.Clone()
            span.MoveEndpointByRange(1, span, 0)
            if start: span.MoveEndpointByUnit(0, 0, start)
            span.MoveEndpointByUnit(1, 0, length)
            bounds = list(span.GetBoundingRectangles())
            if len(bounds) < 4 or bounds[2] <= 1: raise RuntimeError("No target text geometry")
            x,y,width,height = [int(v) for v in bounds[:4]]
            mouse.press(button="left", coords=(x+1,y+height//2));time.sleep(.25)
            for step in range(1,9):
                mouse.move(coords=(x+1+max(1,width-2)*step//8,y+height//2));time.sleep(.06)
            time.sleep(.25);mouse.release(button="left",coords=(x+max(2,width-1),y+height//2))
        else:
            keyboard.send_keys("^{HOME}")
            if start: keyboard.send_keys("{RIGHT " + str(start) + "}")
            keyboard.send_keys("+{RIGHT " + str(length) + "}")
        time.sleep(4)
        after = lab.inspect_windows(proc.pid)
        result["dialogAfterSelection"] = lab.connected_dialog(after)
        lab.emit(OUT / (label + "-after-selection-ui.json"), after)
        lab.screenshot(label + "-after-selection")
        if not result["dialogAfterSelection"]:
            sample = copy_sample()
            result["copiedSyntheticSample"] = sample[:80]
            result["selectedSyntheticTextConfirmed"] = bool(sample.strip()) and lab.SYNTHETIC[start:].startswith(sample.strip())
        result["result"] = "COMPLETED"
    except Exception as exc:
        result["result"] = "BLOCKED_AUTOMATION"
        result["error"] = type(exc).__name__ + ": " + str(exc)[:500]
        lab.screenshot(label + "-error")
    finally:
        keyboard.send_keys("{ESC}{ESC}")
        if app is not None:
            try: app.Quit(0)
            except Exception: pass
        subprocess.run(["taskkill","/PID",str(proc.pid),"/T","/F"],capture_output=True,check=False)
        time.sleep(2)
    lab.emit(OUT / (label + ".json"), result)
    print(json.dumps(result, ensure_ascii=False),flush=True)
    return result


def valid(r: dict) -> bool:
    return (r.get("result") == "COMPLETED" and r.get("selectedSyntheticTextConfirmed") is True and
            r.get("dialogAfterSelection") is False and r.get("fileReadonlyObserved") is True and
            r.get("documentReadOnly") is True and r.get("documentFullNameMatches") is True and
            r.get("sourcePathMatches") is True and r.get("viewType") == 3)


def main() -> int:
    if os.name != "nt" or os.environ.get("GITHUB_ACTIONS") != "true" or os.environ.get("RUNNER_ENVIRONMENT") != "github-hosted":
        raise SystemExit("Disposable hosted Windows only")
    word = lab.find_word(); privacy = lab.privacy_snapshot()
    baseline = {"writable":lab.run_scenario(word,False,"baseline-writable"),"readonly":lab.run_scenario(word,True,"baseline-readonly")}
    trials = [run_ui(word,"native-menu-readonly","readonly"),
              run_ui(word,"native-menu-protected-keyboard","protected"),
              run_ui(word,"native-menu-protected-mouse","protected",method="mouse"),
              run_ui(word,"native-menu-protected-structure","protected",method="mouse",structured=True),
              run_ui(word,"native-menu-protected-chinese","protected",method="mouse",chinese=True)]
    if valid(trials[0]):
        trials.extend([run_ui(word,"native-menu-readonly-repeat","readonly"),
                       run_ui(word,"native-menu-readonly-chinese","readonly",method="mouse",chinese=True),
                       run_ui(word,"native-menu-readonly-structure","readonly",method="mouse",structured=True)])
    unchanged = privacy == lab.privacy_snapshot() and all(v==2 for v in privacy.values())
    positive = baseline["writable"]["selectedSyntheticTextConfirmed"] and baseline["readonly"]["dialogAfterSelection"]
    result = {"version":"3.1","baseline":baseline,"trials":trials,"privacyUnchangedAndDisabled":unchanged,
              "verifiedNativeProtectedMenu":positive and all(valid(t) and t.get("protectedViewCount")==1 for t in trials[1:5]) and unchanged,
              "verifiedNativeReadonlyMenu":positive and len(trials)==8 and all(valid(trials[i]) and trials[i].get("protectedViewCount")==0 for i in (0,5,6,7)) and unchanged,
              "rootCauseFixed":False}
    lab.emit(OUT/"attempt.json",result);print(json.dumps(result,ensure_ascii=False,indent=2),flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
