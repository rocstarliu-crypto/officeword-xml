"""Synthetic Word UI reproduction attempt; no account or user documents.

Flow: create synthetic documents -> open real Word -> record startup UI ->
select text with keyboard -> detect exact connected-experiences dialog.
No-popup results in an unsigned-in installation are inconclusive, never fixed.
"""
from __future__ import annotations

import ctypes
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import time
import zipfile
import winreg
import csv

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SYNTHETIC = "WORDLAB_SYNTHETIC_TEXT_1234567890 这是只读选字测试文档，不包含个人内容。"
OUT = Path(sys.argv[1] if len(sys.argv) > 1 else "lab-results").resolve()
OUT.mkdir(parents=True, exist_ok=True)


def emit(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def make_docx(path: Path) -> None:
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/document.xml": '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' + SYNTHETIC + '</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
    }
    parts["word/_rels/document.xml.rels"] = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'
    parts["word/settings.xml"] = '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content.encode("utf-8"))


def find_word() -> Path | None:
    for root in [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]:
        if root:
            for rel in ["Microsoft Office/root/Office16/WINWORD.EXE", "Microsoft Office/Office16/WINWORD.EXE"]:
                path = Path(root) / rel
                if path.is_file():
                    return path
    return None


def scan_flag_identifiers(word: Path) -> list[dict]:
    """Collect identifiers only, not proprietary source excerpts or account data."""
    hits = []
    root = word.parent / "sdxs"
    if not root.is_dir():
        return hits
    pattern = re.compile(r"Microsoft\.Office\.[A-Za-z0-9_.]{1,140}")
    for path in root.rglob("*"):
        if path.suffix.lower() not in {".js", ".json", ".xml", ".html"} or path.stat().st_size > 15_000_000:
            continue
        try:
            data = path.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        flags = sorted({s for s in pattern.findall(data) if re.search("explain|readonly|selection|copilot", s, re.I)})
        if flags:
            hits.append({"file": str(path.relative_to(word.parent)), "identifiers": flags[:100]})
        if len(hits) >= 100:
            break
    return hits


def inspect_windows(process_id: int) -> list[dict]:
    from pywinauto import Desktop
    found = []
    for window in Desktop(backend="uia").windows(process=process_id):
        try:
            names = []
            for item in window.descendants()[:250]:
                info = item.element_info
                if info.name:
                    names.append({"name": info.name[:500], "type": info.control_type, "id": info.automation_id, "visible": item.is_visible(), "enabled": item.is_enabled()})
            found.append({"title": window.window_text(), "class": window.class_name(), "elements": names})
        except Exception as exc:
            found.append({"error": type(exc).__name__ + ": " + str(exc)[:250]})
    return found


def connected_dialog(windows: list[dict]) -> bool:
    names = [w.get("title", "") for w in windows]
    names += [e["name"] for w in windows for e in w.get("elements", [])]
    return any(re.search(r"使用\s*Office\s*连接体验|Use Office connected experiences|Use connected experiences", name, re.I) for name in names)


def screenshot(label: str) -> None:
    from PIL import ImageGrab
    try:
        ImageGrab.grab().save(OUT / (label + ".png"))
    except Exception as exc:
        emit(OUT / (label + "-screenshot-error.json"), {"error": type(exc).__name__, "detail": str(exc)})


def finish_privacy_onboarding(main, process_id: int, label: str) -> dict:
    events = []
    for index in range(10):
        time.sleep(2)
        windows = inspect_windows(process_id)
        emit(OUT / (label + "-onboarding-" + str(index) + ".json"), windows)
        visible_names = [e["name"].strip() for w in windows for e in w.get("elements", []) if e.get("visible")]
        has_intro = any(any(title in name for title in ("Microsoft 尊重你的隐私", "Your privacy matters", "助推你的体验", "Powering your experience")) for name in visible_names)
        if has_intro:
            for box in main.descendants(control_type="CheckBox"):
                if box.is_visible() and box.is_enabled() and re.search(r"连接体验|connected experience|可选诊断|optional diagnostic", box.window_text(), re.I):
                    if box.get_toggle_state() == 1:
                        events.append({"disableOptionalCheckbox": box.window_text()})
                        box.toggle()
        buttons = [e for e in main.descendants(control_type="Button") if e.is_visible() and e.is_enabled()]
        chosen = None
        for pattern in (r"不发送.*可选", r"不.*使用.*可选.*连接", r"Don't send optional", r"Do not send optional"):
            chosen = next((b for b in buttons if re.search(pattern, b.window_text(), re.I)), None)
            if chosen is not None:
                break
        if chosen is None and has_intro:
            # These advance explanatory/privacy pages, never a sign-in or activation page.
            for name in ("下一步", "Next", "完成", "Done"):
                chosen = next((b for b in buttons if re.fullmatch(re.escape(name) + r"(?:\([A-Za-z]\))?", b.window_text().strip())), None)
                if chosen is not None:
                    break
        if chosen is not None:
            events.append({"button": chosen.window_text()})
            chosen.click_input()
            continue
        startup_blockers = [n for n in visible_names if re.search(r"登录以设置 Office|Sign in to set up Office|激活向导|Activation Wizard|Microsoft 尊重你的隐私|Your privacy matters|助推你的体验|一起变得更好", n, re.I)]
        body = [item for item in main.descendants(control_type="Edit") if item.element_info.automation_id == "Body"]
        if not body:
            body = [e for e in main.descendants(control_type="Document") if e.is_visible() and e.is_enabled()]
        ready = not startup_blockers and bool(body) and body[0].is_enabled()
        screenshot(label + "-onboarding-final")
        return {"ready": ready, "actions": events, "blockers": startup_blockers[:10]}
    return {"ready": False, "actions": events, "blockers": ["ONBOARDING_STEP_LIMIT"]}


def scan_native_identifiers(word: Path) -> list[dict]:
    hits = []
    pattern = re.compile(r"Microsoft\.Office\.[A-Za-z0-9_.]{1,150}")
    for name in ("WINWORD.EXE", "WWLIB.DLL", "mso.dll", "mso20win32client.dll", "mso30win32client.dll", "mso40uiwin32client.dll"):
        path = word.parent / name
        if not path.is_file() or path.stat().st_size > 160_000_000:
            continue
        raw = path.read_bytes()
        names = set()
        for encoding in ("latin-1", "utf-16-le"):
            data = raw.decode(encoding, errors="ignore")
            names.update(s for s in pattern.findall(data) if re.search("explain|readonly|selection|copilot|contextual", s, re.I))
        hits.append({"file": name, "identifiers": sorted(names)[:500]})
    return hits


def run_scenario(word: Path, readonly: bool, label: str | None = None, com_options: dict | None = None, view_name: str | None = None, view_type: int | None = None, restore_options: bool = True, input_method: str = "keyboard", app_options: dict | None = None, command: str | None = None) -> dict:
    from pywinauto import Desktop, keyboard
    import win32clipboard
    label = label or ("readonly" if readonly else "writable")
    path = OUT / ("synthetic-" + label + ".docx")
    print("START " + label, flush=True)
    make_docx(path)
    if readonly:
        os.chmod(path, stat.S_IREAD)
    proc = subprocess.Popen([str(word), "/n", str(path)])
    result = {"scenario": label, "fileReadonly": readonly, "selectedSyntheticTextConfirmed": False,
              "dialogBeforeSelection": False, "dialogAfterSelection": False,
              "result": "INCONCLUSIVE_NO_MATCHING_AUTHENTICATED_PROFILE"}
    try:
        time.sleep(20)
        startup = inspect_windows(proc.pid)
        emit(OUT / (label + "-startup-ui.json"), startup)
        screenshot(label + "-startup")
        result["dialogBeforeSelection"] = connected_dialog(startup)
        # Escape can dismiss an activation prompt; it never opts in to connectivity.
        for _ in range(2):
            keyboard.send_keys("{ESC}")
            time.sleep(1)
        windows = Desktop(backend="uia").windows(process=proc.pid, class_name="OpusApp")
        if not windows:
            result["result"] = "BLOCKED_NO_WORD_DOCUMENT_WINDOW"
            return result
        main = windows[0]
        main.set_focus()
        onboarding = finish_privacy_onboarding(main, proc.pid, label)
        result["onboarding"] = onboarding
        if not onboarding["ready"]:
            result["result"] = "BLOCKED_STARTUP_UI"
            return result
        main.maximize()
        if view_name:
            tabs = [e for e in main.descendants(control_type="TabItem") if e.is_visible() and e.window_text() == "视图"]
            if not tabs:
                result["result"] = "BLOCKED_NO_VIEW_TAB"
                return result
            tabs[0].click_input()
            time.sleep(1)
            emit(OUT / (label + "-view-tab-ui.json"), inspect_windows(proc.pid))
            buttons = [e for e in main.descendants() if e.is_visible() and e.is_enabled() and re.fullmatch(re.escape(view_name) + r"(?:\([A-Za-z]\))?", e.window_text())]
            if not buttons:
                result["result"] = "BLOCKED_NO_REQUESTED_VIEW"
                return result
            buttons[0].click_input()
            time.sleep(2)
            emit(OUT / (label + "-view-selected-ui.json"), inspect_windows(proc.pid))
            screenshot(label + "-view-selected")
            result["requestedView"] = view_name
        if com_options or view_type is not None or app_options or command or input_method == "mouse":
            import win32com.client
            main.minimize()
            time.sleep(1)
            try:
                app = win32com.client.GetActiveObject("Word.Application")
                result["comHandleMatches"] = int(app.ActiveWindow.Hwnd) == int(main.handle)
                result["comOptions"] = {}
                if not result["comHandleMatches"]:
                    raise RuntimeError("COM application window differs from trial")
                result["documentReadOnlyCOM"] = bool(app.ActiveDocument.ReadOnly)
                result["viewTypeBeforeCOM"] = int(app.ActiveWindow.View.Type)
                result["openAttachmentsBeforeCOM"] = bool(app.OpenAttachmentsInFullScreen)
                if app_options:
                    result["appOptions"] = {}
                    for option, value in app_options.items():
                        before_value = getattr(app, option)
                        setattr(app, option, value)
                        result["appOptions"][option] = {"before": before_value, "after": getattr(app, option)}
                if command:
                    app.CommandBars.ExecuteMso(command)
                    result["executedCommand"] = command
                if view_type is not None:
                    app.ActiveWindow.View.Type = view_type
                    result["viewTypeAfterCOM"] = int(app.ActiveWindow.View.Type)
                for option, value in (com_options or {}).items():
                    before_value = getattr(app.Options, option)
                    setattr(app.Options, option, value)
                    result["comOptions"][option] = {"before": before_value, "after": getattr(app.Options, option)}
            except Exception as exc:
                result["comError"] = type(exc).__name__ + ": " + str(exc)[:250]
            finally:
                main.restore()
                main.maximize()
                main.set_focus()
        body = [item for item in main.descendants(control_type="Edit") if item.element_info.automation_id == "Body"]
        if not body:
            body = [e for e in main.descendants(control_type="Document") if e.is_visible() and e.is_enabled()]
        if not body:
            result["result"] = "BLOCKED_NO_DOCUMENT_BODY"
            return result
        body[0].click_input()
        before = inspect_windows(proc.pid)
        result["dialogBeforeSelection"] |= connected_dialog(before)
        emit(OUT / (label + "-before-selection-ui.json"), before)
        if result["dialogBeforeSelection"]:
            result["result"] = "STARTUP_DIALOG_IS_NOT_SELECTION_REPRODUCTION"
            return result
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
        finally:
            win32clipboard.CloseClipboard()
        result["fileReadonlyObserved"] = bool(path.stat().st_file_attributes & stat.FILE_ATTRIBUTE_READONLY)
        result["inputMethod"] = input_method
        if input_method == "mouse":
            from pywinauto import mouse
            rect = None
            try:
                span = body[0].iface_text.DocumentRange.Clone()
                span.MoveEndpointByRange(1, span, 0)
                span.MoveEndpointByUnit(1, 0, 12)
                bounds = list(span.GetBoundingRectangles())
                if len(bounds) >= 4 and bounds[2] > 1:
                    rect = bounds[:4]
                    result["selectionGeometrySource"] = "UIA text range"
            except Exception as exc:
                result["uiaGeometryError"] = type(exc).__name__
            if rect is None and "app" in locals():
                import pythoncom
                import win32com.client
                coordinates = [win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0) for _ in range(4)]
                app.ActiveWindow.GetPoint(*coordinates, app.ActiveDocument.Range(0, 12))
                rect = [v.value for v in coordinates]
                result["selectionGeometrySource"] = "Word range screen coordinates"
            if rect is None:
                result["result"] = "BLOCKED_NO_TEXT_GEOMETRY"
                return result
            x, y, width, height = [int(v) for v in rect]
            result["selectionRect"] = [x, y, width, height]
            mouse.press(button="left", coords=(x+1, y+max(1,height//2)))
            mouse.move(coords=(x+max(2,width-1), y+max(1,height//2)))
            time.sleep(0.5)
            mouse.release(button="left", coords=(x+max(2,width-1), y+max(1,height//2)))
        else:
            keyboard.send_keys("^{HOME}")
            keyboard.send_keys("+{RIGHT 12}")
        time.sleep(4)
        after = inspect_windows(proc.pid)
        result["dialogAfterSelection"] = connected_dialog(after)
        result["explainEntryObserved"] = any(e.get("name") in ("说明", "Explain") and e.get("visible") for w in after for e in w.get("elements", []))
        emit(OUT / (label + "-after-selection-ui.json"), after)
        screenshot(label + "-after-selection")
        if result["dialogAfterSelection"]:
            result["result"] = "TARGET_DIALOG_OBSERVED_AFTER_SELECTION"
            return result
        keyboard.send_keys("^c")
        time.sleep(1)
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(13):
                    text = win32clipboard.GetClipboardData(13)
                    result["copiedSyntheticSample"] = text[:80]
                    result["selectedSyntheticTextConfirmed"] = bool(text) and SYNTHETIC.startswith(text.strip())
            finally:
                win32clipboard.CloseClipboard()
        except Exception as exc:
            result["clipboardCheckError"] = type(exc).__name__
        return result
    except Exception as exc:
        result["result"] = "BLOCKED_UI_AUTOMATION_ERROR"
        result["error"] = type(exc).__name__ + ": " + str(exc)[:400]
        screenshot(label + "-error")
        return result
    finally:
        if com_options and restore_options and "app" in locals() and result.get("comHandleMatches"):
            try:
                keyboard.send_keys("{ESC}{ESC}")
                for option, values in result.get("comOptions", {}).items():
                    setattr(app.Options, option, values["before"])
                result["comOptionsRestored"] = True
            except Exception as exc:
                result["comRestoreError"] = type(exc).__name__
        if app_options and "app" in locals() and result.get("comHandleMatches"):
            try:
                keyboard.send_keys("{ESC}")
                if restore_options:
                    for option, values in result.get("appOptions", {}).items():
                        setattr(app, option, values["before"])
                app.Quit(0)
                result["wordQuitGracefully"] = True
                time.sleep(1)
            except Exception as exc:
                result["quitError"] = type(exc).__name__
        subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True, check=False)
        time.sleep(2)


FEATURE = "Microsoft.Office.Word.EnableCopilotOptionsInSelectionFloatie"
OVERRIDE = r"Software\Microsoft\Office\16.0\Common\ExperimentConfigs\ExternalFeatureOverrides\word"
PRIVACY = r"Software\Policies\Microsoft\office\16.0\common\privacy"


def privacy_snapshot() -> dict:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, PRIVACY) as key:
        return {n: winreg.QueryValueEx(key, n)[0] for n in (
            "disconnectedstate", "usercontentdisabled", "downloadcontentdisabled", "controllerconnectedservicesenabled")}


def override_snapshot() -> dict:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, OVERRIDE) as key:
            value, kind = winreg.QueryValueEx(key, FEATURE)
            return {"exists": True, "value": value, "kind": kind}
    except FileNotFoundError:
        return {"exists": False}


def set_override(value: str) -> None:
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, OVERRIDE, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, FEATURE, 0, winreg.REG_SZ, value)


def restore_override(original: dict) -> None:
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, OVERRIDE, 0, winreg.KEY_SET_VALUE) as key:
        if original["exists"]:
            winreg.SetValueEx(key, FEATURE, 0, original["kind"], original["value"])
        else:
            try:
                winreg.DeleteValue(key, FEATURE)
            except FileNotFoundError:
                pass


def no_popup_and_copy(result: dict) -> bool:
    return (result.get("fileReadonlyObserved") is True and
            result.get("selectedSyntheticTextConfirmed") is True and
            result.get("dialogBeforeSelection") is False and
            result.get("dialogAfterSelection") is False)


def traced_readonly(word: Path) -> dict:
    tool = os.environ.get("WORDLAB_PROCMON")
    work = Path(os.environ["RUNNER_TEMP"]) / "word-popup-regtrace-v1-7"
    work.mkdir(exist_ok=True)
    pml, csv_path = work / "word.pml", work / "word.csv"
    trace_info = {"available": bool(tool), "filteredRegistryRows": []}
    if not tool:
        return run_scenario(word, True, "control-readonly-no-trace")
    proc = subprocess.Popen([tool, "/AcceptEula", "/Quiet", "/Minimized", "/BackingFile", str(pml), "/Runtime", "65"])
    try:
        time.sleep(2)
        result = run_scenario(word, True, "control-readonly-traced")
    finally:
        subprocess.run([tool, "/Terminate", "/Quiet"], capture_output=True, timeout=30, check=False)
        try:
            proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            proc.kill()
    trace_info["pmlExists"] = pml.is_file()
    trace_info["pmlBytes"] = pml.stat().st_size if pml.is_file() else 0
    try:
        export = subprocess.run([tool, "/AcceptEula", "/Quiet", "/OpenLog", str(pml), "/SaveAs", str(csv_path)], capture_output=True, timeout=180, check=False)
        trace_info["exportExit"] = export.returncode
        if csv_path.is_file():
            with csv_path.open("rb") as stream:
                bom = stream.read(2)
            encoding = "utf-16" if bom in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
            with csv_path.open(encoding=encoding, errors="replace", newline="") as stream:
                for row in csv.DictReader(stream):
                    if row.get("Process Name", "").lower() != "winword.exe" or not row.get("Operation", "").startswith("Reg"):
                        continue
                    path = row.get("Path", "")
                    if re.search(r"ExperimentConfig|Copilot|Explain|PersistentSelection|DABConsumption|Allow.*Float|Office\\16\.0\\Word\\Options|Office\\16\.0\\Common\\Privacy", path, re.I):
                        trace_info["filteredRegistryRows"].append(row)
    except Exception as exc:
        trace_info["error"] = type(exc).__name__ + ": " + str(exc)[:300]
    emit(OUT / "word-registry-trace.json", trace_info)
    return result


def group_snapshot(names: list[str]) -> dict:
    global FEATURE
    result = {}
    for name in names:
        FEATURE = name
        result[name] = override_snapshot()
    return result


def group_set(names: list[str], value: str) -> None:
    global FEATURE
    for name in names:
        FEATURE = name
        set_override(value)


def group_restore(values: dict) -> None:
    global FEATURE
    for name, before in values.items():
        FEATURE = name
        restore_override(before)


def dump_ecs_identifiers() -> dict:
    root = r"Software\Microsoft\Office\16.0\Common\ExperimentConfigs\Ecs\word\ConfigContextData"
    values = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, root) as key:
            for i in range(winreg.QueryInfoKey(key)[1]):
                name, value, kind = winreg.EnumValue(key, i)
                if isinstance(value, bytes):
                    values.append((name, value))
    except FileNotFoundError:
        pass
    names = set()
    p = re.compile(r"[A-Za-z][A-Za-z0-9_.]{3,200}")
    details = []
    for label, raw in values:
        found = set()
        for enc in ("latin-1", "utf-16-le"):
            found.update(s for s in p.findall(raw.decode(enc, errors="ignore")) if re.search("explain|floatie|readonly|copilot|consumption", s, re.I))
        names.update(found)
        if found:
            details.append({"chunk": label, "bytes": len(raw), "identifiers": sorted(found)})
    result = {"binaryValueCount": len(values), "identifiers": sorted(names), "chunks": details}
    emit(OUT / "ecs-identifiers.json", result)
    return result


def mapping_set(values: dict[str, str]) -> None:
    global FEATURE
    for key, value in values.items():
        FEATURE = key
        set_override(value)


def main() -> int:
    if os.name != "nt" or os.environ.get("GITHUB_ACTIONS") != "true" or os.environ.get("RUNNER_ENVIRONMENT") != "github-hosted":
        raise SystemExit("Limited to disposable GitHub-hosted Windows runner.")
    word = find_word()
    if word is None:
        return 2
    stage = sys.argv[2] if len(sys.argv) > 2 else "baseline"
    if stage == "baseline":
        result = {"writable": run_scenario(word, False, "control-writable"), "readonly": run_scenario(word, True, "control-readonly")}
        emit(OUT / "baseline.json", result)
        dump_ecs_identifiers()
        print(json.dumps(result, ensure_ascii=False), flush=True)
        return 0
    baseline = json.loads((OUT / "baseline.json").read_text(encoding="utf-8"))
    positive = baseline["writable"]["selectedSyntheticTextConfirmed"] and baseline["readonly"]["dialogAfterSelection"]
    privacy_before = privacy_snapshot()
    uci = {"Microsoft.Office.UCI.ShowFloatieChangeGate": "false"}
    dab = {"Microsoft.Office.UXPlatform." + name: "false" for name in ("CopilotDABEnabled", "CopilotDABConsumerEnabled", "CopilotDABConsumerProEnabled", "CopilotDABCommercialProEnabled")}
    word_flags = {"Microsoft.Office.Word.CoPilot.Chat": "false", "Microsoft.Office.Word.CopilotStart": "false", "Microsoft.Office.Shared.Copilot.EnableM365CopilotForConsumers": "false"}
    block = {"Microsoft.Office.Word.BlockCopilotChatNoAugloop": "true"}
    admin = {"Microsoft.Office.Word.EnableCopilotDisabledByAdminPolicy": "true"}
    groups = [uci, dab, word_flags, block, admin, {**uci, **dab, **word_flags, **block, **admin}]
    trials = []
    verified = None
    for index, values in enumerate(groups):
        before = group_snapshot(list(values))
        trial = {"values": values, "verified": False}
        try:
            mapping_set(values)
            trial["B"] = run_scenario(word, True, "actual-ECS-" + str(index) + "-B")
            if positive and no_popup_and_copy(trial["B"]):
                trial["B2"] = run_scenario(word, True, "actual-ECS-" + str(index) + "-B2")
        finally:
            group_restore(before)
        if "B2" in trial:
            trial["A2"] = run_scenario(word, True, "actual-ECS-" + str(index) + "-A2-restored")
            trial["verified"] = no_popup_and_copy(trial["B2"]) and trial["A2"]["dialogAfterSelection"] and privacy_before == privacy_snapshot()
        trial["restored"] = before == group_snapshot(list(values))
        trials.append(trial)
        emit(OUT / "actual-ecs-progress.json", trials)
        print(json.dumps(trial, ensure_ascii=False), flush=True)
        if trial["verified"]:
            verified = values
            break
    view_fallback = {"verified": False}
    if not verified:
        view_fallback["B"] = run_scenario(word, True, "reading-mouse-max-B", view_name="阅读视图", input_method="mouse")
        view_fallback["B2"] = run_scenario(word, True, "reading-mouse-max-B2", view_name="阅读视图", input_method="mouse")
        view_fallback["A2"] = run_scenario(word, True, "print-mouse-max-A2", view_name="页面视图", input_method="mouse")
        view_fallback["verified"] = positive and no_popup_and_copy(view_fallback["B"]) and no_popup_and_copy(view_fallback["B2"]) and view_fallback["A2"]["dialogAfterSelection"]
        emit(OUT / "view-fallback.json", view_fallback)
    result = {"version": "1.9", "viewFallback": view_fallback, "baseline": baseline, "trials": trials, "verifiedValues": verified,
              "fixedInTestEnvironment": bool(verified), "privacyUnchangedAndDisabled": privacy_before == privacy_snapshot() and all(v == 2 for v in privacy_snapshot().values())}
    emit(OUT / "attempt.json", result)
    print(json.dumps(result, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
