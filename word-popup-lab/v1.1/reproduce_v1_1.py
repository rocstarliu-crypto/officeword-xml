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

SYNTHETIC = "WORDLAB_SYNTHETIC_TEXT_1234567890 这是只读选字测试文档，不包含个人内容。"
OUT = Path(sys.argv[1] if len(sys.argv) > 1 else "lab-results").resolve()
OUT.mkdir(parents=True, exist_ok=True)


def emit(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def make_docx(path: Path) -> None:
    parts = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/document.xml": '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' + SYNTHETIC + '</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>',
    }
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
                    names.append({"name": info.name[:500], "type": info.control_type, "id": info.automation_id})
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


def run_scenario(word: Path, readonly: bool) -> dict:
    from pywinauto import Desktop, keyboard
    import win32clipboard
    label = "readonly" if readonly else "writable"
    path = OUT / ("synthetic-" + label + ".docx")
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
        keyboard.send_keys("^{HOME}")
        keyboard.send_keys("+{RIGHT 12}")
        time.sleep(4)
        after = inspect_windows(proc.pid)
        result["dialogAfterSelection"] = connected_dialog(after)
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
        subprocess.run(["taskkill", "/PID", str(proc.pid), "/T", "/F"], capture_output=True, check=False)
        time.sleep(2)


def main() -> int:
    if os.name != "nt" or os.environ.get("GITHUB_ACTIONS") != "true" or os.environ.get("RUNNER_ENVIRONMENT") != "github-hosted":
        raise SystemExit("This attempt is limited to a disposable GitHub-hosted Windows runner.")
    word = find_word()
    if word is None:
        emit(OUT / "attempt.json", {"status": "BLOCKED_WORD_NOT_INSTALLED", "reproduced": False, "fixed": False})
        return 2
    emit(OUT / "observed-feature-identifiers.json", scan_flag_identifiers(word))
    scenarios = [run_scenario(word, False), run_scenario(word, True)]
    baseline = scenarios[0]["selectedSyntheticTextConfirmed"] and not scenarios[0]["dialogAfterSelection"]
    observed = baseline and scenarios[1]["result"] == "TARGET_DIALOG_OBSERVED_AFTER_SELECTION"
    result = {"version": "1.1", "status": "TARGET_OBSERVED_DIFFERENT_PROFILE" if observed else "INCONCLUSIVE",
              "reproducedInUnsignedProfile": observed, "reproducedInMatchingUserEnvironment": False,
              "fixed": False, "scenarios": scenarios,
              "limitations": ["No activated matching Microsoft 365 Family account", "No authenticated OneDrive sync", "Windows Server differs from the user's Windows client"]}
    emit(OUT / "attempt.json", result)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as stream:
            stream.write("\n## Real Word UI attempt V1.1\n\n```json\n" + json.dumps(result, ensure_ascii=False, indent=2) + "\n```\n")
    return 0 if observed else 2


if __name__ == "__main__":
    raise SystemExit(main())
