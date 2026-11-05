"""Collect only diagnostic feature and setting identifiers, no source excerpts."""
import json, os, re, sys
from pathlib import Path
sys.stdout.reconfigure(encoding="utf-8")
if os.name != "nt" or os.environ.get("GITHUB_ACTIONS") != "true":
    raise SystemExit("Disposable Windows runner only")
out=Path(sys.argv[1]);out.mkdir(exist_ok=True,parents=True)
roots=[]
for base in (os.environ.get("ProgramFiles"),os.environ.get("ProgramFiles(x86)")):
    if base:
        for rel in ("Microsoft Office/root/Office16", "Common Files/Microsoft Shared/OFFICE16"):
            p=Path(base)/rel
            if p.is_dir(): roots.append(p)
pattern=re.compile(r"[A-Za-z][A-Za-z0-9_.]{3,180}")
interest=re.compile(r"explain|floatie|readonly|readmode|readingmode|viewmode|copilot.*(?:enabl|disabl)|(?:enabl|disabl).*copilot",re.I)
results=[]
for root in roots:
    for path in root.rglob("*"):
        if not path.is_file() or path.suffix.lower() not in {".exe",".dll",".js",".json"} or path.stat().st_size>200_000_000:
            continue
        try:
            raw=path.read_bytes()
            hits=set()
            for enc in ("latin-1","utf-16-le"):
                hits.update(t for t in pattern.findall(raw.decode(enc,errors="ignore")) if interest.search(t))
            if hits:
                results.append({"file": str(path),"identifiers":sorted(hits)[:1500]})
        except OSError:
            pass
(out/"wide-identifiers.json").write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding="utf-8")
print("Scanned roots:",*[str(x) for x in roots],"files with identifiers",len(results))
