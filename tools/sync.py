#!/usr/bin/env python3
"""Pull each trainer into apps/<id>/ and add the TechCo home button.

    python3 tools/sync.py                 # all apps listed in apps.json
    python3 tools/sync.py bizerba         # just one
    python3 tools/sync.py --local ../Bizerba-Trainer bizerba
    python3 tools/sync.py --reinject      # only put the Main menu button back

Each app keeps its own index.html exactly as its own repo built it; the only
change is a small home button injected between the TECHCO markers at the end of
<body>, so re-running this never stacks copies of it.
"""
import json, pathlib, shutil, sys, urllib.request

ROOT = pathlib.Path(__file__).resolve().parent.parent
START, END = "<!--TECHCO:HOME-START-->", "<!--TECHCO:HOME-END-->"

HOME_BUTTON = START + """
<style>
#techco-home{position:fixed;z-index:2147483000;right:calc(12px + env(safe-area-inset-right,0px));bottom:calc(12px + env(safe-area-inset-bottom,0px));display:flex;align-items:center;gap:7px;height:34px;padding:0 13px 0 11px;border-radius:999px;background:#111922;color:#fff;border:1px solid rgba(255,255,255,.2);text-decoration:none;font:600 12.5px/1 "IBM Plex Sans",system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:.01em;opacity:.8;transition:opacity .15s,transform .15s;box-shadow:0 2px 10px rgba(0,0,0,.3)}
#techco-home:hover,#techco-home:focus-visible{opacity:1;transform:translateY(-1px);outline:none}
#techco-home:focus-visible{box-shadow:0 0 0 2px #0b5fa5}
#techco-home svg{width:15px;height:15px;flex:none}
@media print{#techco-home{display:none}}
</style>
<a id="techco-home" href="../../index.html" title="Back to the TechCo Trainer menu (Alt+H)" aria-label="Back to the TechCo Trainer main menu"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 11l9-8 9 8"/><path d="M5 10v10h14V10"/></svg>Main menu</a>
<script>document.addEventListener("keydown",function(e){if(e.altKey&&!e.ctrlKey&&!e.metaKey&&(e.key==="h"||e.key==="H")){e.preventDefault();e.stopPropagation();location.href="../../index.html"}},true);</script>
""" + END


def inject(html: str) -> str:
    if START in html and END in html:
        html = html[: html.index(START)] + html[html.index(END) + len(END) :]
    tag = "</body>"
    i = html.lower().rfind(tag)
    if i == -1:
        return html + "\n" + HOME_BUTTON + "\n"
    return html[:i] + HOME_BUTTON + "\n" + html[i:]


def fetch(src: str) -> str:
    if src.startswith("http"):
        with urllib.request.urlopen(src, timeout=60) as r:
            return r.read().decode("utf-8")
    return pathlib.Path(src).read_text(encoding="utf-8")


def main() -> int:
    args = [a for a in sys.argv[1:]]
    reinject = "--reinject" in args
    if reinject:
        args.remove("--reinject")
    local = None
    if "--local" in args:
        k = args.index("--local")
        local = args[k + 1]
        del args[k : k + 2]
    apps = json.loads((ROOT / "apps.json").read_text(encoding="utf-8"))
    wanted = args or [a["id"] for a in apps]
    for app in apps:
        if app["id"] not in wanted:
            continue
        out = ROOT / "apps" / app["id"]
        out.mkdir(parents=True, exist_ok=True)
        if reinject:
            f = out / "index.html"
            if f.exists():
                f.write_text(inject(f.read_text(encoding="utf-8")), encoding="utf-8")
                print("%-9s Main menu button in place" % app["id"])
            continue
        src = app.get("source")
        if local:
            src = str(pathlib.Path(local) / "index.html")
        if src and not src.startswith("http") and not pathlib.Path(src).is_absolute():
            src = str(ROOT / src)
        if not src:
            print("%-9s no source in apps.json — left as it is" % app["id"])
            continue
        html = fetch(src)
        (out / "index.html").write_text(inject(html), encoding="utf-8")
        print("%-9s %7.0f KB  <- %s" % (app["id"], len(html) / 1024, src))
        for d in app.get("dirs", []):
            srcdir = pathlib.Path(src).parent / d
            if srcdir.is_dir():
                shutil.rmtree(out / d, ignore_errors=True)
                shutil.copytree(srcdir, out / d)
                print("          %s/ copied" % d)
        for extra in app.get("files", []):
            base = src.rsplit("/", 1)[0]
            try:
                (out / extra).parent.mkdir(parents=True, exist_ok=True)
                if src.startswith("http"):
                    (out / extra).write_text(fetch(base + "/" + extra), encoding="utf-8")
                else:
                    shutil.copy(pathlib.Path(src).parent / extra, out / extra)
            except Exception as e:  # a missing extra file is not fatal
                print("          skipped %s (%s)" % (extra, e))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
