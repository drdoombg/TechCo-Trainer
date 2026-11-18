#!/usr/bin/env python3
"""Build index.html (the home screen) from src/home.inner.html.

    python3 tools/build_home.py

src/home.inner.html holds the page itself: <title>, <style>, the plates and the
small script. This wraps it in the document shell GitHub Pages needs and writes
index.html. dist/artifact-home.html is the same page without the shell, for
publishing as a Claude artifact.
"""
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
inner = (ROOT / "src" / "home.inner.html").read_text(encoding="utf-8").strip()

SHELL = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#eaeef2">
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="assets/icon.svg" type="image/svg+xml">
<style>
:root{color-scheme:light;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
body{margin:0;font:14px system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
img{max-width:100%%}
[hidden]{display:none!important}
</style>
%s
</head>
<body>
%s
</body>
</html>
"""

head_tags, body_html = [], []
for line in inner.splitlines():
    s = line.strip()
    if s.startswith(("<title", "<meta", "<link")) and not body_html:
        head_tags.append(line)
    else:
        body_html.append(line)

# the <style> block belongs in the head too; it is the first block of the body list
text = "\n".join(body_html)
if text.lstrip().startswith("<style"):
    end = text.index("</style>") + len("</style>")
    head_tags.append(text[: end].strip())
    text = text[end:]

(ROOT / "index.html").write_text(
    SHELL % ("\n".join(head_tags), text.strip()), encoding="utf-8"
)
(ROOT / "dist").mkdir(exist_ok=True)
(ROOT / "dist" / "artifact-home.html").write_text(inner + "\n", encoding="utf-8")
print("index.html %.1f KB · dist/artifact-home.html %.1f KB"
      % ((ROOT / "index.html").stat().st_size / 1024, len(inner) / 1024))
