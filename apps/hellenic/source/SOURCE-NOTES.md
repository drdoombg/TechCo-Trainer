# Hellenic trainer — build sources

These are the sources that built `../index.html` (Scale terminal, FOPS 8, FOPSX Web Editors, Comms VNC, Label Designer).

```
node tools/build.js     # concatenates src/p1-head.html + src/p2-hx.js … into index.html
node tools/test-run.js  # click-through check
```

`screenshots/` holds the terminal photos the screens were matched to, and `tools/photo-checklist.html` lists the photos still wanted. `Handoff/HANDOFF.md` is the full build history.

After a rebuild, copy the built file in and put the home button back:

```
python3 ../../../tools/sync.py --local <folder holding the built index.html> hellenic
```
