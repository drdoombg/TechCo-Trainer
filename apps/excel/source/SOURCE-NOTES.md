# Excel Trainer — source

This trainer is a multi-file app with no build step, so `../index.html`, `../css`, `../js` and `../data` are the source. Edit them in place.

`01-Formula-Catalog.md` is the catalogue of formulas pulled out of the real work spreadsheets that the toolkits and levels were built from. The `scan/` scripts that produced it stay on the work laptop with the spreadsheets.

`../tests.html` runs the engine tests in the browser. `../docs/` holds the app's own README and open questions.

To refresh this folder from the working copy on the laptop:

```
python3 tools/sync.py --local "D:/Work/3. Technical Co-ordinator/_Apps/Excel Trainer/app" excel
```

That copies `index.html`, `css/`, `js/`, `data/` and `docs/` in, and puts the Main menu button back into `index.html`.
