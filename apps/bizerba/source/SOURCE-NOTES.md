# Bizerba trainer — build sources

These are the V6 sources that built `../index.html` (GLM-I scale terminal, Bizerba Brain 2, Data Management).

```
python3 build.py        # writes labeller.html and "Labeller Line Trainer V6.html"
```

`build.py` expects `manual.json`, `figs.json` and `prices.json` beside it.

- `manual.json` — 726 sections of the GLM-I instruction manual, used to build the menu tree.
- `figs.json` — 43 manual figures as data URIs.
- **`prices.json` is not in this repository.** It is the site price list, internal customer data, so it stays on the work laptop. `plu-demo.json` / `plu-demo.csv` are fictional stand-ins, and `tools/pricelist_to_plu.py` turns a real price list into that format. The parser drops the staff email addresses in the sheet; do not add them back.

`HANDOFF_V6.md` is the full build history and the list of what came from photos versus what is assumed. `HANDOFF_V5.md` is the version before it.

After a rebuild, copy the built file in and put the home button back:

```
python3 ../../../tools/sync.py --local <folder holding the built index.html> bizerba
```
