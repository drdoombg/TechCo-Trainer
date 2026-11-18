# TechCo Trainer

One home screen for the trainers used at 2 Sisters Food Group, Scunthorpe. Open the page, pick a trainer, start. Everything runs in the browser with no install and no connection to any real factory system.

| Plate | What it practises | Programs inside |
|---|---|---|
| **Bizerba** | GLM-I 40 / 70, C-Wrap applicator (FP2 Lines 3–6, FP1 Line 3) | GLM-I scale terminal · Bizerba Brain 2 · Data Management |
| **Hellenic** | Hellenic Systems OCM, Requiem v9 | Scale terminal · FOPS 8 · FOPSX Web Editors · Comms (VNC) · Label Designer |
| **Ishida** | WPL weigh-price labeller, ICE Zodiac Hawk printer (Line 4) | Office PC (CommsBridge, CLARiSOFT) · Line 4 WPL · ICE Hawk printer |
| **Excel** | The formula dialect of the work spreadsheets | Formula sandbox with the shelf life, nutritional and micro schedule toolkits |

Each trainer is the same app it has always been. The combined site only adds the home screen and a **Main menu** button in the corner of each trainer (also **Alt+H**).

This repository is now the only home for all of them: the built trainers, the sources they are built from, and the handoff notes.

## Layout

```
index.html                   home screen (built from src/home.inner.html)
src/home.inner.html          the home screen itself: title, styles, plates, script
tools/build_home.py          src/home.inner.html -> index.html
tools/sync.py                copy a trainer in and inject the home button
apps.json                    where each trainer comes from
dist/artifact-home.html      the home screen without the document shell

apps/bizerba/index.html      the trainer that runs
apps/bizerba/source/         what builds it, plus HANDOFF_V6.md
apps/hellenic/index.html
apps/hellenic/source/        build sources, terminal photos, HANDOFF.md
apps/ishida/index.html
apps/ishida/source/          the file as delivered
apps/excel/                  multi-file app: index.html, css/, js/, data/, docs/
apps/excel/source/           the formula catalogue it was built from
```

Saved data (settings, passwords, drill progress) lives in the browser, one set of keys per trainer, so they never overwrite each other. Anything already saved carries over as long as the site keeps the same `github.io` host.

## Running it

- **Hosted:** GitHub Pages serves the repository root. The home screen is the site root; each trainer is a folder under `apps/`.
- **Offline:** copy the whole folder to the laptop and open `index.html`. Every link is relative, so it works straight from the file system.

## Rebuilding a trainer

Each trainer builds the same way it always did — see `SOURCE-NOTES.md` in its `source/` folder. Afterwards, copy the built file in and put the Main menu button back:

```bash
python3 tools/sync.py --local <folder holding the built index.html> bizerba
python3 tools/sync.py ishida      # Ishida builds from apps/ishida/source/index.html
python3 tools/sync.py --local "D:/Work/3. Technical Co-ordinator/_Apps/Excel Trainer/app" excel
python3 tools/sync.py --reinject  # only put the Main menu button back, in every trainer
```

For a multi-file trainer, list its folders under `"dirs"` in `apps.json` and `--local` copies those too.

`sync.py` writes `apps/<id>/index.html` and injects the Main menu button between the `TECHCO:HOME` markers at the end of `<body>`. Running it again replaces that block instead of adding another, so the file stays exactly as its build produced it.

## Adding a trainer

1. Put the trainer's `index.html` in `apps/<id>/`, its build sources in `apps/<id>/source/`, and any assets it loads beside it.
2. Add an entry to `apps.json` so `sync.py` can re-inject the Main menu button later. A multi-file app lists its folders under `"dirs"`.
3. Copy a plate in `src/home.inner.html`, change the name, model line, program list, bottom strip and `data-n` / `data-id`, then run `python3 tools/build_home.py`.
4. Update the trainer and program counts in the header.

## Editing the home screen

Edit `src/home.inner.html` and run `python3 tools/build_home.py`. Do not edit `index.html` by hand — it is generated.
