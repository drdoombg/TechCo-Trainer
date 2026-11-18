# Changelog

## 2026-11-18 — Excel Trainer added
- Fourth plate: the Excel Trainer (formula sandbox, shelf life / nutritional / micro schedule toolkits), copied in from the work laptop with its engine, levels and formula catalogue.
- The corner button in every trainer now reads **Main menu** instead of being an icon on its own.
- `tools/sync.py --reinject` puts that button back in every trainer, and `"dirs"` in `apps.json` lets a multi-file trainer be refreshed folder and all.

## 2026-11-18 — TechCo Trainer
This repository was the Bizerba trainer. It now holds all of the machine trainers behind one home screen, and is renamed to match.

- Home screen with a plate per machine: Bizerba, Hellenic, Ishida, plus a free slot for the next one.
- Each trainer moved under `apps/<id>/`, unchanged apart from a home button in the corner (Alt+H).
- Each trainer's own build sources and handoff notes moved to `apps/<id>/source/`, so the separate Bizerba and Hellenic repositories are no longer needed.
- `tools/build_home.py` builds the home screen; `tools/sync.py` re-injects the home button after a trainer is rebuilt.

Trainer histories up to this point are in `apps/bizerba/source/CHANGELOG.md` and `apps/hellenic/source/CHANGELOG.md`.
