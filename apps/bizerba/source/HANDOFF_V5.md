# Labeller Line Trainer V5 — handoff (26/10/2026)

Builds on repo v4.4 (github.com/drdoombg/Bizerba-Trainer, src/app.js) + the V3 price list, plus the 31 FP2-L04 photos (Bizerba_Scale_More_images.pdf) and one BRAIN2 screen photo. No passwords in the code.

## What changed
- **App tabs:** GLM-I Scale terminal · BRAIN2 Label Designer · BRAIN2 Data management (same BRAIN2 window, different start module).
- **FP2 Line 3 and 4 merged** into one tab with an L3/L4 switch (devices 2: FP2-L0x-Master, 3: FP2-L0x-Slave). Old drill progress for both lines carried over.
- **Passwords (V5.1):** label checker password → PLU change on level 9 (# macros plu change, m&s plu only), training default 1111. Higher password → levels 2, 3, 4, 5, training default 2222. Levels 0, 1, 9 open; 183 from standby.
- **Photographed menu** (Settings can switch back to the manual layout): Article data / Production data / Display / Configuration / Free / Free / Device (Stop|Start) / Service. Photographed branches built key for key; the rest filled from the manual (marked assumed).
- **New terminal screens from photos:** Load / Load record dialog, Edit current article data + Article data maintenance (tabs General, Texts, Barcode, Label parameters, Aut.lab. parameters, Package selection; Master + Slave columns), Templates + "Template with this number does not exist!", full-screen text editor, Load weight class with "Loading data", breadcrumb tab, Free keys.
- **PLU 6419** set to photo values (tare 62 g, min 2458, max 2874, shelf-life 9, label 205, weight class 7292, code substring 00015035). Text field 10 formula enforced: shelf-life outside 7–10 days → "Invalid Use By Date" lock (GLDA read as shelf-life: assumption).
- **Label layouts** are now data (layout 205 M&S NFO 2 drawn as on the preview: COOKS IN, USE BY, GB 2037, Not for EU, SC1606). The terminal prints whatever layout BRAIN2 last exported to that line.
- **BRAIN2:** Product data (Code Structure, Label Designer, Layout Management, Master Data, Master Data Exceptions, Text), Process data (Devices with download/upload/compare, Weight Classes, Automation/Label Parameters, Data Transfer, Device Access), Result data (Individual Values, Totals, Give-away, Rejects — fed by packs run on the terminal), Reporting, … (Audit Trail, Users & Groups, Backup, Licences, About). FP1-L03 accepts layouts but refuses master data.
- **8 BRAIN2 drills** added to the Drills tab.

## Files
template.html, app_main.js (terminal), b2a–b2e.js (BRAIN2), app_boot.js, build.py → labeller.html (artifact) and the standalone V5 html.

## Still open
Real layouts for non-M&S customers, the label-and-hand key, contents of Production data/Display/Configuration/Device/Service at each level (photos wanted), BRAIN2 screens other than Label Designer (photos wanted), what Text field 6 ("6") and the code line mean.

## V5.1 additions (from github.com/taubedonner/gxnet, MIT, GxNet protocol notes)
- Text field 10 token is GL0A = GGL_HALTBAR1 (shelf-life 1); photo blur made it look like GLDA (= clock 3).
- GX status messages 107 layout missing, 120 code build rule missing, 129 code substring missing (+ two fault drills).
- Package error flags PSL_PCK_ERR_FLAGS (PL13) on every pack, shown in BRAIN2 Individual Values and Rejects.
- BRAIN2 Devices: GX status (PL03), labeller state (SW98), last package (SV70); memory card polled every 5 s.
- Code Structure uses real build-rule identifiers (CODE_KONST, CODE_TEILSTR1-7, CODE_ETX, CODE_GS, CODE_FNC1).
- Text limits: article text 64, code substring 30.
- Configuration › Communication configuration (Del. remote softkeys) from a newer GLM manual.
- About BRAIN2 notes CVE-2026-16246 (BRAIN2 < 3.09) and CVE-2025-12507 (BCS).
