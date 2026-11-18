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

## V6 (27/10/2026): Bizerba Brain 2 and Data Management rebuilt from the 27/10 photos and notes
Sources: Photos.pdf (31 photos: BRAIN2 Label Designer on layout 077-77-test, dataMaintenance.BRAIN 3.92 screens, the printed 7777 test labels, Victor's annotated Co-op label), Label_Format_Creation.docx (Birat's notes), the Export-to photo, Birat's answers.

- **App tabs:** GLM-I Scale terminal · **Bizerba Brain 2** · **Data Management** (separate program, own window, not BRAIN2 Master Data).
- **Who does what (notes):** BRAIN2 makes the label format (fields and field types). Data Management assigns values to text fields, matches PLU → label parameter set (the "slot", e.g. 33, whose Label type = Customer Label NN = layout), keeps devices/IPs, sends to the terminals. Layouts can be sent from BRAIN2 (Export to…) and label parameters from DM. Make the slot in DM first, then fill it from BRAIN2.
- **Field rule (Birat):** fixed text and predefined fields (weight, price, unit price, date, clock, device number, code, logo) are set in BRAIN2. **Text field 01–10** content comes from the article's **General text no. 1–10** in DM. Date field 1 = Julian (day of year), Date field 2 = use by. Designer shows 23.09.76 / 14.12.01 only as preview.
- **BRAIN2 shell:** bottom bar … / Product data / Process data / Result data / Reporting / Jobs / System data; footer Admin (Two Sisters Administrator) · STBIZERBA.2SFG.NET:2015.
- **Label Designer:** File Edit View Elements Properties Window, three toolbar rows (icons drawn from the photo, blurred ones marked in tooltips), field-selection combo, Project view (003.3 … 077.77-test, 082.82 booths ac, 083, 085.85 OCADO new, 087.87- M&S BDF PZ C WR, 091.91 M&S EVENT, 092.92 M&S Truffle, 173.173 CO-OP …) + Label window tab, Property window (Categorized/Alphabetical, mm/Dot; General · Position · Appearance · Settings; Resource props Label width 520 Dot / length 2120 Dot / Distance 10 Dot / Extended layout / Resolution 8 Dots/mm), document tabs NNN-name, print legend, Text field editor, status bar (1009 · Gx labeling mode · [Dots] x,y · CAP NUM SCRL). Right-click › Change field type. Field types incl. **Multi** (photo, Co-op layout).
- **Layouts:** 077 77-test as photographed; 173 Co-op from Victor's label (code line = Device number, Clock 1, Text field 6, Date 1 Julian); 205 rebuilt the same way; 082–092 names only (contents assumed). Layout 206 is not pre-made: the drill makes it.
- **Data Management:** tree Product Maintenance (Labeler: Articles, Article Exceptions, Tare Values, General Text, Date Text, Code Structures; Inspection System: Articles) · Device Maintenance (Auto Labeler Parameters, Label Parameters, Weight Classes) · System Configuration (Locations, Lines, Devices) · Upload / Download (Upload Data, Download Data). Articles with all 10 tabs as photographed; 0 = keep previously loaded data, -1 = remove / deactivate / No text; "Please select a text" lookup; Send format Yes/No asked on article save (Yes = small text format, No = layout font; wording assumed). Download Data exactly as photo (tables GGDAT, GGATX, GGDTX, STATTXT, GGCAB, GGSCAN, AMPAR, GGETIKET, GGMINMAX, TENDREG; Only send changed records; Set Selected Devices As Up-To-Date).
- **PLU 7777 TEST TEST** seeded as photographed in DM (codes/logo 0 = keep) and as it stood on the terminal (Code 1 price code, substring 123456, Logo GB 2037, £15.99 fixed). Printed label matches the photos: 11.26 unit price, 1.708, £15.99, 1.420kg, 11, GB 2037, 06NOV, TEST 1, clock, barcode 1 234562 015991, text fields 2–9; text field 10 (=$%$ GL0A check) is a command and does not print.
- **Code structures** rebuilt from the test labels: 1 price code, 3 price code 4-digit, 5 EAN, 9 weight code (check digits verified). 2/4/6 assumed.
- **General texts** from the DM photo (1 L1 … 24073, 666666666 = 6, 777777777 kill-line text, 100000000 formula; 11100–11159 "1hr N mins", 12100–12159 "2hrs N mins" assumed). Text 1000 had a person's name: placeholder.
- **FP2 lines 05 and 06** added to the terminal (L05 master only, L06 master + slave) because DM lists FP2 Line03–Line06.
- **Weight class 7929** (PLU 6706) with cooking-time texts per band (bands assumed).
- **Assumed:** DM unit price 0.00 keeps the terminal's price; device numbers (L03 = 11, L04 = 12, L05 = 15, L06 = 16); IPs; Tare/Date Text/Code Structures/Article Exceptions/Inspection/Auto Labeler/Weight Class/Locations/Lines/Devices/Upload screens (not photographed).
- **Drills (Data Management group):** weight code on 7777, 0 keeps / -1 removes (logo), change a text field value, the M&S fixed cooking-time job (205 → 206, slot 206, Weight Class -1, General text 1 = 2hrs 5mins), Up-To-Date trap, make the slot before the format.
- **Labels & log tab:** "Printed label at full size" for long labels like 77-test.
- **Open:** why Date field 1 printed blank on the 27/10 test labels (trainer prints the Julian day); exact Send format question wording; the weight class 7929 screen; DM screens not photographed; device numbers.

Files: template.html, v6.css, app_main.js, b2a.js (layouts/fields/rendering), b2b.js (BRAIN2 shell/state), b2c.js (Label Designer), b2d.js (other BRAIN2 modules), dm.js (Data Management), b2e.js (events), app_boot.js, build.py → labeller.html + "Labeller Line Trainer V6.html".
