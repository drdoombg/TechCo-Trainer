# Excel Trainer

Offline formula trainer for the dialect of Excel described in
`../01-Formula-Catalog.md`. Section references below (§, P#, D#) are that document.

## The two ways through it

**Learn by formula** (the `Learn by formula` tab) — one function at a time, in the
order the four toolkits actually use them: IF, VLOOKUP, IFERROR, SUMIF, COUNTIF,
SUM, MIN/MAX, MATCH, COUNT, COUNTIFS, INDEX, AND/OR, ISERROR, TEXTAFTER, VALUE,
XLOOKUP, `&`, RIGHT, AVERAGE, ROUND, dates. 22 lessons, 38 graded tasks, each on a
real toolkit worksheet. Defined in `js/trainer/formulas.js`.

**The course** (the `Progress` tab, or `Follow the course`) — levels 0-18, one
pipeline end to end. Unchanged, in `js/trainer/levels.js`.

## The four toolkits

These are the workbooks the job actually runs on, and they come first in the
scenario list:

| | Sheets it rebuilds | What it teaches |
|---|---|---|
| **T1 Shelf life toolkit** | MICRO PASTE → SET MOD → SAMPLE ID NO → COP → MICRO MOD → Test Profile 1 | XLOOKUP looking left, the paired `.` result columns, turning `<10` into a number |
| **S5 Neck & Caeca labels toolkit** | LBI Paste → Despatch Hidden → LBI Hidden → Batch Data Information → Label Info → Error Coding | four sources, a hidden workings layer, a self-audit sheet |
| **T3 Nutritional toolkit** | PASTE → LABELS INFOS → background → 2sfg UKP Nutritional Toolkit | one flat key→value paste, three reps averaged, the tolerance test |
| **T4 Micro schedule** | Swab Results Database → Risk Assessment → QTR 1 → Point List → Home | the Y/N risk ladder, the 13-week grid, SUMIF and COUNTIF |

Fixtures for T1/T3/T4 are in `data/toolkits.js`; S5 is in `data/fixtures.js`.
Still synthetic — structure copied, values invented.

## Run it

Double-click **`app/index.html`**. No server, no install, no internet.
All scripts are classic `<script src>` (not ES modules, which don't load from
`file://`) and all data is JS assigning to `XT.data`, so nothing is fetched.

Open **`app/tests.html`** for the conformance suite: green/red/amber table.
Currently **300 cases, 296 pass, 0 fail, 4 open** (the 4 are listed in
`OPEN_QUESTIONS.md` — things that need checking against real Excel).

## Keys

| | |
|---|---|
| F2 | edit the cell |
| **F4** | cycle `$` on the reference under the cursor — `A1 → $A$1 → A$1 → $A1` |
| Ctrl+D / Ctrl+R | fill down / fill right |
| Ctrl+Shift+Enter | enter as a legacy array formula |
| Ctrl+\` | show all formulas |
| Ctrl+Z / Ctrl+Y | undo / redo |
| Ctrl+Enter | check the exercise |
| drag / double-click the fill handle | fill, with reference translation |
| Ctrl+V on the grid | paste tab-separated text |

## The panels (right-hand side)

Five tabs, not ten. Two of them group the rest behind pill sub-tabs.

- **Task** — the current task, hints one at a time, the grader's feedback, and
  next/previous within whichever track you are on.
- **Learn by formula** — the 22 lessons. Click one to open it: what the function
  does, what it does *in your files*, the trap, and its graded tasks.
- **Explain this cell** — *Evaluate* (Excel's Evaluate Formula, one substitution at
  a time), *Layers* (the formula as guard → route → join), *Inspector* (what the
  cell really holds, and for an error *which part* produced it — including
  **errors an `IFERROR` has swallowed**, which Excel cannot show), *Trace*
  (precedents and dependents, cross-sheet included), *Shape* (the formula in
  catalog notation with its matching P#).
- **Progress** — formula-track tasks passed, then course coverage weighted by the
  catalog's measured formula counts.
- **Reference** — *Functions*, *Patterns*, *Toolkits*.

The toolbar is the same idea: **Check** (primary), Hint, Show me, Reset, the two
fills, Formulas, Undo — and **More** for Paste export and the trace arrows.

## What's in it

**Engine** (`js/engine/`) — ~70 functions, the §12.3 list plus the estate's extras.
Correct `EMPTY` vs `""` vs `0` semantics (§7.4), Excel's cross-type comparison
(text > any number), criteria strings built with `&`, 1900 date serials,
whole-column clipping, legacy CSE arrays, structured `Table[[#This Row],[Col]]`,
external `[1]Sheet!A1` (parses, reports "link not available"), and `#REF!` as a
live value. Functions from §8 (FILTER, SORT, LET, LAMBDA, IFS…) are *registered as
deliberately unsupported* so a pasted formula explains itself instead of `#NAME?`.

**Trainer** (`js/trainer/`) — levels 0–18 per §11, 34 exercises, plus 22
formula lessons and 38 formula tasks in `formulas.js`. Grader checks
**values and shape**: hard-coding fails, and so does a formula that works in the
first cell but breaks when filled. Level 17 is graded structurally (did you keep
the five layers separate?); level 18 is five seeded catalog defects to find and fix.
Levels 14–15 (conditional formatting, dropdowns) are scoped to v2.

**Fixtures** (`data/fixtures.js`) — all five scenarios, generated and synthetic.
No real workbook data anywhere. What is copied is structure: sheet names,
column positions (MIMS SAVE FILE keys on M with a second key in V; Volumes Paste
has three rival key columns and dates across row 13), and the dirt — trailing
spaces, `CHAR(160)`, numbers stored as text, zero padding, blank rows, padding zeros.

## How to add things

- **A function** — `XT.defineFunction({name, min, max, impl, sig, doc, estate, pats})`
  in a file under `js/engine/functions/`, then add it to the script list in both
  `index.html` and `tests.html`. Use `arrayArgs: [1,2]` for argument positions that
  must evaluate in array context (that's what makes `XLOOKUP(k, A:A&B:B, C:C)` work).
  Use `lazy: true` if it must not evaluate all its arguments (IF, IFERROR, IFNA).
- **A formula lesson or task** — add to `XT.data.formulaTrack` in
  `js/trainer/formulas.js`. A lesson needs `fn`, `title`, `cells`, `scenario`,
  `what`, `estate`, `trap` and `drills`; each drill is an ordinary exercise
  object wrapped in `drill({…})`, so the same grader applies.
- **An exercise** — push onto a level's `exercises` in `js/trainer/levels.js`.
  Give it `setup` (stands in for earlier layers, locked on the grid), `targets`,
  `solution` (the grader reads expected values off this — never hand-write answers),
  `shapeRules`, three `hints`, and a `realWorldNote`.
- **A scenario** — a builder in `data/fixtures.js` plus an entry in `XT.data.scenarios`.
- **A bug fix** — add a case to `js/tests.js` first.

## Catalog section → file

| Catalog | Implemented in |
|---|---|
| §7.1 reference forms | `engine/refs.js`, `engine/tokenizer.js`, `engine/parser.js` |
| §7.2 operators, precedence | `engine/parser.js`, `engine/evaluator.js` |
| §7.3 errors | `engine/values.js` (`XErr`), `functions/logic.js` |
| §7.4 EMPTY vs `""` | `engine/values.js`, and the whole first test group |
| §7.5 coercion | `engine/values.js` (`toNum`, `toText`, `compare`, `numToText`) |
| §7.6 dates | `engine/dates.js` |
| §7.8 number formats | `engine/formats.js` |
| §7.9 fill translation, clipping | `engine/parser.js` (`translate`), `engine/refs.js` (`clip`), `ui/grid.js` |
| §3 five layers | `data/fixtures.js` (sheet `layer`), `trainer/grader.js` (`gradeLayers`) |
| §5 patterns P1–P20 | `data/patterns.js`, `engine/shape.js` |
| §9 defects D1–D14 | level 18 in `trainer/levels.js`, `ui/panels.js` (inspector) |
| §10 scenarios S1–S5 | `data/fixtures.js` |
| §11 curriculum | `trainer/levels.js` |
| §12.7 grader | `trainer/grader.js` |

## What has been verified

Headless runners, all green:

- **300** conformance cases: 296 pass, 0 fail, 4 open (`OPEN_QUESTIONS.md`).
- **33 of 34** course exercises pass when graded against their own model answers.
  The 34th is level 18, the defect hunt, which has no model answer by design.
- **38 of 38** formula tasks pass against their own model answers, and their
  result values were read back and checked by eye for sense (no column of
  `#VALUE!`, no column of blanks).
- Every JS file passes `node --check`.

**The browser UI has still not been opened.** If something throws on first load,
the console will say which file.
