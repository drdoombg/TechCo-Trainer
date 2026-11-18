# Formula Catalog — D:\Work Excel Estate

**Purpose:** a blueprint for an HTML Excel simulator / sandbox trainer.
**Built from:** an automated scan of every `.xlsx` / `.xlsm` workbook under `D:\Work`.
**Date of scan:** 2026-11-15

---

## 0. How to read this document

This is not a general Excel tutorial. It is a *description of the specific dialect of Excel
used in this job*, measured rather than guessed, so that an HTML simulator can be built that
drills exactly that dialect and nothing else.

Three things it gives you:

1. **What the engine must support** (§6, §7) — the feature list a JavaScript formula engine
   has to implement before any lesson can run.
2. **What the lessons must teach** (§5, §8, §9) — the formula patterns, in order of how often
   they actually appear in the work.
3. **What the practice data must look like** (§10) — real end-to-end pipelines, column by
   column, ready to become sandbox scenarios.

Everything in §5 and §10 is quoted verbatim from the workbooks. Where a formula is shown with
`SH!`, `R`, `N`, `"S"` it is a *normalised shape* (sheet ref / cell ref / number / string
literal) used for counting; the real example always follows.

---

## 1. Scope of the scan

| Measure | Count |
|---|---|
| Workbooks read (`.xlsx` + `.xlsm`) | **152** |
| Worksheets | **1,816** |
| — of which hidden / veryHidden | **509** (28%) |
| Total formula cells | **7,041,214** |
| Distinct formula texts extracted | **3,973,548** |
| Workbooks containing VBA | **71** |
| Excel Tables (ListObjects) | 108 |
| PivotTables | 32 |
| Workbooks with external workbook links | 20 |
| Legacy files **not** scanned (`.xls`, `.xlsb`) | 15 |

**Not covered by this scan** (see §13): the 15 legacy `.xls`/`.xlsb` files (binary formats —
formulas need Excel itself to extract), VBA macro bodies, Power Query / `M` code, and chart
definitions.

**Role folders scanned:**
`1. Compliance Manager`, `2. Lab Testing Co-ordinator`, `3. Technical Co-ordinator`,
`4. KPI Administrator`, `5. Process Technologist`, `6. Safety Manager`, `QA`,
`xDocuments - Procedures`, `zCompany Details`.

The formula mass is concentrated in three of them:

| Role folder | Formula signature |
|---|---|
| **2. Lab Testing Co-ordinator** | `IF` 779k · `IFERROR` 324k · `VLOOKUP` 232k · `XLOOKUP` 127k · `SUMIF` 39k · `COUNTIF` 29k · `HOUR/MINUTE/SECOND/TIME` 2.7k each |
| **3. Technical Co-ordinator** | `IF` 807k · `RIGHT` 203k · `LEFT` 200k · `VLOOKUP` 54k · `TEXTAFTER`/`TEXTBEFORE`/`VALUE` 7.8k each — **text surgery on machine logs** |
| **4. KPI Administrator** | `IF` 648k · `IFERROR` 408k · `VLOOKUP` 403k · `WEEKDAY` 177k · `COUNTIF(S)` 76k · `SUMIFS` 32k · `AVERAGEIFS` 1.6k — **period/week aggregation** |

---

## 2. The one-sentence finding

> **This estate is not a calculator. It is a hand-built ETL system made of worksheets.**

The evidence:

- **956,499 formula cells contain no function at all** — they are bare pass-throughs like
  `=SamplesinWB!A2` or `='CAMPY MICRO PASTE'!A5`. That is ~1 in 7 of all formulas, and the
  single largest formula shape in the estate (`SH!R`, 867,462 cells across 53 workbooks).
- The top five functions by volume are `IF`, `IFERROR`, `VLOOKUP`, `RIGHT`, `LEFT`. Not one of
  them computes anything. They **route, guard, look up and slice**.
- Arithmetic barely registers: `SUM` 23,631 uses, `ROUND` 9, `AVERAGE` 48. Compare `VLOOKUP`
  at 689,179.
- 509 hidden sheets with names like `Workings Hidden`, `LBI Hidden`, `Line 1 Hidden`,
  `scanMODWB` — explicit intermediate stages.
- The most-referenced sheet names in the whole estate are literally named after the act of
  pasting:

| Referenced sheet | Times referenced |
|---|---|
| `LVALL` | 1,199,998 |
| `Despatch Paste` | 944,804 |
| `Complaints Paste` | 675,305 |
| `Calendars` | 307,326 |
| `MIMS SAVE FILE` | 248,121 |
| `SamplesinCUT` / `SamplesinWB` | 173,124 / 157,220 |
| `Volumes Paste` | 131,912 |
| `ECOLI MICRO PASTE` | 101,920 |
| `Micro paste` | 71,042 |
| `GMP paste` | 65,536 |

So the simulator's whole reason for existing is correct: the skill being trained is
**reshaping pasted exports into a presentable form**, not finance maths.

---

## 3. The architecture every workbook follows

Nearly every significant workbook is the same five-layer machine. Sheet names across 152
workbooks fall into these roles with remarkable consistency:

```
  ┌──────────────────────────────────────────────────────────────────┐
  │ 1. LANDING / PASTE            raw export dumped in, untouched    │
  │    'Micro paste'  'Complaints Paste'  'LBI Paste'  'MIMS SAVE    │
  │    FILE'  'Despatch Paste'  'KILL DATA PASTE'  'Bone QA Check    │
  │    Paste'  'FDC Paste'  'LVALL'  'Samplesin*'                    │
  └──────────────────────────────────────────────────────────────────┘
                              ↓  pass-through + guard
  ┌──────────────────────────────────────────────────────────────────┐
  │ 2. MIRROR / MOD              1:1 copy that cleans blanks & zeros │
  │    'scanMODWB'  'scanMODCUT'  'Despatch Hidden'  'Data Extract'  │
  │    'Complaints Paste Review'                                     │
  └──────────────────────────────────────────────────────────────────┘
                              ↓  row selection
  ┌──────────────────────────────────────────────────────────────────┐
  │ 3. FILTER / SELECT           today's rows, this line, this shift │
  │    'FILTERWBTODAY'  'FILTERCUT9AM'  'FILTERPREVIOUSDAY'          │
  │    'FILTERFRESHKILL'                                             │
  └──────────────────────────────────────────────────────────────────┘
                              ↓  enrichment by key
  ┌──────────────────────────────────────────────────────────────────┐
  │ 4. WORKINGS (hidden)         joins to reference data, derived    │
  │    'Workings Hidden'  'LBI Hidden'  'Line 1 Hidden'  'Micro      │
  │    HELPER TAB'  'Final LBI Values'  'Ecoli Final Data'           │
  │    ← joins against: 'Calendars'  'picklists'  'List Menu'        │
  │      'Label Info'  'REF'  'Error Coding'                         │
  └──────────────────────────────────────────────────────────────────┘
                              ↓  presentation
  ┌──────────────────────────────────────────────────────────────────┐
  │ 5. OUTPUT                    the thing a human/printer consumes  │
  │    'FINALEMAIL WB'  'Label for Micro Neck'  'Reporting Template' │
  │    'Site KPI Reporting'  'QA Sheet'  'Cover'  'Home'             │
  │    'SIS_Micro_Sub_DATE' (a file to upload back to the LIMS)      │
  └──────────────────────────────────────────────────────────────────┘
```

**This five-layer shape is the single most important thing the simulator should teach.** Not
`VLOOKUP` in isolation — the habit of building *layers*, each with one job.

Supporting conventions observed:

- **`picklists`** sheet (26 workbooks) holding one column per dropdown, exposed as defined
  names: `Failure_type`, `Shift`, `Water`, `Environmental_Swabbing`,
  `Finished_Product_Microbiology`, `Raw_Material_Microbiology`, `Regulatory_Microbiology`,
  `Answer`. All of the form `picklists!$A$2:$A$16`.
- **`Calendars`** sheet mapping dates → week number → period → retailer calendar. The join key
  is built by concatenation: `=F2&B2`, then read back with `=VLOOKUP(V2&AA2,Calendars!$A$1:$H$470,8,0)`.
- **`Cover` / `Home` / `How to Use It`** sheets as the user-facing front panel, usually driven
  by `COUNTIFS` against the data sheets.
- **`Error Coding`** sheet that collects the workbook's own validation failures into one place.
- Version numbers live in the filename (`v17`, `V20`, `v49`), not inside the file.

---

## 4. Quantitative function catalog

Ranked by number of formula cells containing the function. `files` = how many of the 152
workbooks use it at all — this is the better signal for *generality*.

### Tier 1 — the irreducible core (teach first, in this order)

| Function | Uses | Files | Role in this estate |
|---|---:|---:|---|
| `IF` | 2,235,460 | 73 | Blank-guard, routing, label selection |
| `IFERROR` | 750,917 | 38 | Swallow `#N/A` from lookups → `""` |
| `VLOOKUP` | 689,179 | 25 | The join. Almost always `,FALSE` / `,0` |
| `XLOOKUP` | 128,323 | 21 | The modern join; used for multi-key via `&` |
| `COUNTIF` | 71,070 | 25 | Existence check, duplicate numbering |
| `COUNTIFS` | 39,511 | **39** | Dashboard counting (most *widespread* aggregator) |
| `SUMIFS` | 32,717 | 26 | Period/date-window totals |
| `SUM` | 23,631 | 33 | Row and block totals |
| *(no function)* | 956,499 | — | Bare pass-through `=Sheet!A1` |

### Tier 2 — text surgery on machine output

| Function | Uses | Files | Role |
|---|---:|---:|---|
| `RIGHT` | 204,153 | 21 | Trailing status token off a log line |
| `LEFT` | 200,242 | 4 | Leading `JOB…` prefix test |
| `TEXTAFTER` | 10,081 | 17 | Field *n* of a space-delimited log line |
| `TEXTBEFORE` | 7,830 | 5 | Paired with the above |
| `VALUE` | 7,915 | 8 | Text-number → number after slicing |
| `TEXT` | 3,346 | 30 | `TEXT(AC2,"MMM-YY")` period labels |
| `SUBSTITUTE` + `CHAR` | 1,564 + 1,564 | 1 | `SUBSTITUTE(D1,CHAR(160)," ")` — strip non-breaking spaces from a PLU export |
| `LEN` | 1,127 | 6 | Blank detection via `LEN(TRIM(x))=0` |
| `FIND` | 103 | 1 | — |

### Tier 3 — dates, shifts and the week grid

| Function | Uses | Files | Role |
|---|---:|---:|---|
| `WEEKDAY` | 177,467 | 13 | `(date - WEEKDAY(date,1)) + 7` → week-ending date |
| `TODAY` | 3,872 | 24 | "Today's samples" filters |
| `HOUR` `MINUTE` `SECOND` `TIME` | ~2,775 each | 12–14 | Rebuild a time from parts to compare against a shift cut-off |
| `MONTH` | 41 | 5 | — |
| `NETWORKDAYS` | 11 | 1 | — |
| `WEEKNUM` | **1** | 1 | *Note: the estate does **not** use `WEEKNUM`. Week numbers come from the `Calendars` lookup table instead — the retailer weeks don't match ISO.* |

### Tier 4 — index/match and ranking

| Function | Uses | Files | Role |
|---|---:|---:|---|
| `ROW` | 18,468 | 10 | Row sequence generator inside array tricks |
| `MATCH` | 13,108 | 20 | Position lookup |
| `INDEX` | 11,814 | 21 | Paired with `MATCH`; also `INDEX/SMALL/IF` "get all matches" |
| `MIN` / `MAX` | 5,870 / 3,688 | 12 / 13 | Earliest/latest date in a pasted block |
| `SMALL` | 396 | 1 | The *n*-th match trick |
| `RANK` | 896 | 11 | Worst-5 tables |
| `SUMPRODUCT` | 1,763 | 3 | Pre-`COUNTIFS` multi-criteria counting |
| `AVERAGEIFS` | 1,689 | 7 | Mean result per week per site |
| `INDIRECT` | 554 | 1 | Sheet name built from a cell |
| `OFFSET` | 7 | 5 | `OFFSET(AG3,,MATCH(S3,AH3:AM3,0))` retailer column pick |

### Tier 5 — present but rare (support in the engine, don't build lessons around)

`AND` 5,173 · `OR` 100 · `NOT` 346 · `COUNT` 4,119 · `COUNTA` 176 · `COUNTBLANK` 74 ·
`ISTEXT` 238 · `ISERROR` 193 · `ISBLANK` 139 · `ISNUMBER` 12 · `IFNA` 45 · `NA` 17 ·
`INT` 39 · `ROUND` 9 · `ROWS` 12 · `TRANSPOSE` 3 · `UNIQUE` 2 · `TAKE` 12 ·
`ANCHORARRAY` 12 · `IMAGE` 87 · `CONCATENATE` 2 · `UPPER` 2 · `LOOKUP` 355 · `NOW` 5.

> **Design note:** `UNIQUE` appears twice and `FILTER` never. Dynamic arrays are essentially
> absent — the work is done with fill-down and helper columns. The simulator should therefore
> make **fill-down the primary interaction**, not array spilling.

### Formula mechanics distribution

| | Count |
|---|---:|
| Normal formulas | 3,883,029 |
| Shared formulas (fill-down groups) | 66,542 masters → 3.07M child cells |
| **Legacy array formulas (CSE)** | **23,977** |

---

## 5. The pattern catalog

These are the actual recurring *shapes*, ranked by occurrence. Each entry is a candidate
lesson. `n` = cells, `files` = workbooks.

---

### P1 · Bare pass-through — `=OtherSheet!A2`
**n = 867,462 · files = 53 · THE most common formula in the estate**

```excel
=SamplesinWB!A2
='CAMPY MICRO PASTE'!A5
```

**What it does:** mirrors a column from the landing sheet onto a working sheet.
**Why it exists:** keeps the pasted export pristine. Everything downstream points at the
mirror, so re-pasting never breaks the formulas. This is the estate's version of "don't edit
your source data".
**Simulator must support:** cross-sheet references, relative fill-down across rows,
sheet names with spaces requiring `'quotes'`.

---

### P2 · Zero-suppressing mirror — `=IF(Src!A6>0,Src!A6,"")`
**n = 472,402 · files = 2 (but 944k refs to `Despatch Paste`)**

```excel
=IF('Despatch Paste'!A6>0,'Despatch Paste'!A6,"")
```

**What it does:** mirror, but show nothing instead of `0`.
**Why it exists:** exports pad empty cells with zeros; zeros print on labels and look like data.
**Teaching point:** `>0` is a *numeric* test being used on possibly-text data — a latent bug
(see §9). The honest form is `=IF(Src!A6="","",Src!A6)`.

---

### P3 · Blank-guarded mirror — `=IF(A1="","",Src!B2)`
**n = 399,998 + 328,036 variants · files = 6**

```excel
=IF(E1="","",LVALL!C2)
=IF('Volumes Paste'!A1="","",'Volumes Paste'!A1)
```

**What it does:** only pull if the anchor cell has something in it.
**Why it exists:** stops a 100,000-row filled-down column from showing junk below the data.
**Simulator must support:** `""` as a distinct value from a truly empty cell — the single most
important semantic subtlety in this whole estate (see §7.4).

---

### P4 · Belt-and-braces mirror — `=IFERROR(IF(Src!A1="","",Src!A1),"")`
**n = 329,778 · files = 1 (`Complaints Analysis Weekly Database2.xlsm`)**

```excel
=IFERROR(IF('Complaints Paste'!A1="","",'Complaints Paste'!A1),"")
```

**Teaching point:** the `IFERROR` here is defensive cargo-cult — a mirror of a cell cannot
error. Good lesson on *reading* a formula to decide whether a wrapper earns its place.

---

### P5 · The guarded join — `=IFERROR(VLOOKUP(key,Sheet!$A:$DN,n,FALSE),"")`
**n = 237,873 · files = 9 — the workhorse**

```excel
=IFERROR(VLOOKUP(C18,PASTE!$A:$DN,2,FALSE),"")
=IFERROR(VLOOKUP(A16,'Risk Assessment'!A:B,2,0),"")
```

**What it does:** enrich a row from a reference table; blank if not found.
**Why it exists:** every single pipeline needs it — product code → description, barcode →
product, farm → site.
**Note the house style:** whole-column ranges `$A:$DN`, exact match always spelled `FALSE`
or `0`, result always `IFERROR`-wrapped to `""`.
**Simulator must support:** whole-column ranges, `col_index_num` counted from the *first*
column of the range (the classic off-by-one teaching moment), absolute/relative `$` mixing.

---

### P6 · Log-line slicing — `RIGHT` / `LEFT` classification
**n = 400,000 across two shapes · files = 2 (`M&S Vision` LV Verification)**

```excel
=IF(RIGHT(LVALL!C1,29)="FLASHLABEL; Invalid: Mismatch","Missing Sleeve",
    IF(RIGHT(LVALL!C1,8)="Mismatch","Mismatch",""))
=IF(LEFT(LVALL!C1,3)="JOB",LEFT(LVALL!C1,45),"")
```

**What it does:** turns a vision-system text log into typed columns — a one-column CSV parsed
by character position.
**Why it exists:** the export is a flat log, not a table.
**Teaching point:** magic numbers (`29`, `8`, `45`, `3`) are counted by hand and break when
the vendor changes a message. The modern replacement is right there in the same file (P7).

---

### P7 · Delimiter-based field extraction
**n = 6,256 × 2 · files = 2 / 17**

```excel
=IFERROR(_xlfn.TEXTAFTER(F1," ",3),"")
=IFERROR(VALUE(_xlfn.TEXTBEFORE(G1," ",1)),"")
=_xlfn.TEXTAFTER(M2," ",-1)          ← negative n = count from the end
```

**What it does:** "give me the 3rd space-delimited field", then coerce to a number.
**Why it matters pedagogically:** P6 and P7 solve the same problem in the same workbook.
Showing them side by side is the best single lesson in the catalog.
**Simulator must support:** `TEXTBEFORE`/`TEXTAFTER` with negative instance numbers, and the
`_xlfn.` prefix (see §7.7).

---

### P8 · Week-ending date lookup
**n = 177,447 · files = 1, but `WEEKDAY` appears in 13**

```excel
=IF(A4="","-",VLOOKUP(($B4-WEEKDAY($B4,1))+7,Calendars!$H:$J,3,FALSE))
```

**What it does:** rounds a date to its week-ending Saturday, then looks that up in the retailer
calendar to get the period/week label.
**Why it exists:** the business runs on retailer calendars, not ISO weeks. `WEEKNUM` is useless
here — hence 1 use in the whole estate against 177k uses of this pattern.
**Simulator must support:** date serial arithmetic, `WEEKDAY(d, return_type)`, lookup into a
date column.

---

### P9 · Multi-key join by concatenation
**n = 102,397 (`VLOOKUP`) + 8,564 (`XLOOKUP`) · files = 6 / 9**

```excel
=VLOOKUP(V2&AA2,Calendars!$A$1:$H$470,8,0)
=IFERROR(_xlfn.XLOOKUP(FILTERWBTODAY!$B2&FILTERWBTODAY!G2,
         'MIMS SAVE FILE'!$M:$M&'MIMS SAVE FILE'!V:V,
         'MIMS SAVE FILE'!J:J),"")
```

**What it does:** join on two columns at once (barcode + shift, week + site) by gluing them
into one key.
**Two dialects in the same estate:** `VLOOKUP` needs a *physical* helper column holding
`=F2&B2`; `XLOOKUP` can build the key array inline with `&`. Teach both — the helper-column
version is what they'll meet in old files.
**Simulator must support:** `&` concatenation producing arrays when applied to ranges, and
`XLOOKUP` over a computed lookup array.

---

### P10 · `XLOOKUP` column-fan-out
**n = 58,772 + 40,611 guard variants · files = 9**

```excel
=IFERROR(_xlfn.XLOOKUP(FILTERWBTODAY!$B2,'MIMS SAVE FILE'!$M:$M,'MIMS SAVE FILE'!A:A),"")
… repeated for B:B, C:C, D:D … AK:AK   (37 columns, same key, one column each)
=IF(IFERROR(_xlfn.XLOOKUP(…,'MIMS SAVE FILE'!S:S),"")=0,"")   ← zero-blanking variant
```

**What it does:** rebuilds an entire 37-column record next to a scanned barcode.
**Why it exists:** the lab scanner gives a barcode; the submission file needs the full product
record in a fixed column order.
**Teaching point:** this is the right answer, done 37 times by hand. A great "now do it once
and fill right" lesson — and a great argument for the `$` anchoring that makes it possible.

---

### P11 · Existence flag then conditional join
**n = 30,000 + 10,000 · files = 1**

```excel
=IF(K2="","",IF(COUNTIF('Volumes Paste'!$A:$A,'Complaints Paste Review'!K2)>=1,1,""))
=IFERROR(IF(L2=1,VLOOKUP(K2,'Volumes Paste'!$A:$G,7,FALSE),
         IF(N2=1,VLOOKUP(M2,'Volumes Paste'!$B:$G,6,FALSE),
         IF(P2=1,VLOOKUP(O2,'Volumes Paste'!$C:$G,5,FALSE),
         "Check"&" "&"-"&" "&O2))),"Check"&" "&"-"&" "&O2)
```

**What it does:** three candidate keys are tested for existence with `COUNTIF`; the first one
that exists is used for the real join; otherwise emit a human instruction `"Check - <code>"`.
**Why it exists:** complaint records identify a product by one of three different code systems
depending on retailer.
**Teaching point:** cascading fallback keys — a genuinely sophisticated pattern that deserves
its own lesson. Also note the output is *a message to a person*, not an error value. That is
the estate's error-handling philosophy throughout.

---

### P12 · Date-window `SUMIFS` across a date *row*
**n = 11,379 + 5,124 + 1,440 · files = 3**

```excel
=SUMIFS('Volumes Paste'!17:17,'Volumes Paste'!$13:$13,">=" &K$14,
                              'Volumes Paste'!$13:$13,"<="&K$15)
=SUMIFS('FDC Paste'!$AB2:$AB10000,'FDC Paste'!$B2:$B10000,">=" &$K$4,
                                  'FDC Paste'!$B2:$B10000,"<=" &$M$4,
                                  'FDC Paste'!$N2:$N10000,$G$62,
                                  'FDC Paste'!$O2:$O10000,$H$62)
```

**What it does:** totals a week/period window. Note the first one sums a **row** (`17:17`) with
criteria on **row 13** — the pasted volumes file is transposed, dates across the top.
**Simulator must support:** criteria as a *string built by concatenation* (`">=" & K14`) —
this is where most learners fail. Also row-oriented ranges, not just columns.

---

### P13 · Multi-criteria `COUNTIFS` dashboard cell
**n = 10,416 + 3,458 + 2,357 + 1,094 · files = 2–3**

```excel
=COUNTIFS('Micro paste'!$M:$M,$D$86,
          'Micro paste'!$H:$H,$C$86,
          'Micro paste'!$B:$B,'Reporting Template'!$Q$23,
          'Micro paste'!$AA:$AA,"Yes")
=COUNTIFS('Micro Paste'!$K:$K,B$133,'Micro Paste'!$E:$E,$A$132,
          'Micro Paste'!$EA:$EA,$D$132,
          'Micro Paste'!$B:$B,">=" &$B$1,'Micro Paste'!$B:$B,"<="&$C$1)
```

**What it does:** one dashboard number = count of paste rows matching test code + site + week +
flag. The entire KPI reporting layer is built from these.
**Simulator must support:** `COUNTIFS` with 4–6 criteria pairs, mixed absolute anchoring
(`B$133` vs `$A$132` — row-locked vs column-locked so one formula fills in both directions).
**This is the single highest-value lesson for the KPI Administrator role.**

---

### P14 · Duplicate-instance numbering (unique key manufacture)
**n = 7,614 + 2,968 · files = 2**

```excel
=C5&":"&COUNTIF($C5:C$16,C5)
=H19&COUNTIF($H$19:H19,H19)
```

**What it does:** turns a repeating code into `CODE:1`, `CODE:2`, `CODE:3` by counting how many
times it has appeared *so far* — the expanding-range trick (`$C5:C$16` grows as it fills).
**Why it exists:** `VLOOKUP` only finds the first match, so a unique key has to be manufactured
to retrieve the 2nd, 3rd … sample for the same product.
**Simulator must support:** the expanding range. Teaching the `$` placement that makes
`$C5:C$16` grow one way and not the other is essential and non-obvious.

---

### P15 · "Get the *n*-th match" — `INDEX` + `SMALL` + `IF` array
**n = 2,000+ across the Barcode toolkits · files = 1 (array formula)**

```excel
{=IFERROR(INDEX(FILTERCUT9AM!Q:Q,
   SMALL(IF((FILTERCUT9AM!B:B=FILTERCUT9AM!B2)*(FILTERCUT9AM!F:F=FILTERCUT9AM!F2),
            ROW(FILTERCUT9AM!Q:Q)),
         COUNTIFS(FILTERCUT9AM!$B$2:B2,FILTERCUT9AM!B2,
                  FILTERCUT9AM!$F$2:F2,FILTERCUT9AM!F2))),"")}
```

**What it does:** retrieves *all* matching rows, not just the first, by combining the boolean
product of two criteria, `ROW()` as a position generator, `SMALL()` to pick the k-th, and
`COUNTIFS` on an expanding range to supply k.
**This is the ceiling of the estate's complexity.** It is the pre-`FILTER` idiom.
**Simulator must support:** CSE array evaluation, boolean×boolean multiplication as AND,
`ROW(range)` returning an array.

---

### P16 · Boundary-match lookup (date falls inside a range)
**n = 9,100 · files = 5**

```excel
=INDEX(Calendars!F:F, MATCH(1, (C2>=Calendars!G:G)*(C2<=Calendars!H:H), 0))
```

**What it does:** finds which calendar week a date belongs to, where the calendar is stored as
start/end date pairs rather than a single key.
**Teaching point:** `MATCH(1, (…)*(…), 0)` — the array-as-AND idiom again. Also the honest
alternative: `XLOOKUP(C2, Calendars!G:G, Calendars!F:F, , -1)` with approximate match.

---

### P17 · Match-verification / self-audit
**n = ~3,000 · files = 4**

```excel
=K2&"_"&M2&"_"&O2                                    ← build a signature
=IF(D2=Q2,1,0)                                       ← compare to the expected one
=IF(R2=1,"DATA MATCH","PLEASE REVIEW AND AMEND")     ← speak to the human
=IF(COUNTIF(Volumes!I:I, B5) > 0, "CORRECT", "INCORRECT")
```

**What it does:** the workbook checks itself and prints a verdict for the operator.
**Why it matters:** this is the estate's quality culture encoded in formulas. The simulator
should treat "add a self-check column" as a graded objective, not an afterthought.

---

### P18 · Score → band (nested `IF` ladder)
**n = 86,282 + 2,532 · files = 3–4**

```excel
=IF('QTR 1'!$K$16=1,1,IF('QTR 1'!$K$16=2,2,""))
=IF(D16="N",0,IF(D16="Y",2,0))
=IF(S3="Aldi",U3,IF(S3="CO-OP",U3,IF(S3="Tesco",V3,IF(S3="M&S","2015/16",
   IF(S3="Sainsbury",V3,IF(S3="Asda",V3,))))))
```

**What it does:** maps a category to a value through nested `IF`s — depth observed up to 6.
**Teaching point:** the last example is a lookup table written as code. The lesson is
converting an `IF` ladder into a two-column table + `VLOOKUP`/`XLOOKUP`. Note the trailing
`,))))` — an empty `value_if_false`, which returns `FALSE`, not blank. A real bug in a live
file.

---

### P19 · Sequence and spine generators
**n = 19,683 (`+R+N`) + 17,238 (`ROW()`) + 1,053 · files = 17 / 6 / 6**

```excel
=+K13+1                                        ← incrementing week/day column
=ROW()                                         ← row ordinal
=+H3+1   /   =+G2+6                            ← build Calendars start/end dates
=IF(A2="","",IF((A2+1)<=$L$1,(A2+1),""))       ← self-limiting counter
=IF(ROW()<COUNTA($B:$B)-49,"",ROW()-(COUNTA($B:$B)-50))   ← last-50-rows window
```

**What it does:** manufactures the row/column spine a report hangs off — dates, batch numbers,
rolling windows.
**Simulator must support:** `ROW()` with no argument, `COUNTA` on a whole column.

---

### P20 · Rate and percentage finishers
**n = 2,855 + 1,083 + 1,074 · files = 5–12**

```excel
=IFERROR((K58/K57)*100,"")
=T22/S22
=SUMIFS($Y:$Y,$V:$V,$V2,$AA:$AA,$AA2)/COUNTIFS($V:$V,$V2,$AA:$AA,$AA2)
=IF(U2>500,1,0)   then   =AVERAGEIFS($U:$U,$V:$V,$V2,$AA:$AA,$AA2)
```

**What it does:** the only real arithmetic in the estate — % compliance, % over limit, mean
result. Note the idiom: flag each row `1/0` first, then average the flags to get a rate.
**Teaching point:** `IFERROR(x/y,"")` for divide-by-zero, and the flag-then-average pattern as
an alternative to `COUNTIFS/COUNTIFS`.

---

## 6. Non-formula features the simulator needs

### 6.1 Data validation — 218 rules

| Type | Count | Example |
|---|---:|---|
| `list` | 159 | `$AI$2:$AI$5`, `picklists!$E$2:$E$15` |
| `whole` (integer) | 27 | `≥ 0`, `≥ 1` |
| `date` | 7 | `≥ 30682` (serial for 1984-01-01) |
| `decimal` | 6 | `≥ -100` |
| `custom` | 2 | `SUM($J29:$U29)<=$I29` |

Dropdowns are the primary data-entry control in this estate. A simulator without working
dropdowns bound to a `picklists` range will not feel like the real thing.

### 6.2 Conditional formatting — 292,327 rules

| Rule type | Count | Note |
|---|---:|---|
| `cellIs` / `between` | 180,990 | **often with text criteria: `"<500"`, `"<=2"`** |
| `cellIs` / `greaterThanOrEqual` | 88,532 | e.g. `1000` — the E. coli action limit |
| `cellIs` / `equal` | 12,482 | `"AGRI"`, `"TRANS"` |
| `expression` (formula) | 8,465 | `$J16="Removed"`, `$J$7<TODAY()` |
| `duplicateValues` | 678 | |
| `containsText` | 339 | `NOT(ISERROR(SEARCH("p",D7)))` |
| `containsBlanks` | 32 | `LEN(TRIM(A2))=0` |
| `colorScale` / `dataBar` | 14 / 1 | essentially unused |

**Takeaway:** colour *is* the output. A result sheet that isn't red/amber/green hasn't done its
job. The simulator must render CF, and must support the `expression` type with `$`-anchored
row-relative formulas (`$J16="Removed"` applied over `A66:W110`).

### 6.3 Defined names

Two populations, both worth teaching:

**Good:** semantic picklist ranges — `Failure_type`, `Shift`, `Water`, `Answer`,
`Environmental_Swabbing`, `Finished_Product_Microbiology` → `picklists!$A$2:$A$16`.

**Bad (and real):** `boom`, `xxxxx`, `jgjgjgjhgjhg`, `ghhghg`, `bbb`, `tESt`, `phil`,
`_____jul1` — most pointing at `#REF!`. These are inherited corpses from decades of copied
workbooks. 229 defined names in one file alone.

**Lesson:** Name Manager hygiene, and how `#REF!` names propagate silently.

### 6.4 Excel Tables (108 ListObjects)

Used, but inconsistently. Calculated columns observed:

```excel
Table1[Change]    = SamplesinWB!A2            ← structured table, unstructured formula
Table1[Time VALUE]= TIME(HOUR(P2),MINUTE(P2),SECOND(P2))
Table4[PRODUCT]   = IFERROR(_xlfn.XLOOKUP(FILTERWBTODAY!#REF!,'MIMS SAVE FILE'!$M:$M,…),"")
```

And consumed with structured references:

```excel
=IF(Table613[[#This Row],[MANUFTIME]]<TIME(HOUR(scanMODCUT!$R$1),…),…)
=IFERROR(N2+Table613[[#This Row],[SUPPLIER]],"")
```

**Simulator should support:** `Table[[#This Row],[Column]]` syntax at minimum, because it
appears in live production files mixed with A1 refs in the same formula.

### 6.5 Other

- **71 workbooks contain VBA** — the simulator cannot and should not replicate this, but
  lessons should mark the boundary: *"in the real file a button does this; here you do it by
  hand."*
- **32 PivotTables, 14 color scales** — low priority.
- **`IMAGE()` function, 87 uses in 3 files** — QR-code label sheets pulling images by URL.
- **External workbook links in 20 workbooks**, resolving to `H:\Technical\…` network paths,
  `\\2sfg.net\filestore\RedirectedFiles\…`, SharePoint URLs, and — revealingly — Outlook
  attachment temp folders (`AppData\Local\Microsoft\Olk\Attachments\…`). Several are stale,
  giving `[1]Pivot!#REF!`. **The `[n]` external reference syntax must at least parse.**

---

## 7. Engine requirements (the hard part)

Derived from what actually appears. Anything not listed here can be skipped in v1.

### 7.1 Reference forms that must work

| Form | Example | Seen |
|---|---|---|
| Relative / absolute / mixed | `A1`, `$A$1`, `A$1`, `$A1` | everywhere; mixed anchoring is load-bearing in P13/P14 |
| Range | `K16:M56` | everywhere |
| Whole column | `$A:$DN`, `C:C` | the house style for lookups |
| Whole row | `17:17`, `$13:$13` | P12 |
| Cross-sheet | `Sheet1!A1` | 2.4M+ refs |
| Quoted sheet | `'MIMS SAVE FILE'!$M:$M` | names with spaces & `&` are common |
| External workbook | `[1]CODES!$G$4:$I$4975`, `'[5]Complaints Trending'!$J$11` | 20 workbooks |
| Structured table | `Table4[[#This Row],[SUPPLIER]]` | mixed with A1 in one formula |
| Defined name | `picklists!$E$2:$E$15` via `Environmental_Swabbing` | 26 workbooks |
| Error refs | `#REF!` **inside live formulas** | thousands — must evaluate, not crash |

### 7.2 Operators

`+ - * / ^` · `& ` (concatenation — heavily used) · `= <> > < >= <=` ·
**unary leading `+`** (`=+K13+1`, `=+'Site KPI Reporting'!Q12` — 19,683+14,568 cells; a Lotus
habit that survives, and it must parse) · `%` · range `:` · union `,` (in defined names) ·
implicit intersection in legacy array contexts.

### 7.3 Error values

`#N/A` `#REF!` `#VALUE!` `#DIV/0!` `#NAME?` `#NULL!` `#NUM!` — all must exist as first-class
values, propagate through operators, and be catchable by `IFERROR`/`ISERROR`/`IFNA`.

### 7.4 ⚠ The `""` vs truly-empty distinction

**This is the most important semantic in the entire estate.** 1.2M+ formulas end in `,"")`.

The simulator must reproduce Excel's actual behaviour:

| Test on a cell containing `=""` | Excel result |
|---|---|
| `ISBLANK(x)` | `FALSE` |
| `x=""` | `TRUE` |
| `COUNTA(x)` | `1` |
| `COUNT(x)` | `0` |
| `COUNTBLANK(x)` | `1` |
| `x>0` | **`TRUE`** — `""` is text, and in Excel any text is greater than any number. This is why P2's `>0` guard is a latent bug (D5) |
| `SUM(x)` | `0`, no error |
| `x&"!"` | `"!"` |

Get this wrong and every lesson built on P2/P3/P5 will give the wrong answer. Build a
conformance test suite for this cell semantics before writing any lesson.

### 7.5 Type coercion rules needed

- Text that looks like a number in a comparison: `"10"=10` → `FALSE`; `"10">5` → `TRUE`
  (text always greater than any number).
- `VALUE("  12.5 ")` → `12.5`; `VALUE("abc")` → `#VALUE!`.
- `&` coerces numbers to text using General formatting — `=1/3&""` gives
  `"0.333333333333333"` (15 sig figs).
- Criteria strings: `">=" & 46234` must be parsed by `SUMIFS`/`COUNTIFS` as an operator plus a
  number. Observed in 20k+ cells.
- Booleans in arithmetic: `TRUE*1=1`. P15 and `IFERROR(IF(…,AND(…),"")*1,0)` depend on it.

### 7.6 Dates

- 1900 serial system; `1 = 1900-01-01`.
- Times as fractional days; `TIME(h,m,s)` returns a fraction.
- `WEEKDAY(d, return_type)` with `return_type` 1 (the only one used).
- Date arithmetic on serials: `(B4 - WEEKDAY(B4,1)) + 7`, `O2-2`, `U2-31`, `$F$2+365`.
- Display is a *format*, not a value — the estate uses `dd/mm/yyyy` (UK), `m/d/yyyy`, `ddd`,
  `dddd`, `yyyy.mm.dd hh:mm:ss`, `[$-F400]h:mm:ss AM/PM`.
- `TEXT(d,"MMM-YY")` for period labels.
- **Watch for `0/1/1900`** — appears in live formulas as `IF(D2=0/1/1900,…)`, which Excel
  parses as the *arithmetic* `0÷1÷1900 = 0`, not a date. See §9.

### 7.7 The `_xlfn.` prefix

Every modern function appears in the XML as `_xlfn.XLOOKUP`, `_xlfn.TEXTAFTER`,
`_xlfn.TEXTBEFORE`, `_xlfn.ANCHORARRAY`, `_xlfn.TAKE`. The simulator's parser should accept
both `XLOOKUP(` and `_xlfn.XLOOKUP(` so that formulas pasted straight out of a file work. Worth
a short lesson too: this prefix is why a file opened in an older Excel shows `#NAME?`.

### 7.8 Number formats to implement

`General` · `dd/mm/yyyy` · `m/d/yyyy` · `ddd` · `dddd` · `mmmm` · `yyyy` ·
`yyyy.mm.dd hh:mm:ss` · `[$-F400]h:mm:ss AM/PM` · `£0.00` · `£#,##0.00` · accounting
(`_("£"* #,##0.00_);…`) · `0.0%` · `0.0` / `0.000` / `0.0000E+00` · `0;[Red]0` · `@` (text) ·
`d/m/yy;@`.

### 7.9 Calculation engine

- Dependency-graph recalculation; cycles must be detected and reported, not hung on.
- **Fill-down / fill-right with reference translation is the core interaction**, not a nicety.
  The estate has 3.07M shared-formula child cells and essentially no dynamic arrays.
- Legacy CSE array formulas (23,977 cells) — needed for P15 and P16.
- Whole-column ranges over a 1,048,576-row grid must not be evaluated literally. Track the used
  range per sheet.

---

## 8. What *isn't* here (so don't build it)

Measured absences, each one a scope saving:

- **No `FILTER`, no `SORT`, no `SEQUENCE`, no `LET`, no `LAMBDA`.** `UNIQUE` appears twice.
  Dynamic arrays are not part of this dialect.
- **No financial functions at all** — no `NPV`, `PMT`, `IRR`, `XIRR`.
- **No statistical depth** — no `STDEV`, `MEDIAN`, `PERCENTILE`, `CORREL`, `FORECAST`.
- **Almost no rounding** — `ROUND` 9 uses, `ROUNDUP`/`ROUNDDOWN`/`MROUND` zero.
- **No `SWITCH`, no `IFS`, no `CHOOSE`.** Nested `IF` does that work.
- **No `TRIM`/`CLEAN` as functions** (only inside CF formulas) despite obvious need — see §9.
- **No `XMATCH`, no `SORTBY`, no `GROUPBY`/`PIVOTBY`.**
- Goal Seek, Solver, Scenario Manager, Power Pivot, array constants: absent.

> A simulator that implements §7 and the 20 patterns in §5 covers **>99% of the formula mass of
> this entire estate.** That is a tractable v1.

---

## 9. Real defects found — the honest teaching material

These are live bugs in production files. They are the best lesson content available because
they are *theirs*, and fixing them in a sandbox has transfer value.

| # | Defect | Evidence | The lesson |
|---|---|---|---|
| D1 | **`#REF!` inside live formulas** | `=IFERROR(VLOOKUP(V4&AA4,#REF!,8,0),"NA")` (4,051 cells); `Table4[PRODUCT] = …XLOOKUP(FILTERWBTODAY!#REF!,…)`; `=SUM(#REF!)`; `+'QTR 2 Old'!#REF!` (1,504 cells) | What `#REF!` means, how deleting a column creates it, how `IFERROR` **hides it forever** |
| D2 | **`IFERROR` masking real breakage** | D1 above — the lookup is permanently broken but the sheet shows `"NA"` and looks fine | Use `IFNA` to catch "not found" while letting structural errors shout |
| D3 | **`IF(D2=0/1/1900,…)`** | `Complaints Paste Review!C2`, `AL2` | `0/1/1900` is division, = `0`. A date literal needs `DATE(1900,1,1)` or a cell |
| D4 | **Empty `value_if_false`** | `=IF(S3="Aldi",U3,…IF(S3="Asda",V3,))` | Returns `FALSE`, not blank. Shows as `FALSE` on the report |
| D5 | **`>0` as a blank test on text data** | P2, 472,402 cells | Text compares greater than any number in Excel; the guard doesn't do what it looks like it does |
| D6 | **Magic-number string slicing** | `RIGHT(…,29)`, `LEFT(…,45)` — 400k cells | Breaks silently when the vendor changes a message. Compare with `TEXTAFTER` in the same workbook |
| D7 | **Non-breaking space contamination** | `=SUBSTITUTE(D1,CHAR(160)," ")` then `VALUE()` | Why a number pasted from a web/PLU export won't add up |
| D8 | **Orphan defined names** | `boom`, `xxxxx`, `jgjgjgjhgjhg` → `#REF!`; 229 names in one file | Name Manager as a maintenance task |
| D9 | **Version control by filename** | `v7…v15` of Barcode Sample Submission; `Microbiology1`…`77`; `Copy of Nono detail investigation` ×7 | Why the trainer teaches a single file with a dated paste log instead |
| D10 | **External links to Outlook temp folders** | `…\AppData\Local\Microsoft\Olk\Attachments\ooa-16d79c87…` | A link into a temp attachment path dies the moment Outlook cleans up |
| D11 | **Duplicated toolkits across folders** | `Campy WB and Necks labels toolkit v17.xlsm` byte-identical in `Caeca and Necks/` and `HELPING TOOLS/` (441,660 formulas each) | Single source of truth |
| D12 | **Redundant wrappers** | `IFERROR(IF(Src!A1="","",Src!A1),"")` — 329,778 cells | Read a formula critically before copying it 300,000 times |
| D13 | **Whole-column lookups at scale** | `$A:$DN` over 1M rows × 118 columns, in 237k formulas | Why the file takes 40 seconds to open, and what to do instead |
| D14 | **Pasted-over formula ranges** | `MIMS DAILY WB OLD!A2` keyed on `FILTERWBTODAY!B30` while the 9AM twin uses `B2` | Off-by-28 drift between copies of the same sheet |

---

## 10. Case studies — five real pipelines to become sandbox scenarios

Each of these is a complete, working, nameable job. They are ordered by teaching difficulty and
are offered as the simulator's scenario set.

---

### S1 · Barcode Sample Submission — *the enrichment fan-out*
`2. Lab Testing Co-ordinator/Apps/.../Barcode Sample Submission v15.xlsm`
19 sheets · 14 tables · 85,163 formulas · VBA

**Job:** a lab operator scans sample barcodes; the workbook produces a submission file for the
external lab (`SIS_Micro_Sub_DATE`) and a summary email body.

**Pipeline:**
```
Samplesin / SamplesinWB / SamplesinCUT     (scanner output lands here)
   → scanMODWB / scanMODCUT                =SamplesinWB!A2              [P1]
   → FILTERWBTODAY / FILTERWB9AM /         (row selection by date/shift)
     FILTERCUTTODAY / FILTERCUT9AM /
     FILTERPREVIOUSDAY / FILTERFRESHKILL
   → MIMS DAILY WB OLD / WB 9AM /          37 × XLOOKUP against
     CUT OLD / CUT 9AM                     'MIMS SAVE FILE'             [P10]
   → FINALEMAIL WB / FINALEMAIL CUT /
     FINALEMAILQ2 WB 9AM                   (presentation)
   → SIS_Micro_Sub_DATE                    (upload file)
```

**The 37-column recipe** (`MIMS DAILY WB OLD`, verbatim):

| Col | Formula | Pattern |
|---|---|---|
| A–F | `=IFERROR(XLOOKUP(FILTERWBTODAY!$B2,'MIMS SAVE FILE'!$M:$M,'MIMS SAVE FILE'!A:A),"")` | P10 |
| G | `=IFERROR(XLOOKUP(FILTERWBTODAY!$B2&FILTERWBTODAY!$F2,'MIMS SAVE FILE'!$M:$M&'MIMS SAVE FILE'!V:V,'MIMS SAVE FILE'!G:G),"")` | **P9** two-key |
| H | key is `F2` alone → different source column `V:V` | P10 |
| L | `=IFERROR(XLOOKUP(K2,'MIMS SAVE FILE'!K:K,'MIMS SAVE FILE'!L:L),"")` | **chained** — keys off its own column K |
| N | `=IF(Table4[[#This Row],[MANUFTIME]]<TIME(HOUR(scanMODWB!$R$1),MINUTE(scanMODWB!$R$1),SECOND(scanMODWB!$R$1)),IF(scanMODWB!$Y$1=2,O2-2,O2-1),TODAY())` | **shift cut-off → back-date 1 or 2 days** |
| O, X | `=TODAY()` | |
| P | `=IFERROR(INDEX(FILTERWBTODAY!Q:Q,SMALL(IF((…B:B=B2)*(…F:F=F2),ROW(…)),COUNTIFS($B$2:B2,B2,$F$2:F2,F2))),"")` | **P15** n-th match |
| Q | `=IFERROR(N2+Table4[[#This Row],[SUPPLIER]],"")` | date + shelf-life days |
| W | `=TEXTAFTER(M2," ",-1)` | **P7** last field |
| Z–AK | `=IF(IFERROR(XLOOKUP(…),"")=0,"")` | zero-blanking variant |

**Why it's the best first scenario:** one key, many columns. Teaches `$` anchoring, `IFERROR`,
fill-right, and the payoff of a well-built lookup table — then escalates to two-key joins and
the shift cut-off logic.

---

### S2 · Complaints Analysis Weekly Database — *the fallback-key join*
`4. KPI Administrator/KPI/Complaints Analysis Weekly Database2.xlsm`
8 sheets · 515,992 formulas · 25 defined names · 21 external links · VBA

**Job:** turn a weekly complaints export + a volumes export into complaints-per-million by
site, retailer and period.

**Pipeline:**
```
Complaints Paste  +  Volumes Paste          (two unrelated exports)
   → Complaints Paste Review                mirror + 3-way fallback key   [P4, P11]
   → Data Extract                            period windows               [P12]
   → Site KPI Reporting                       SUMIFS rollup
   ← List Menu / Period Date Sheet            reference data
```

**Key formulas, verbatim:**
```excel
G1:AW1   =IFERROR(IF('Complaints Paste'!A1="","",'Complaints Paste'!A1),"")     [P4]
A2       =+AC2&""&"-"&""&B2                       build a complaint ref
B2       =IF(C2="",B1-1,C2)                       sequence that fills gaps upward
D2       =IF(AC2=AC1,0,AC2)                       first-of-group marker
C2       =IF(D2=0/1/1900,"",COUNTIF(AC:AC,D2))    ← D3 defect lives here
F2       =IF(AC2="","",TEXT(AC2,"MMM-YY"))        period label
L2       =IF(K2="","",IF(COUNTIF('Volumes Paste'!$A:$A,K2)>=1,1,""))   key-1 exists? [P11]
N2, P2   … same for key-2, key-3
Q2       =IFERROR(IF(L2=1,VLOOKUP(K2,'Volumes Paste'!$A:$G,7,FALSE),
                   IF(N2=1,VLOOKUP(M2,'Volumes Paste'!$B:$G,6,FALSE),
                   IF(P2=1,VLOOKUP(O2,'Volumes Paste'!$C:$G,5,FALSE),
                   "Check"&" "&"-"&" "&O2))),"Check"&" "&"-"&" "&O2)       [P11]
S2       =IF(VLOOKUP(Q2,'Volumes Paste'!$G:$K,5,FALSE)>1,
             "Supplied Within 60 Days","Not Supplied Within 60 Days")
L17      =SUMIFS('Volumes Paste'!17:17,'Volumes Paste'!$13:$13,">=" &K$14,
                                       'Volumes Paste'!$13:$13,"<="&K$15)   [P12]
```

**Why it matters:** this is the hardest *reasoning* in the estate — three code systems, a
fallback cascade, and a human-readable failure message. It also contains D3 and D13, so it
doubles as the debugging scenario.

---

### S3 · E. coli / Campy trending — *the week grid*
`2. Lab Testing Co-ordinator/.../All Lab helpers and speeds and toolkits.xlsm` (47,685 formulas)
and `4. KPI Administrator/KPI/Live E.coli Trending - all sites.xlsm` (667,402 formulas, 19 sheets, 12 external links)

**Job:** micro results per site per week, % over the 1,000 cfu action limit, trended.

```excel
Calendars!A2   =F2&B2                                   composite key  [P9]
Calendars!G2   =+H3+1          H2  =+G2+6               week start/end spine  [P19]
Calendars!E2   =MONTH(G2+2)

'Ecoli Final Data'!A2:S2  ='ECOLI MICRO PASTE'!A5 …     mirror, reordered   [P1]
U2  =IF('ECOLI MICRO PASTE'!AB5<=10,10,'ECOLI MICRO PASTE'!AH5)   limit-of-detection floor
V2  =INDEX(Calendars!F:F, MATCH(1,(C2>=Calendars!G:G)*(C2<=Calendars!H:H),0))  [P16]
W2  =SUMIFS($U:$U,$V:$V,$V2,$AA:$AA,$AA2)                         [P12]
X2  =AVERAGEIFS($U:$U,$V:$V,$V2,$AA:$AA,$AA2)
Y2  =IF(U2>500,1,0)                                      flag                [P20]
Z2  =SUMIFS($Y:$Y,$V:$V,$V2,$AA:$AA,$AA2)/COUNTIFS($V:$V,$V2,$AA:$AA,$AA2)   rate
AB2 =VLOOKUP(V2&AA2,Calendars!$A$1:$H$470,8,0)                    [P9]

Willand!T4 =IF(A4="","-",VLOOKUP(($B4-WEEKDAY($B4,1))+7,Calendars!$H:$J,3,FALSE))  [P8]
'Current month'!T22 =COUNTIFS(B:B,">="&$S$13,B:B,"<="&$S$14,O:O,">=1000")          [P13]
'Current month'!R2  =IF(ROW()<COUNTA($B:$B)-49,"",ROW()-(COUNTA($B:$B)-50))        [P19]
```

**Why it matters:** this is the pure "reorganise, don't calculate" scenario, and the only place
real statistics appear. It also teaches the limit-of-detection floor (`<=10 → 10`), which is
domain knowledge worth encoding.

---

### S4 · M&S Vision LV Verification — *parsing a machine log*
`3. Technical Co-ordinator/M&S Vision/26th October LV Verification - Verified - V20.xlsx`
4 sheets · 1,000,076 formulas (!) · no VBA

**Job:** a vision system emits one text line per pack; turn it into a count of label mismatches
and missing sleeves per product, and flag whether enough packs were verified.

```excel
Sheet4!C1 =IF(RIGHT(LVALL!C1,29)="FLASHLABEL; Invalid: Mismatch","Missing Sleeve",
             IF(RIGHT(LVALL!C1,8)="Mismatch","Mismatch",""))           [P6]
Sheet4!D1 =IF(LEFT(LVALL!C1,3)="JOB",LEFT(LVALL!C1,45),"")             [P6]
Sheet4!E1 =IF(D1="","",RIGHT(D1,8))                                     product code
Sheet4!F1 =IF(E1="","",LVALL!C2)                                        next line = detail  [P3]
Sheet4!G1 =IFERROR(TEXTAFTER(F1," ",3),"")                             [P7]
Sheet4!H1 =IFERROR(VALUE(TEXTBEFORE(G1," ",1)),"")                     [P7]
Sheet4!A1 =IFERROR(IF(E1="",A1_prev,E1),"")                            fill-down-last-seen

Home!B4  =MAX(FOPSALL!I:I)
Home!H4  =COUNTIFS(FOPSALL!$C:$C,E4,FOPSALL!$I:$I,$B$4,FOPSALL!$D:$D,94)*F4
          + COUNTIFS(FOPSALL!$C:$C,C4,FOPSALL!$I:$I,$B$4,FOPSALL!$D:$D,94)*D4
Home!K4  =SUMIFS(Sheet4!H:H,Sheet4!E:E,E4)+SUMIFS(Sheet4!H:H,Sheet4!E:E,C4)
Home!J4  =IF(K4<I4,1,"")                                               shortfall flag
Home!L4  =IF(J4="","",(K4-I4)&" "&"Packs")                             human message  [P17]
Home!M4  =IF(COUNTIFS(Sheet4!$A:$A,E4,Sheet4!$C:$C,"Mismatch")+…=0,"",…)
```

**Why it matters:** the best possible demonstration of old-vs-new text handling (P6 beside P7
in one file), and it shows the "fill-down last seen value" idiom. Also the clearest case of D13
— a million formulas over four sheets.

---

### S5 · Campy WB & Necks labels toolkit — *the many-source assembly*
`2. Lab Testing Co-ordinator/Toolkits and Files/Caeca and Necks/Campy WB and Necks labels toolkit v17.xlsm`
48 sheets (most hidden) · 441,660 formulas · 29 defined names · VBA

**Job:** from four pasted exports (`LBI Paste`, `Kill Data Paste`, `Despatch Paste`,
`Transaction Listings Paste`), assemble printable sample labels with correct batch, use-by and
kill-time data, and report its own data errors.

```
LBI Paste · Kill Data Paste · Despatch Paste · Transaction Listings Paste
   → Despatch Hidden        =IF('Despatch Paste'!A6>0,'Despatch Paste'!A6,"")      [P2]
   → LBI Hidden, Line 1/2 Hidden (×4 each), Workings Hidden
   → Kill Data Information (×17 copies!), Batch Data Information, Kill Time Calculator
   → Label for Micro Neck, Label Info
   → Error Coding
```

```excel
'Batch Data Information'!H13 =IF(K12=L12,"Batch Use By Date:","Multiple Use By Dates")
'Batch Data Information'!E6  =MIN('Transaction Listings Paste'!I:I)
'Batch Data Information'!N28 =IFERROR(INDEX('Transaction Listings Paste'!$A$7:$A$19998,
   MATCH(0,IF(ISBLANK(…$A$7:$A$19998),1,COUNTIF($N$9:$N27,…$A$7:$A$19998)),0)),"-")
                                                    ← unique-list extraction, array
'Kill Data Information'!A27 =IF(MIN(A28:A77)=0,"",MIN(A28:A77))
'Kill Data Information'!D28 =IFERROR(VLOOKUP($C28,'Line 1 Hidden (2)'!$A:$Z,2,FALSE),"")
'Kill Data Information'!O13 =SUMIFS('Workings Hidden'!H:H,'Workings Hidden'!F:F,O14,
                                    'Workings Hidden'!G:G,'Kill Data Information'!J9)
'Kill Data Information'!X28 =IFERROR(IF(N28=$U$17,AND(Q28>=$W$13,O28<=$W$17),"")*1,0)
                                                    ← AND() coerced to 1/0 by *1
'LBI Hidden'!R2 =IF(M2="L1",N2+'Kill Data Information'!$AE$8,
                   IF(M2="L2",O2+'Kill Data Information'!$AE$11,""))   line-specific offset
'Error Coding'!C1 =IF(MIN('Batch Data Information'!D3,'Batch Data Information'!F3,
                        'Kill Data Information'!A27,…)="","",MIN(…))   error aggregation [P17]
```

**Why it matters:** the capstone. Four sources, a hidden workings layer, a unique-list array
formula, line-specific business rules, and a self-audit sheet. Also the clearest example of
D11 (identical copy in two folders) and of sheet-duplication-as-copy-paste
(`Kill Data Information (2)` … `(17)`).

---

## 11. Proposed curriculum ladder

Mapped to the patterns above, ordered by dependency rather than by Excel's own menus.

| Level | Title | Patterns | Engine features needed |
|---|---|---|---|
| **0** | The grid, the formula bar, F2, Enter | — | A1 refs, display vs value |
| **1** | Pointing at another sheet | P1 | cross-sheet refs, quoted names, fill-down |
| **2** | Showing nothing instead of noise | P2, P3, P4 | `IF`, `""` semantics (§7.4) |
| **3** | `$` — the anchor | — | relative/absolute/mixed translation on fill |
| **4** | The join | P5 | `VLOOKUP` exact, whole-column ranges, `#N/A` |
| **5** | Guarding the join | P5, D1, D2 | `IFERROR`, `IFNA`, `#REF!` as a value |
| **6** | The modern join | P10 | `XLOOKUP`, `_xlfn.` prefix |
| **7** | Two keys at once | P9, P14 | `&`, array `&`, expanding ranges |
| **8** | Counting what matters | P13, P11 | `COUNTIF`, `COUNTIFS` ×6 criteria, mixed anchors |
| **9** | Totalling a window | P12, P20 | `SUMIFS`, criteria-as-string, row ranges, `AVERAGEIFS` |
| **10** | Dates, weeks and shifts | P8, P19 | serials, `WEEKDAY`, `TIME/HOUR/MINUTE`, `TEXT` |
| **11** | Cutting up text | P6, P7, D6, D7 | `LEFT/RIGHT/MID/LEN`, `TEXTBEFORE/AFTER`, `VALUE`, `SUBSTITUTE/CHAR` |
| **12** | Banding and routing | P18, D4 | nested `IF` → lookup-table refactor |
| **13** | Making the sheet check itself | P17 | signature columns, verdict strings |
| **14** | Colour as output | §6.2 | conditional formatting incl. `expression` |
| **15** | Dropdowns and picklists | §6.1 | data validation lists, defined names |
| **16** | `INDEX`/`MATCH` and the n-th match | P15, P16 | `INDEX`, `MATCH`, `SMALL`, `ROW`, CSE arrays |
| **17** | Building the five layers | §3 | multi-sheet scenario, graded structure |
| **18** | Debugging someone else's workbook | D1–D14 | trace precedents, error inspector |

Levels 0–9 cover ~85% of the formula mass. Levels 1, 2, 4, 5 alone cover ~70%.

---

## 12. Simulator build requirements

Pulled together as a spec, in priority order.

### Must have (v1)
1. **Multi-sheet workbook** with a sheet tab bar, hidden-sheet support, and `'quoted'!` refs.
2. **Formula engine** per §7: the reference forms in 7.1, operators in 7.2 (including leading
   unary `+`), errors in 7.3, the `""` semantics in 7.4, coercion in 7.5, dates in 7.6, the
   `_xlfn.` tolerance in 7.7.
3. **Function set:** `IF IFERROR IFNA VLOOKUP XLOOKUP INDEX MATCH COUNTIF COUNTIFS SUMIF
   SUMIFS AVERAGEIFS SUM COUNT COUNTA COUNTBLANK MIN MAX AND OR NOT TRUE FALSE LEFT RIGHT MID
   LEN TRIM VALUE TEXT CONCAT & SUBSTITUTE CHAR FIND SEARCH TEXTBEFORE TEXTAFTER TODAY NOW
   DATE WEEKDAY MONTH YEAR DAY HOUR MINUTE SECOND TIME ROW ROWS COLUMN SMALL LARGE RANK
   SUMPRODUCT ISBLANK ISERROR ISNUMBER ISTEXT NA ROUND INT ABS`. That's ~60 functions and it
   covers the estate.
4. **Fill-down / fill-right** with reference translation, as a first-class, mouse-and-keyboard
   interaction. This is the muscle being trained.
5. **Number formats** per §7.8, applied as display only.
6. **Paste** a block of tab-separated text into a sheet — the trainer's defining gesture.
   Ideally with "paste as values" vs "paste over formulas" as distinct, consequential actions.
7. **A grader**: compare the learner's sheet against an expected result *and* against expected
   formula shape, so that hard-coding an answer fails.

### Should have (v2)
8. Conditional formatting (`cellIs`, `expression`, `containsText`, `duplicateValues`).
9. Data validation list dropdowns bound to a range or defined name.
10. Defined-name manager, including showing broken `#REF!` names.
11. CSE array formulas (`Ctrl+Shift+Enter`) for P15/P16.
12. Structured table references `Table[[#This Row],[Col]]`.
13. Trace precedents / dependents, and a formula-evaluation stepper — the debugging level (18)
    needs it.
14. External-reference syntax `[1]Sheet!A1` that parses and reports "link not available", so D1
    and D10 can be taught.

### Won't have
Dynamic arrays, PivotTables, charts beyond a sparkline, VBA, Power Query, Solver. All absent
from the estate (§8); the hidden 15 legacy files may change this verdict only slightly.

### Design principles taken from the data
- **Layers over cleverness.** The grader should reward a readable five-layer build over a single
  monstrous formula, because that is what makes these workbooks survivable.
- **Blank is a decision.** Every lesson from level 2 up should force a choice about what an
  empty source cell produces.
- **Errors are messages to people.** `"Check - 40231"`, `"PLEASE REVIEW AND AMEND"`,
  `"Multiple Use By Dates"` — the estate's own convention. Teach it; don't teach `#N/A`
  suppression as the goal.
- **The data is dirty on purpose.** Sandbox paste-blocks should include trailing spaces,
  `CHAR(160)`, numbers stored as text, zero-padding, and inconsistent date formats. The real
  exports do.

---

## 13. Gaps and next steps

**Gaps in this catalog:**
1. **15 legacy files unscanned** — `.xls` and `.xlsb`. Notably `MIMS_SAVE FILE - v49.xls`
   (the master product reference), `Allergen & Species Risk Assessment Site V48.5.xls`,
   `database_sis_al/ch/mi/sl.xlsb`, `PDH KPI WC *.xlsb`. These are *reference databases* rather
   than formula-heavy toolkits, so the function profile is unlikely to shift much — but
   `MIMS_SAVE FILE` is the join target for 248,121 formulas, so its *column layout* is worth
   documenting for the sandbox fixtures. Requires Excel or LibreOffice to convert.
2. **VBA not read** (71 workbooks). The macros define what the buttons do, which is the boundary
   between "simulator teaches this" and "a macro does this". Worth a separate pass.
3. **Power Query / `M`** — `ExternalData_1` / `ExternalData_2` defined names and
   `queryTables` entries suggest some refreshable queries. Not catalogued.
4. **Column *meanings*** — this catalog documents formula *shapes*, not what `'MIMS SAVE
   FILE'!$M:$M` actually contains. Building realistic sandbox fixtures needs a short data
   dictionary for the half-dozen recurring paste layouts (`MIMS SAVE FILE`, `Micro paste`,
   `LBI Paste`, `Kill Data Paste`, `Despatch Paste`, `Complaints Paste`, `Volumes Paste`,
   `LVALL`).

**Suggested order from here:**
1. Pick the **engine scope** — confirm the §12 "must have" list.
2. Build the **`""`/blank conformance test suite** (§7.4) first. Everything depends on it.
3. Write the **data dictionary** for the 8 paste layouts above (gap 4) — half a day, and it
   unlocks realistic fixtures.
4. Build the engine + grid, validated against the test suite.
5. Author levels 1–5 against scenario **S1** (Barcode Submission), which is the gentlest.
6. Then S4 (text parsing), S3 (week grid), S2 (fallback keys), S5 (capstone).

---

## Appendix A — raw scan artefacts

Produced by the scan; regenerate as needed.

| Artefact | Contents |
|---|---|
| `formulas.jsonl` | 4,266,245 records — every formula, data-validation and conditional-format rule, with file/sheet/cell |
| `meta.jsonl` | 152 records — per-workbook sheets, dimensions, defined names, tables, external links, VBA flag |
| `an/functions.json` | function → (uses, files) |
| `an/shapes.json` | top 400 normalised formula shapes with a real example each |
| `an/crosssheet.json` | referenced-sheet-name frequency |
| `an/dv.json`, `an/cf.json` | validation and conditional-format inventories |
| `an/per_file_func.json`, `an/role_func.json` | function profile per workbook / per role folder |
| `exemplars.txt` | per-column formula recipes for the 7 case-study workbooks |

Scan scripts: `extract.py` (zip/XML streaming extractor), `analyze.py` (aggregator),
`ex.py` (exemplar recipe builder).

All of the above except `formulas.jsonl` (1.0 GB) have been copied to
`D:\Work\Excel Trainer\scan\`. To regenerate the big one:

```
python extract.py formulas.jsonl meta.jsonl     # ~6 min over 152 workbooks
python analyze.py formulas.jsonl meta.jsonl an  # ~4 min
python ex.py > exemplars.txt
```
