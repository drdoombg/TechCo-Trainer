# Open questions

Cases where I am not certain what real Excel returns. Each one is wired into
`tests.html` as an **amber (OPEN)** case: the engine has a behaviour, the suite
exercises it, but it is **not counted as a pass** until confirmed.

To settle one: type the formula into Excel at work, note the answer, and tell me.
I will either confirm the engine or change it and move the case to green.

---

## Q1 — Does unary `+` leave text alone?

```excel
=+"abc"
```

**Engine says:** `abc` — unary plus is a no-op that does not coerce.

**Why it matters:** 34,251 cells in the estate start with a bare `+` (`=+K13+1`,
`=+Home!AS1`, `=+'QTR 2 Old'!#REF!`). If unary `+` coerces to a number, then
`=+A1` over a text cell gives `#VALUE!` and a lot of those cells would be
breaking. I believe it does not coerce, but it is worth one keystroke to be sure.

**Also check:** `=+TRUE` — does it give `TRUE` or `1`?

---

## Q2 — Is `^` left- or right-associative?

```excel
=2^3^2
```

**Engine says:** `64`, i.e. `(2^3)^2`, left-associative.

If Excel gives `512` it is right-associative and `parsePow` in
`js/engine/parser.js` needs to recurse on the right instead of looping.

**Why it matters:** barely. The estate has no `^` at all. It is in the suite
because an engine should not be wrong about something this cheap to check.

---

## Q3 — `=IF(FALSE,1,)` versus `=IF(FALSE,1)`

```excel
=IF(1=2,1,)      ' third argument present but written as nothing
=IF(1=2,1)       ' third argument absent entirely
```

**Engine says:** `0` for the first, `FALSE` for the second.

**This one matters, and it contradicts the catalog.** Catalog defect **D4** says
the trailing-comma form returns `FALSE`. The real formula it was describing is:

```excel
=IF(S3="Aldi",U3,IF(S3="CO-OP",U3,IF(S3="Tesco",V3,IF(S3="M&S","2015/16",
   IF(S3="Sainsbury",V3,IF(S3="Asda",V3,))))))
```

That innermost `,))))` is a *present but empty* argument, so by the engine's
reading an unmatched retailer prints **`0`**, not `FALSE`. Either way it prints
something wrong on a report, so the defect stands — but the lesson text should
say the right thing.

**What to check:** put `=IF(1=2,1,)` and `=IF(1=2,1)` in two cells and read them.
If they differ, D4 in the catalog needs a one-word correction and the level 12
lesson copy follows it.

---

## Q4 — What does `=TEXT(123,)` give?

```excel
=TEXT(C1,)
```

**Engine says:** the value as plain text, i.e. the same as `=C1&""`.

Appears 1,564 times in `NEW FOPS 8 PLU DATABASE v5.xlsm`, always paired with a
`SUBSTITUTE(…,CHAR(160)," ")` on the next column, so the intent was clearly "give
me this as text". But an empty `format_text` could equally be `#VALUE!` or `""`.

**What to check:** `=TEXT(123,)` and `=TEXT(123,"")` — are they the same?

---

## Settled, for the record

These were uncertain while building and are now confirmed by the suite against
documented Excel behaviour. Listed so they are not re-litigated.

| Behaviour | Answer | Where |
|---|---|---|
| `=""` cell: `ISBLANK` | `FALSE` | §7.4 group |
| `=""` cell: `A1=""` | `TRUE` | §7.4 group |
| `=""` cell: `A1>0` | `TRUE` — text beats any number | §7.4 group |
| empty cell: `A1=""` **and** `A1=0` | both `TRUE` | §7.4 group |
| `COUNTBLANK` counts `""` | yes; `COUNTA` also counts it | §7.4 group |
| `"10"=10` | `FALSE` | §7.5 group |
| `COUNTIF(A1,5)` where A1 is text `"5"` | `1` — criteria coerce, `=` does not | §7.5 group |
| `VLOOKUP(5002874, …)` against text keys | `#N/A` — exact match does not coerce | lookups group |
| `1/3&""` | `0.333333333333333` (15 s.f.) | §7.5 group |
| `TRIM` and `VALUE` versus `CHAR(160)` | neither removes it | §7.5 group, defect D7 |
| `0/1/1900` | `0` — it is division | §7.6 group, defect D3 |
| `WEEKDAY(36526,1)` | `7` (2000-01-01 was a Saturday) | §7.6 group |
| circular reference | shows `0`, warns, never hangs | §7.3 group |
| `IFNA` vs `IFERROR` on `#REF!` | `IFNA` lets it through | §7.3 group, defect D2 |
