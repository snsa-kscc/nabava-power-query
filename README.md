# NABAVA — Power Query setup

Internal notes. Everything here runs on your Excel machine once.

---

## Why the queries aren't already inside the file

Power Query code lives in a binary `DataMashup` part of the `.xlsx`, which
openpyxl cannot write. So `NABAVA_PQ.xlsx` ships as a wired skeleton — settings,
named ranges, the manual table, the conditional formatting — and the queries are
added once, on a machine that has Excel.

`import_queries.ps1` does that through Excel's own object model in a few seconds.
Pasting by hand is the fallback and takes about ten minutes. Either way, once the
queries are in, the file is self-contained and the end user never touches any of it.

---

## Setup, once

1. Put `NABAVA_PQ.xlsx` in the folder with the ERP exports, and **save it once**
   (`Ctrl+S`). Until it is saved, `CELL("filename")` returns empty and
   `POSTAVKE!B4` cannot resolve the folder. The query raises a clear Croatian
   error if you forget.

2. Add the queries — **either** by script **or** by hand.

   ### By script (Windows, Excel installed)

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\import_queries.ps1
   ```

   It reads the `/*@ query: … */` markers in `NABAVA_QUERIES.m`, adds all ten
   queries in order, loads `qNabava` to `NABAVA!$A$4`, refreshes and saves. It is
   re-runnable: queries of the same name are dropped first, so after every edit to
   the `.m` file just run it again. `qStanjePrethodno` is generated from `qStanje`
   automatically, so there is no "change the marked line" step to forget.

   Before running it, on any machine, with no Excel needed:

   ```bash
   python3 check_queries.py
   ```

   which parses the `.m` exactly as the importer does and checks block boundaries,
   brace balance, the derive anchor, and that no query references one created later.

   ### By hand

   For each block in `NABAVA_QUERIES.m`, in this order:
   `Podaci → Dohvati podatke → Iz drugih izvora → Prazan upit`
   → `Prikaz → Napredni uređivač` → paste → `Gotovo` → rename the query.

   Order matters, each one references the previous:

   | # | Query name | Load as |
   |---|---|---|
   | 1 | `fnNorm` | Samo veza |
   | 2 | `fnStupac` | Samo veza |
   | 3 | `Putanja` | Samo veza |
   | 4 | `Postavke` | Samo veza |
   | 5 | `fnDatoteke` | Samo veza |
   | 6 | `qStanje` | Samo veza |
   | 7 | `qStanjePrethodno` | Samo veza — same code as `qStanje`, change `idx = 0` to `idx = 1` on the marked line (the importer does this for you) |
   | 8 | `qProdaja` | Samo veza |
   | 9 | `qDolasci` | Samo veza |
   | 10 | `qNabava` | **Load to sheet NABAVA, cell A4** |

3. Load `qNabava` with `Zatvori i učitaj u… → Tablica → Postojeći radni list →
   `=NABAVA!$A$4`. The conditional formatting is already bound to `Q5:Q4000`
   (STATUS) and `I5:I4000` (Zaliha traje), so it lights up on first load.

4. **Allow the privacy firewall.** `Putanja` reads the folder out of a cell and
   `fnDatoteke` then reads the disk, and Power Query refuses to combine a query
   reference with a direct data-source hit:

   > Query 'fnDatoteke' (step 'fnDatoteke') references other queries or steps, so
   > it may not directly access a data source. Please rebuild this data combination.

   `Podaci → Dohvati podatke → Opcije upita → Privatnost`
   (`Data → Get Data → Query Options → Privacy`) → **ignore privacy levels**.

   Set it under **Current Workbook**, not only Global: the workbook-scoped setting
   is stored in the file, so the end user gets it without touching any options. The
   Global setting only covers the machine it was set on.

5. `Ctrl+Alt+F5` to refresh everything. Check the row count matches ~930.

---

## Language independence

The end user's Excel language was never established, so nothing depends on it.

**M is identical in every language version.** Function names in M (`Table.Group`,
`Number.From`) never localise — unlike worksheet functions, which do. This is the
main reason the logic lives in Power Query rather than in cells.

**Column names are matched after normalising.** `fnNorm` strips Croatian
diacritics and lowercases, `fnStupac` looks up a column through it. So
`Dobavljač`, `DOBAVLJAC` and `dobavljac` all resolve to the same column, and a
column that disappears returns `null` instead of throwing.

**Every conversion carries an explicit culture.** `Number.From(_, "hr-HR")` and
`Table.TransformColumnTypes(…, "hr-HR")`. Without this, a machine set to en-US
reads `1.234` as one-point-two-three-four instead of one thousand two hundred
thirty four. This is the classic silent corruption on mixed-locale machines.

**Sheets are taken by position, never by name** — `listovi{0}[Data]` — because the
export sheet name carries the export date.

**Worksheet formulas are stored in English inside the file** and rendered in the
UI language, so `=LEFT(CELL("filename")…)` in `POSTAVKE!B4` displays as
`=LIJEVO(ĆELIJA(…))` on a Croatian install and works identically.

---

## Minimum Excel version

**Microsoft 365 or Excel 2021.**

The binding constraint is the `??` null-coalescing operator, used in `fnNorm`,
`qStanje` and `qDolasci`. It entered the Power Query engine around mid-2020, so
it ships in Excel 2021 and Microsoft 365 but **not in Excel 2019 or 2016**.
`Text.Select` in `fnDatoteke` is the softer one — added around 2017, so it is
fine on 2019 but not on 2016.

Everything else used predates all of these.

The worksheet formulas are not the limiting factor. `NABAVA_model.xlsx` uses only
`SUMIFS`, `COUNTIFS`, `OFFSET`, `INDEX`, `MATCH`, `IFERROR`, `TEXT`, `ROUND`,
`SUM`, `IF`, `AND`, and `NABAVA_PQ.xlsx` only `CELL`, `FIND`, `LEFT`. All of that
runs on 2016 and later.

If the end user turns out to be below 2021, two small edits drop the requirement to
2016: replace each `x ?? 0` with `try x otherwise 0`, and rebuild `Text.Select`
as `Text.Combine(List.Select(Text.ToList(...), each List.Contains({"0".."9","_","."}, _)))`.
Worth also wrapping `Table.Buffer` around `qDolasci` on an older engine, since it
is scanned once per article.

---

## How the folder drives the data

The end user drops files into one folder. Nothing else.

**`Stanje skladišta`** — `fnDatoteke("stanje")` finds every file whose normalised
name starts with `stanje`, reads the export date out of the filename
(`DD_MM_YYYY`, falling back to the file timestamp), and sorts newest first.
`qStanje` takes `idx = 0`, `qStanjePrethodno` the same query with `idx = 1`. The
index is clamped to the number of files found, so a folder with a single export
falls back to the current one and movement reads zero instead of erroring. So:

- drop in a newer export → it becomes "today", yesterday's becomes "previous",
  and the `Izlaz` column shows the movement between them
- export daily → daily movement; keep the 1st-of-month file and one recent one →
  movement since the start of the month
- the model never needs to know which files exist; it re-decides on every refresh

**`Analiza prodaje`** — all of them are read, not just the newest. Each is tagged
with its year, the newest file per year wins, months are unpivoted into one row
each and indexed as `year*12 + month`. The average covers the last `RollingN`
(12) **completed** months, ending before the newest export's month.

That is the year-rollover answer. In January 2027 the window is Feb–Dec 2026 plus
whatever 2027 has completed, so the average never divides by one month. The only
manual act is leaving the 2026 file in the folder.

The current partial month is always excluded — the 03.09. export had 3 days in
Rujan, which would have dragged every average down.

**`ROBA_U_DOLASKU`** — the in-workbook table `tblDolasci`, pre-filled with the 70
open lines from the five ND orders. The end user edits it directly. `qDolasci` drops
empty rows, ranks shipments chronologically per article and keeps the first three.

**Order files are deliberately NOT read.** Arrival dates are maintained by hand in
`ROBA_U_DOLASKU`; reading the order folder automatically is out of scope here.

**The workbook excludes itself** — `fnDatoteke` skips anything starting with
`nabava` and any `~$` lock files, so it never tries to read itself as an export.

**Subfolders work.** `Folder.Files` recurses, and the filter only looks at the
file name, never its path — so keeping the exports in a `data\` subfolder beside
the workbook behaves exactly like keeping them loose. The one rule is that each
export exists in **exactly one place**: two copies of the same `Stanje` file in
different folders both parse to the same export date, so `qStanjePrethodno` would
silently compare the current stock against a copy of itself and report no
movement at all. `Analiza prodaje` is immune — it keeps only the newest file per
year — but `Stanje` is not.

---

## Failure modes, and what happens

| Situation | Behaviour |
|---|---|
| Workbook never saved | Clear Croatian error: "Spremite datoteku (Save) prije osvjezavanja." |
| Privacy levels not ignored | `Formula.Firewall`: "references other queries or steps, so it may not directly access a data source". Expected, not a bug — see step 4. It is the price of resolving the folder from a cell instead of hard-coding a path. |
| Only one `Stanje` file in folder | No error. The index clamps to the only file, so `qStanjePrethodno` equals `qStanje` and `Izlaz` shows zero movement until a second export arrives. |
| No `Stanje` / `Analiza` file at all | Named error rather than a cryptic one |
| Column renamed in a future ERP version | `fnStupac` returns `null`, that column comes through empty, the rest still works |
| Folder moved or renamed | Self-resolves, nothing to edit |
| Article only in sales, not in stock | Included, stock 0 — the 139-article case |
| Two copies of the same `Stanje` export (e.g. one loose, one in `data\`) | **Silent.** Both parse to the same date, so movement reads zero everywhere. Keep each export in one place only. |
| Same article twice in ROBA_U_DOLASKU, same month | Both kept, ranked 1 and 2, never merged |
| More than three shipments per article | Only the three earliest are shown |

---

## Verified before shipping

The status logic in `qNabava` was tested against all five of the agreed scenarios
plus three-container edge cases:

| Case | Result |
|---|---|
| 2000 stock, no order | Sve u redu |
| 1000 stock, no order | Naruci odmah |
| 1000 + 1000 @ 4mj | Rupa |
| 1500 + 300 @ 4mj | Kolicina nije dovoljna |
| 2000 + 1500 @ 4mj | Sve u redu |
| 900 + 300@2 + 300@4 + 900@9 | Rupa nakon 2. dolaska |
| 900 + 150@2 + 900@6 | Rupa nakon 1. dolaska |
| 900 + 900@2 + 900@5 + 900@8 | Kolicina nije dovoljna (break-even supply, 4.0 mj cover) |
| avg 0 | Nema prodaje |

---

## Files in this folder

| File | Runs on | What it is |
|---|---|---|
| `NABAVA_PQ.xlsx` | Excel | The deliverable. Skeleton until the queries are imported. |
| `NABAVA_QUERIES.m` | — | The source of truth for all M code. Marker comments make it machine-readable; it is still paste-able by hand. |
| `import_queries.ps1` | Windows + Excel | Imports every query into the workbook. Re-runnable. |
| `check_queries.py` | anywhere | Validates the `.m` without Excel. Run before importing. |
| `make_test_stanje.py` | anywhere | Generates a synthetic older `Stanje` export for local testing. |
| `data/` | — | The ERP exports. |

### Testing stock movement locally

With a single real `Stanje` export, movement is structurally zero and the `Izlaz`
column proves nothing. `make_test_stanje.py` copies the newest real export, nudges
~70% of `Zaliha` values by ±20% and drops 5 rows (standing in for articles that
did not exist a week earlier), and writes it as a dated older export:

```bash
python3 make_test_stanje.py
```

It is deterministic (`seed=1`), so the same numbers come back every run. The
generated file is named `… TEST-SINTETSKI.xlsx`.

> **Delete it before shipping.** It is fabricated stock data. The suffix contains
> no digits, so it does not disturb the date parsed from the filename — which also
> means nothing about the filename stops Excel from refreshing against it.

---

## After it works

Delete `PRIMJER_*.xlsx` and `tablica.xlsx` from the folder if they're still
there — they match no prefix so they're ignored, but they confuse the folder.

The formula version (`NABAVA_model.xlsx`) stays useful as a reference: same
numbers, no queries, opens anywhere. Worth keeping a copy.
