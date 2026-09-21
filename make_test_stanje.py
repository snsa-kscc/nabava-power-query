#!/usr/bin/env python3
"""Build a synthetic OLDER 'Stanje skladišta' export from the newest real one.

Only for local testing of the stock-movement columns: with a single real export
qStanjePrethodno clamps to the current file and movement is always zero, so
there is nothing to check. NEVER ship the generated file to the client.

The output is a byte-for-byte copy of the real export except for the Zaliha
column, so every other column, style and sheet name stays authentic:
  * ~70% of Zaliha values nudged by -20%..+20% (deterministic, seed=1)
  * 5 rows removed, standing in for articles that did not exist a week earlier

    python3 make_test_stanje.py
"""
import re, random, sys, zipfile
from pathlib import Path

DATA      = Path("data")
OUT       = DATA / "Stanje skladišta 27.08.2026 TEST-SINTETSKI.xlsx"
DROP_ROWS = 5
SEED      = 1

candidates = [p for p in sorted(DATA.glob("Stanje skladišta*.xlsx"))
              if "SINTETSKI" not in p.name]
if not candidates:
    sys.exit("no real Stanje export found in data/")
src = candidates[-1]
print(f"source: {src.name}")

zin   = zipfile.ZipFile(src)
sname = next(n for n in zin.namelist() if re.match(r'xl/worksheets/sheet\d+\.xml$', n))
sheet = zin.read(sname).decode("utf-8")
values = [ "".join(re.findall(r'<t[^>]*>([^<]*)</t>', s))
           for s in re.findall(r'<si>.*?</si>', zin.read("xl/sharedStrings.xml").decode("utf-8"), re.S) ]

rows   = list(re.finditer(r'<row[^>]*r="(\d+)"[^>]*>.*?</row>', sheet, re.S))
header = rows[0]

# locate the Zaliha column from the header row (header cells ARE shared strings)
zal_col = None
for ref, v in re.findall(r'<c r="([A-Z]+)\d+"[^>]*t="s"[^>]*><v>(\d+)</v>', header.group(0)):
    if values[int(v)].strip().lower() == "zaliha":
        zal_col = ref
if not zal_col:
    sys.exit("could not find the Zaliha column in the header row")
print(f"Zaliha column: {zal_col}")

rnd     = random.Random(SEED)
dropped = set(rnd.sample([m.group(1) for m in rows[1:]], DROP_ROWS))

out, cursor, changed, skipped = [], 0, 0, 0
for m in rows:
    out.append(sheet[cursor:m.start()])
    cursor = m.end()
    rnum, body = m.group(1), m.group(0)

    if rnum in dropped:
        continue                                   # row vanishes entirely

    if m is not header and rnd.random() < 0.70:
        # numeric cell: <c r="H2" s="4"><v>1595</v></c>  (no t attribute)
        cell = re.search(rf'(<c r="{zal_col}{rnum}"(?![^>]*t=")[^>]*><v>)([^<]*)(</v>)', body)
        if cell:
            old = cell.group(2)
            try:
                num = float(old)
            except ValueError:
                num = None
            if num is not None:
                new = max(0, int(round(num * (1 + rnd.uniform(-0.20, 0.20)))))
                if new != num:
                    body = body[:cell.start(2)] + str(new) + body[cell.end(2):]
                    changed += 1
        else:
            skipped += 1
    out.append(body)
out.append(sheet[cursor:])

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        zout.writestr(item, "".join(out) if item.filename == sname else zin.read(item.filename))

print(f"wrote {OUT.name}")
print(f"  rows dropped:          {DROP_ROWS} (rows {sorted(dropped, key=int)})")
print(f"  Zaliha values changed: {changed}")
print(f"  rows with no Zaliha cell (left alone): {skipped}")
