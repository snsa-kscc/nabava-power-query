#!/usr/bin/env python3
"""Validate NABAVA_QUERIES.m (and import_queries.ps1) without Excel.

Mirrors the parsing in import_queries.ps1, so a green run here means the
Windows import will see the same 10 queries. Checks marker syntax, block
boundaries, the derive anchor, brace/paren balance, and that every query a
block references is defined before it.
"""
import re, sys

SRC = "NABAVA_QUERIES.m"
text = open(SRC, encoding="utf-8").read()

marker = re.compile(r'^/\*@\s*query:\s*(?P<spec>[^*]+?)\s*\*/\s*$', re.M)
ms = list(marker.finditer(text))
if not ms:
    sys.exit("no markers found")

defs, errors = [], []
for i, m in enumerate(ms):
    end = ms[i + 1].start() if i + 1 < len(ms) else len(text)
    body = text[m.end():end]
    body = re.sub(r'(\s*/\*(?:(?!\*/).)*\*/\s*)+$', '', body, flags=re.S).strip()

    parts = [p.strip() for p in m.group('spec').split('|')]
    d = {"name": parts[0], "load": "connection", "derive": None,
         "find": None, "repl": None, "body": body, "code": ""}
    for p in parts[1:]:
        if p == "load: connection":
            pass
        elif mm := re.match(r'load:\s*sheet\s+(.+?)!(\$?\w+\$?\d+)$', p):
            d["load"] = f"sheet {mm.group(1)}!{mm.group(2)}"
        elif mm := re.match(r'derive-from:\s*(.+)$', p):
            d["derive"] = mm.group(1)
        elif mm := re.match(r'replace-once:\s*(.+?)\s*=>\s*(.+)$', p):
            d["find"], d["repl"] = mm.group(1), mm.group(2)
        else:
            errors.append(f"{d['name']}: unrecognised directive {p!r}")
    defs.append(d)

by_name = {d["name"]: d for d in defs}

for d in defs:
    if d["derive"]:
        src = by_name.get(d["derive"])
        if not src:
            errors.append(f"{d['name']}: derive-from {d['derive']!r} undefined"); continue
        n = src["body"].count(d["find"])
        if n != 1:
            errors.append(f"{d['name']}: anchor {d['find']!r} matched {n}x in {src['name']}, expected 1")
            continue
        d["body"] = src["body"].replace(d["find"], d["repl"])

    b = d["body"]
    if not b:
        errors.append(f"{d['name']}: empty body"); continue
    if not b.startswith("let"):
        errors.append(f"{d['name']}: body does not start with 'let' (starts {b[:20]!r})")

    # comments and string literals removed: used for balance and reference checks
    code = re.sub(r'//[^\n]*', '', re.sub(r'/\*.*?\*/', '', b, flags=re.S))
    code = re.sub(r'"(?:[^"]|"")*"', '""', code)
    d["code"] = code
    for open_c, close_c in (("(", ")"), ("[", "]"), ("{", "}")):
        if code.count(open_c) != code.count(close_c):
            errors.append(f"{d['name']}: unbalanced {open_c}{close_c} "
                          f"({code.count(open_c)} vs {code.count(close_c)})")

# reference order: a query may only call ones created before it
seen = set()
for d in defs:
    for other in by_name:
        if other == d["name"] or other in seen:
            continue
        if re.search(rf'(?<![\w.]){re.escape(other)}(?![\w])', d["code"]):
            errors.append(f"{d['name']}: references {other}, which is created later")
    seen.add(d["name"])

# ---------------------------------------------------------------------------
# M has no compiler here, so a misremembered library name (Number.Min, which
# does not exist - it is List.Min) only surfaces as a refresh error in Excel.
# Every Foo.Bar identifier must be one known to exist in the M standard library.
KNOWN_M = {
    "Date.From", "Date.Month", "Date.Year",
    "Excel.CurrentWorkbook", "Excel.Workbook", "Folder.Files",
    "JoinKind.LeftOuter", "MissingField.Ignore",
    "List.Accumulate", "List.Contains", "List.Count", "List.Distinct",
    "List.First", "List.Max", "List.Min", "List.PositionOf", "List.RemoveNulls",
    "List.Select", "List.Sum", "List.Transform",
    "Number.Abs", "Number.From", "Number.FromText", "Number.Round",
    "Order.Ascending", "Order.Descending",
    "Table.AddColumn", "Table.AddIndexColumn", "Table.Buffer", "Table.Column",
    "Table.ColumnNames", "Table.Combine", "Table.Distinct",
    "Table.ExpandRecordColumn", "Table.FirstN", "Table.Group",
    "Table.NestedJoin", "Table.PromoteHeaders", "Table.RemoveColumns",
    "Table.RenameColumns", "Table.RowCount", "Table.SelectColumns",
    "Table.SelectRows", "Table.Sort", "Table.TransformColumns",
    "Table.TransformColumnTypes", "Table.TransformRows",
    "Table.UnpivotOtherColumns",
    "Text.Combine", "Text.EndsWith", "Text.From", "Text.Length", "Text.Lower",
    "Text.Replace", "Text.Select", "Text.Split", "Text.StartsWith",
    "Text.ToList", "Text.Trim",
}
for d in defs:
    for name in sorted(set(re.findall(r'\b[A-Z][A-Za-z]+\.[A-Za-z]+\b', d["code"]))):
        if name not in KNOWN_M:
            errors.append(f"{d['name']}: {name} is not a known M function "
                          f"- check the name before shipping (add it to KNOWN_M if real)")

# ---------------------------------------------------------------------------
# The importer cannot be executed here, so check the one thing that silently
# breaks it: PowerShell 5.1 reads a BOM-less file as cp1252, and the third byte
# of a UTF-8 em dash decodes to a typographic quote, which the parser accepts as
# a string terminator. Every brace after it then unbalances.
PS1 = "import_queries.ps1"
try:
    data = open(PS1, "rb").read()
except FileNotFoundError:
    data = None

if data is not None:
    body = data[3:] if data[:3] == b"\xef\xbb\xbf" else data
    txt = body.decode("utf-8")
    for lineno, line in enumerate(txt.splitlines(), 1):
        for c in line:
            if ord(c) > 127:
                errors.append(f"{PS1}:{lineno}: non-ASCII {c!r} (U+{ord(c):04X}) "
                              f"- PowerShell 5.1 may read it as a string terminator")
                break
    if data[:3] != b"\xef\xbb\xbf":
        errors.append(f"{PS1}: missing UTF-8 BOM - PowerShell 5.1 will decode it as cp1252")
    for o, c in (("{", "}"), ("(", ")")):
        if txt.count(o) != txt.count(c):
            errors.append(f"{PS1}: unbalanced {o}{c} ({txt.count(o)} vs {txt.count(c)})")

for d in defs:
    print(f"  {d['name']:<18} {d['load']:<22} {len(d['body'].splitlines()):>3} lines"
          + (f"  (derived from {d['derive']})" if d["derive"] else ""))
print(f"\n{len(defs)} queries parsed")

if errors:
    print("\nFAIL")
    for e in errors:
        print("  -", e)
    sys.exit(1)
print("OK - parse clean; import_queries.ps1 is ASCII + BOM")
