#!/usr/bin/env python3
"""Render UPUTE_NABAVA.md to a print-ready PDF.

Deliberately small: it handles only the Markdown this document uses (headings,
tables, blockquotes, lists, bold, code spans, rules) rather than pulling in a
full Markdown implementation. Croatian diacritics are the reason fonts are
pinned to Liberation - it covers Latin Extended-A, which many UI fonts do not.

    pip install weasyprint
    python3 make_pdf.py            # -> UPUTE_NABAVA.pdf
"""
import html, re, sys
from pathlib import Path

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "UPUTE_NABAVA.md")
OUT = SRC.with_suffix(".pdf")

def inline(t):
    t = html.escape(t, quote=False)
    t = re.sub(r'`([^`]+)`', r'<code>\1</code>', t)
    t = re.sub(r'\*\*([^*]+)\*\*', r'<strong>\1</strong>', t)
    return t

def cells(row):
    return [c.strip() for c in row.strip().strip("|").split("|")]

def convert(md):
    out, lines, i = [], md.split("\n"), 0
    while i < len(lines):
        ln = lines[i]

        if not ln.strip():
            i += 1; continue

        if re.match(r'^---+\s*$', ln):
            out.append("<hr>"); i += 1; continue

        if m := re.match(r'^(#{1,3})\s+(.*)$', ln):
            lvl = len(m.group(1))
            out.append(f"<h{lvl}>{inline(m.group(2))}</h{lvl}>"); i += 1; continue

        # table: header row followed by a |---| separator
        if ln.lstrip().startswith("|") and i + 1 < len(lines) and re.match(r'^\s*\|[\s:|-]+\|\s*$', lines[i + 1]):
            head = cells(ln)
            i += 2
            body = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                body.append(cells(lines[i])); i += 1
            out.append("<table><thead><tr>"
                       + "".join(f"<th>{inline(c)}</th>" for c in head)
                       + "</tr></thead><tbody>")
            for r in body:
                r += [""] * (len(head) - len(r))
                out.append("<tr>" + "".join(f"<td>{inline(c)}</td>" for c in r[:len(head)]) + "</tr>")
            out.append("</tbody></table>")
            continue

        if ln.startswith(">"):
            buf = []
            while i < len(lines) and lines[i].startswith(">"):
                buf.append(lines[i].lstrip(">").strip()); i += 1
            out.append("<blockquote>"
                       + "".join(f"<p>{inline(p)}</p>" for p in " \n".join(buf).split("\n") if p.strip())
                       + "</blockquote>")
            continue

        if re.match(r'^\s*[-*]\s+', ln) or re.match(r'^\s*\d+\.\s+', ln):
            ordered = bool(re.match(r'^\s*\d+\.\s+', ln))
            tag = "ol" if ordered else "ul"
            items = []
            while i < len(lines) and (re.match(r'^\s*[-*]\s+', lines[i]) or re.match(r'^\s*\d+\.\s+', lines[i])):
                items.append(re.sub(r'^\s*(?:[-*]|\d+\.)\s+', '', lines[i])); i += 1
                # continuation lines of the same item
                while i < len(lines) and lines[i].startswith("  ") and lines[i].strip() \
                        and not re.match(r'^\s*(?:[-*]|\d+\.)\s+', lines[i]):
                    items[-1] += " " + lines[i].strip(); i += 1
            out.append(f"<{tag}>" + "".join(f"<li>{inline(x)}</li>" for x in items) + f"</{tag}>")
            continue

        para = []
        while i < len(lines) and lines[i].strip() and not re.match(r'^(#{1,3}\s|>|\||---+\s*$|\s*[-*]\s|\s*\d+\.\s)', lines[i]):
            para.append(lines[i].strip()); i += 1
        if para:
            out.append(f"<p>{inline(' '.join(para))}</p>")
    return "\n".join(out)

CSS = """
@page {
  size: A4; margin: 20mm 17mm 18mm;
  @bottom-left  { content: "NABAVA - upute za rad"; font: 8pt "Liberation Sans"; color: #777; }
  @bottom-right { content: "str. " counter(page) " / " counter(pages); font: 8pt "Liberation Sans"; color: #777; }
}
body { font: 10.5pt/1.45 "Liberation Serif", serif; color: #111; hyphens: none; }
h1 { font: bold 22pt "Liberation Sans", sans-serif; margin: 0 0 4mm; padding-bottom: 3mm;
     border-bottom: 2pt solid #222; }
h2 { font: bold 13pt "Liberation Sans", sans-serif; margin: 6mm 0 2.5mm; break-after: avoid;
     page-break-after: avoid; }
h3 { font: bold 11pt "Liberation Sans", sans-serif; margin: 4.5mm 0 1.8mm; break-after: avoid;
     page-break-after: avoid; }
p { margin: 0 0 2.6mm; }
ul, ol { margin: 0 0 3mm; padding-left: 7mm; }
li { margin-bottom: 1.2mm; }
hr { border: none; border-top: 0.5pt solid #ccc; margin: 4.5mm 0; }
/* tight horizontal padding: 1mm left a visible gap before a following
   comma or full stop, which reads as a typo in running text */
code { font: 9.5pt "Liberation Mono", monospace; background: #f2f2f2;
       padding: 0.2mm 0.4mm; border-radius: 0.8mm; }
strong { font-weight: bold; }
table { width: 100%; border-collapse: collapse; margin: 0 0 4mm; font-size: 9.5pt;
        break-inside: auto; }
thead { display: table-header-group; }
th { text-align: left; background: #ececec; border: 0.5pt solid #bbb;
     padding: 1.6mm 2mm; font-family: "Liberation Sans", sans-serif; font-size: 9pt; }
td { border: 0.5pt solid #ccc; padding: 1.6mm 2mm; vertical-align: top; }
tr { break-inside: avoid; page-break-inside: avoid; }
blockquote { margin: 0 0 4mm; padding: 2.5mm 4mm; background: #f7f7f7;
             border-left: 2.5pt solid #888; break-inside: avoid; }
blockquote p { margin: 0 0 1.5mm; }
blockquote p:last-child { margin-bottom: 0; }
"""

md = SRC.read_text(encoding="utf-8")
doc = f"""<!DOCTYPE html><html lang="hr"><head><meta charset="utf-8">
<title>NABAVA - upute za rad</title><style>{CSS}</style></head>
<body>{convert(md)}</body></html>"""

tmp = SRC.with_suffix(".render.html")
tmp.write_text(doc, encoding="utf-8")

from weasyprint import HTML
HTML(string=doc, base_url=str(SRC.parent.resolve())).write_pdf(OUT)
tmp.unlink()
print(f"wrote {OUT} ({OUT.stat().st_size:,} bytes)")
