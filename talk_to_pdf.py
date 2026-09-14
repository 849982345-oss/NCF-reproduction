#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Convert the talk-script markdown into a clean, tablet-friendly PDF via HTML + headless Edge."""
import os, re, subprocess, html as htmllib

SRC = r"C:\research-training\repro-ncf\讲稿与QA备忘.md"
HTML = r"C:\research-training\repro-ncf\讲稿_打印版.html"
PDF = r"C:\Users\Administrator\Desktop\讲稿与QA备忘.pdf"

md = open(SRC, encoding="utf-8").read()

def inline(s):
    s = htmllib.escape(s)
    s = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"`(.+?)`", r"<code>\1</code>", s)
    return s

out = []
lines = md.split("\n")
i = 0
in_table = False
in_ul = False
while i < len(lines):
    ln = lines[i]
    # tables
    if ln.strip().startswith("|") and i + 1 < len(lines) and set(lines[i+1].replace("|", "").strip()) <= set("-: "):
        rows = []
        while i < len(lines) and lines[i].strip().startswith("|"):
            cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
            if not set("".join(cells)) <= set("-: "):
                rows.append(cells)
            i += 1
        out.append("<table>")
        for ri, r in enumerate(rows):
            tag = "th" if ri == 0 else "td"
            out.append("<tr>" + "".join(f"<{tag}>{inline(c)}</{tag}>" for c in r) + "</tr>")
        out.append("</table>")
        continue
    if ln.startswith("### "):
        out.append(f"<h3>{inline(ln[4:])}</h3>")
    elif ln.startswith("## "):
        out.append(f"<h2>{inline(ln[3:])}</h2>")
    elif ln.startswith("# "):
        out.append(f"<h1>{inline(ln[2:])}</h1>")
    elif ln.strip().startswith("> "):
        out.append(f"<blockquote>{inline(ln.strip()[2:])}</blockquote>")
    elif ln.strip() in ("---", "***"):
        out.append("<hr>")
    elif re.match(r"^\s*[-*] ", ln):
        if not in_ul:
            out.append("<ul>"); in_ul = True
        bullet_text = re.sub(r"^\s*[-*] ", "", ln)
        out.append("<li>" + inline(bullet_text) + "</li>")
    else:
        if in_ul:
            out.append("</ul>"); in_ul = False
        if ln.strip():
            out.append(f"<p>{inline(ln)}</p>")
    i += 1
if in_ul:
    out.append("</ul>")

html_doc = """<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>讲稿与QA备忘</title>
<style>
  @page { size: A4; margin: 14mm 13mm; }
  body { font-family: "Microsoft YaHei", "PingFang SC", sans-serif; font-size: 15px; line-height: 1.75; color: #1f2430; }
  h1 { font-size: 22px; color: #1F4E79; border-bottom: 2px solid #C9A227; padding-bottom: 6px; margin: 0 0 14px; }
  h2 { font-size: 18px; color: #1F4E79; margin: 20px 0 8px; page-break-after: avoid; }
  h3 { font-size: 16px; color: #21242B; margin: 14px 0 6px; page-break-after: avoid; }
  p { margin: 6px 0; }
  ul { margin: 6px 0 6px 18px; padding: 0; }
  li { margin: 3px 0; }
  table { border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 13px; }
  th { background: #1F4E79; color: #fff; padding: 6px 8px; text-align: left; }
  td { border-bottom: 1px solid #dde3ea; padding: 5px 8px; vertical-align: top; }
  tr:nth-child(even) td { background: #F4F7FB; }
  blockquote { border-left: 3px solid #C9A227; background: #FBF7EC; margin: 10px 0; padding: 8px 12px; color: #4b5563; font-size: 13.5px; }
  hr { border: none; border-top: 1px solid #e5e7eb; margin: 18px 0; }
  code { background: #eef2f7; padding: 1px 4px; border-radius: 3px; font-size: 13px; }
  strong { color: #123c63; }
</style></head><body>
""" + "\n".join(out) + "\n</body></html>"

open(HTML, "w", encoding="utf-8").write(html_doc)

edge_candidates = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
edge = next((e for e in edge_candidates if os.path.exists(e)), None)
if not edge:
    raise SystemExit("Edge not found")
cmd = [edge, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
       f"--print-to-pdf={PDF}", "file:///" + HTML.replace("\\", "/")]
r = subprocess.run(cmd, capture_output=True, timeout=180)
print("edge exit:", r.returncode, "| pdf exists:", os.path.exists(PDF), "| size:",
      os.path.getsize(PDF) if os.path.exists(PDF) else 0)
