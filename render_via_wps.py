#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Render a .pptx to PNGs via WPS Office COM -> PDF -> pymupdf."""
import os, subprocess, sys, glob

def pptx_to_pdf(pptx_path, pdf_path=None):
    pptx_path = os.path.abspath(pptx_path)
    if pdf_path is None:
        pdf_path = os.path.splitext(pptx_path)[0] + ".pdf"
    pdf_path = os.path.abspath(pdf_path)
    ps = r"""
$ErrorActionPreference = 'Stop'
try {
  $app = New-Object -ComObject Kwpp.Application
  $pres = $app.Presentations.Open('%s', $true, $false, $false)
  $pres.SaveAs('%s', 32)
  $pres.Close()
  $app.Quit()
  Write-Output 'PDF_OK'
} catch {
  Write-Output ('PDF_FAIL: ' + $_.Exception.Message)
  try { $app.Quit() } catch {}
}
""" % (pptx_path.replace("\\", "\\\\"), pdf_path.replace("\\", "\\\\"))
    r = subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps],
                       capture_output=True, timeout=300)
    def dec(b):
        return (b or b"").decode("gbk", errors="replace")
    out = (dec(r.stdout) + dec(r.stderr)).strip()
    if "PDF_OK" in out and os.path.exists(pdf_path):
        return pdf_path
    raise RuntimeError("WPS export failed: " + out[-500:])

def pdf_to_pngs(pdf_path, outdir):
    import fitz
    os.makedirs(outdir, exist_ok=True)
    doc = fitz.open(pdf_path)
    paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=110)
        p = os.path.join(outdir, f"slide_{i+1:02d}.png")
        pix.save(p)
        paths.append(p)
    return paths

if __name__ == "__main__":
    pptx = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else "render"
    pdf = pptx_to_pdf(pptx)
    pngs = pdf_to_pngs(pdf, outdir)
    print("RENDERED", len(pngs), "slides ->", outdir)
    for p in pngs:
        print(p)
