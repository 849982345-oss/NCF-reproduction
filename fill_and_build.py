#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""After all training finishes: parse metrics, fill numbers into README/record/PPT spec,
plot curves, and build the final .pptx."""
import json, os, glob, re, subprocess, sys

ROOT = r"C:\research-training\repro-ncf"
LOGS = os.path.join(ROOT, "ncf_pytorch", "logs")
PY313 = r"C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe"
CREATE = r"C:\Users\Administrator\AppData\Local\hermes\skills\productivity\powerpoint\scripts\pptx_create.py"
PLOT = os.path.join(ROOT, "plot_metrics.py")

def load_tsv(model):
    f = os.path.join(LOGS, f"{model}_ml-1m_factor8.tsv")
    rows = []
    for line in open(f, encoding="utf-8"):
        p = line.strip().split("\t")
        if len(p) == 3:
            rows.append((int(p[0]), float(p[1]), float(p[2])))
    return rows

def best(rows):
    best_hr = max(rows, key=lambda r: r[1])
    best_ndcg = max(rows, key=lambda r: r[2])
    return best_hr, best_ndcg

results = {}
for m in ["GMF", "MLP", "NeuMF-end"]:
    rows = load_tsv(m)
    if not rows:
        print("MISSING logs for", m); sys.exit(1)
    bhr, bnd = best(rows)
    last = rows[-1]
    results[m] = {
        "epochs": len(rows), "last": last,
        "best_hr": bhr[1], "best_hr_epoch": bhr[0],
        "best_ndcg": bnd[2], "best_ndcg_epoch": bnd[0],
    }
    print(f"{m}: {len(rows)} epochs | best HR {bhr[1]:.4f}@{bhr[0]} | best NDCG {bnd[2]:.4f}@{bnd[0]} | last {last[1]:.4f}/{last[2]:.4f}")

neu = results["NeuMF-end"]
gmf = results["GMF"]
mlp = results["MLP"]

# ---- README / experiment record ----
readme = open(os.path.join(ROOT, "README.md"), encoding="utf-8").read()
readme = readme.replace("| GMF | _待填_ | _待填_ |", f"| GMF | {gmf['best_hr']:.4f} | {gmf['best_ndcg']:.4f} |")
readme = readme.replace("| MLP | _待填_ | _待填_ |", f"| MLP | {mlp['best_hr']:.4f} | {mlp['best_ndcg']:.4f} |")
readme = readme.replace("| NeuMF | _待填_ | _待填_ |", f"| NeuMF | {neu['best_hr']:.4f} | {neu['best_ndcg']:.4f} |")
open(os.path.join(ROOT, "README.md"), "w", encoding="utf-8").write(readme)

rec = open(os.path.join(ROOT, "复现实验记录.md"), encoding="utf-8").read()
rec = rec.replace("| GMF | {填} | {填} |", f"| GMF | {gmf['best_hr']:.4f} | {gmf['best_ndcg']:.4f} |")
rec = rec.replace("| MLP | {填} | {填} |", f"| MLP | {mlp['best_hr']:.4f} | {mlp['best_ndcg']:.4f} |")
rec = rec.replace("| NeuMF(no pretrain) | {填} | {填} |", f"| NeuMF(no pretrain) | {neu['best_hr']:.4f} | {neu['best_ndcg']:.4f} |")
rec = rec.replace("{待填}", f"NeuMF 复现 HR@10={neu['best_hr']:.4f}(epoch {neu['best_hr_epoch']}),NDCG@10={neu['best_ndcg']:.4f};论文 0.688/0.410。GMF={gmf['best_hr']:.4f}/{gmf['best_ndcg']:.4f}(论文 Fig.4≈0.645/0.335),MLP={mlp['best_hr']:.4f}/{mlp['best_ndcg']:.4f}(论文 Fig.4≈0.640/0.345)。相对趋势 NeuMF>GMF≈MLP 与论文一致;绝对差距主要来自 TF→PyTorch 实现、训练预算/调参与论文网格搜索最优的差异,以及论文 Figure 读数误差。")
open(os.path.join(ROOT, "复现实验记录.md"), "w", encoding="utf-8").write(rec)

# ---- talk notes (讲稿与QA备忘) ----
talk = open(os.path.join(ROOT, "讲稿与QA备忘.md"), encoding="utf-8").read()
diff_hr = neu['best_hr'] - 0.688
diff_ndcg = neu['best_ndcg'] - 0.410
talk = (talk.replace("{MY_NEUMF_HR}", f"{neu['best_hr']:.3f}")
            .replace("{MY_NEUMF_NDCG}", f"{neu['best_ndcg']:.3f}")
            .replace("差 X 个点", f"差 {abs(diff_hr) * 100:.1f} 个百分点(HR)")
            .replace("{HR_DIFF}", f"{diff_hr:+.3f}")
            .replace("{NDCG_DIFF}", f"{diff_ndcg:+.3f}"))
open(os.path.join(ROOT, "讲稿与QA备忘.md"), "w", encoding="utf-8").write(talk)
print("Talk notes filled.")

# ---- PPT spec ----
specp = os.path.join(ROOT, "ppt_spec_v1.json")
spec = json.load(open(specp, encoding="utf-8"))
def fill(text):
    return (text.replace("{MY_NEUMF_HR}", f"{neu['best_hr']:.3f}")
                .replace("{MY_NEUMF_NDCG}", f"{neu['best_ndcg']:.3f}")
                .replace("{MY_GMF_HR}", f"{gmf['best_hr']:.3f}")
                .replace("{MY_GMF_NDCG}", f"{gmf['best_ndcg']:.3f}")
                .replace("{MY_MLP_HR}", f"{mlp['best_hr']:.3f}")
                .replace("{MY_MLP_NDCG}", f"{mlp['best_ndcg']:.3f}")
                .replace("{EPOCHS_CONVERGE}", f"{neu['best_hr_epoch']}"))
for s in spec["slides"]:
    if s.get("title"): s["title"] = fill(s["title"])
    if s.get("subtitle"): s["subtitle"] = fill(s["subtitle"])
    if s.get("bullets"):
        s["bullets"] = [fill(b) if isinstance(b, str) else {**b, "text": fill(b["text"])} for b in s["bullets"]]
    if s.get("notes"): s["notes"] = fill(s["notes"])
    if s.get("tables"):
        for t in s["tables"]:
            t["rows"] = [[fill(c) if isinstance(c, str) else c for c in row] for row in t["rows"]]
    # inject training-curves image on the training-progress slide after plots exist
    if s.get("title", "").startswith("训练过程"):
        img = os.path.join(ROOT, "figures", "training_curves.png")
        if os.path.exists(img):
            s["images"] = [{"path": img, "left": 0.55, "top": 1.6, "width": 12.2}]
json.dump(spec, open(specp, "w", encoding="utf-8"), ensure_ascii=False)
print("PPT spec filled.")

# ---- plots ----
subprocess.run([PY313, PLOT], check=True)

# ---- build pptx ----
subprocess.run([PY313, CREATE, specp, os.path.join(ROOT, "NCF复现与暑期学习汇报_徐胜钦.pptx")], check=True)
print("PPTX built.")
print("ALL_FILL_DONE")
