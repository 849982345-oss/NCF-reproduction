#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Plot per-epoch HR@10 / NDCG@10 curves from training logs (logs/*.tsv)."""
import glob, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["font.size"] = 13
plt.rcParams["axes.titlesize"] = 15
plt.rcParams["axes.labelsize"] = 13
plt.rcParams["legend.fontsize"] = 12
plt.rcParams["xtick.labelsize"] = 11
plt.rcParams["ytick.labelsize"] = 11

LOGS = os.path.join(os.path.dirname(__file__), "ncf_pytorch", "logs")
OUT = os.path.join(os.path.dirname(__file__), "figures")
os.makedirs(OUT, exist_ok=True)

models = {"GMF": {}, "MLP": {}, "NeuMF-end": {}}
for f in sorted(glob.glob(os.path.join(LOGS, "*.tsv"))):
    base = os.path.basename(f)          # e.g. GMF_ml-1m_factor8.tsv
    model = base.split("_")[0]
    if model not in models:
        continue
    epochs, hrs, ndcgs = [], [], []
    for line in open(f, encoding="utf-8"):
        parts = line.strip().split("\t")
        if len(parts) == 3:
            epochs.append(int(parts[0])); hrs.append(float(parts[1])); ndcgs.append(float(parts[2]))
    models[model] = {"epoch": epochs, "hr": hrs, "ndcg": ndcgs, "file": base}
    print(f"{base}: {len(epochs)} epochs, last HR={hrs[-1]:.4f} NDCG={ndcgs[-1]:.4f}, best HR={max(hrs):.4f}@{epochs[hrs.index(max(hrs))]}")

colors = {"GMF": "#1F4E79", "MLP": "#C9861E", "NeuMF-end": "#2E8B57"}
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12.4, 4.4), dpi=200)

for m, d in models.items():
    if not d:
        continue
    ax1.plot(d["epoch"], d["hr"], label=m, color=colors[m], lw=2.4, marker="o", markersize=2.5, markevery=5)
    ax2.plot(d["epoch"], d["ndcg"], label=m, color=colors[m], lw=2.4, marker="o", markersize=2.5, markevery=5)

for ax in (ax1, ax2):
    ax.grid(alpha=0.25, lw=0.6)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
ax1.set_xlabel("Epoch"); ax1.set_ylabel("HR@10"); ax1.set_title("MovieLens-1M · HR@10 (factor=8)")
ax1.legend(frameon=False, loc="lower right")
ax2.set_xlabel("Epoch"); ax2.set_ylabel("NDCG@10"); ax2.set_title("MovieLens-1M · NDCG@10 (factor=8)")
ax2.legend(frameon=False, loc="lower right")

fig.tight_layout()
p = os.path.join(OUT, "training_curves.png")
fig.savefig(p, dpi=150)
print("saved", p)
