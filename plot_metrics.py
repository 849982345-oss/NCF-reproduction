#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Plot per-epoch HR@10 / NDCG@10 curves from training logs (logs/*.tsv)."""
import glob, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

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

colors = {"GMF": "#1f77b4", "MLP": "#ff7f0e", "NeuMF-end": "#2ca02c"}
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 4.6))

for m, d in models.items():
    if not d:
        continue
    ax1.plot(d["epoch"], d["hr"], label=m, color=colors[m], lw=1.8)
    ax2.plot(d["epoch"], d["ndcg"], label=m, color=colors[m], lw=1.8)

ax1.set_xlabel("Epoch"); ax1.set_ylabel("HR@10"); ax1.set_title("MovieLens-1M · HR@10 (factor=8)")
ax1.grid(alpha=0.3)
ax1.legend()
ax2.set_xlabel("Epoch"); ax2.set_ylabel("NDCG@10"); ax2.set_title("MovieLens-1M · NDCG@10 (factor=8)")
ax2.grid(alpha=0.3)
ax2.legend()

fig.tight_layout()
p = os.path.join(OUT, "training_curves.png")
fig.savefig(p, dpi=150)
print("saved", p)
