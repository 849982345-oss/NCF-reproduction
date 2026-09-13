#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Render NCF core formulas as a clean PNG (mathtext, no LaTeX install needed)."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["mathtext.fontset"] = "cm"

FORMULAS = [
    ("GMF · 广义矩阵分解", r"$\hat{y}_{ui} = a_{out}\left(h^{\top}(p_u \odot q_i)\right)$"),
    ("MLP · 多层感知机", r"$z_1 = \phi_1(p_u,\, q_i),\qquad z_l = \phi_l(W_l z_{l-1} + b_l)$"),
    ("NeuMF · 双塔融合", r"$\hat{y}_{ui} = \sigma\left(h^{\top}\left[\,p_u^{G} \odot q_i^{G}\ ;\ z_L\,\right]\right)$"),
    ("训练目标 · 二分类 log loss(1 正 : 4 负)", r"$\mathcal{L} = -\sum_{(u,i)\in\mathcal{Y}^{+}} \log \hat{y}_{ui} \;-\; \sum_{(u,j)\in\mathcal{Y}^{-}} \log\left(1-\hat{y}_{uj}\right)$"),
    ("HR@10 · 命中率", r"$\mathrm{HR@10} = \frac{1}{|U|}\sum_{u\in U} \mathbb{1}\left(\mathrm{rank}_u \leq 10\right)$"),
    ("NDCG@10 · 归一化折损累计增益", r"$\mathrm{NDCG@10} = \frac{1}{|U|}\sum_{u\in U} \frac{\mathbb{1}\left(\mathrm{rank}_u \leq 10\right)}{\log_2(\mathrm{rank}_u + 2)}$"),
]

fig = plt.figure(figsize=(11.6, 6.05), dpi=200)
fig.patch.set_facecolor("white")
n = len(FORMULAS)
for i, (label, formula) in enumerate(FORMULAS):
    y = 0.965 - i * (0.965 - 0.10) / (n - 1)
    fig.text(0.015, y + 0.035, label, fontsize=12.5, color="#1F4E79", weight="bold", va="top")
    fig.text(0.015, y - 0.045, formula, fontsize=18, color="#21242B", va="top")

out = r"C:\research-training\repro-ncf\figures\formulas.png"
fig.savefig(out, dpi=200, facecolor="white")
print("saved", out)
