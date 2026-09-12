#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Explain HR@10 / NDCG@10 with real data: load trained NeuMF, show per-user example."""
import os, sys, math
import numpy as np
import torch

ROOT = r"C:\research-training\repro-ncf\ncf_pytorch"
sys.path.insert(0, ROOT)
os.chdir(ROOT)

import config
import data_utils

# ---- load data ----
train_data, test_data, user_num, item_num, train_mat = data_utils.load_all()
print(f"users={user_num}, items={item_num}, train interactions={len(train_data)}, test candidates={len(test_data)}")

# ---- load trained NeuMF model ----
model = torch.load(os.path.join("models", "NeuMF-end.pth"), weights_only=False)
model.eval()
device = "cuda" if torch.cuda.is_available() else "cpu"
model = model.to(device)

# ---- evaluate per user (same protocol as evaluate.py) ----
from torch.utils.data import DataLoader
test_dataset = data_utils.NCFData(test_data, item_num, train_mat, 0, False)
test_loader = DataLoader(test_dataset, batch_size=100, shuffle=False, num_workers=0)

ranks = []          # rank of the true item (0-based)
examples = []
with torch.no_grad():
    for batch_idx, (user, item, label) in enumerate(test_loader):
        scores = model(user.to(device), item.to(device)).cpu()   # 100 scores for this user
        order = torch.argsort(scores, descending=True)   # candidate indices by score
        gt_item = item[0].item()                         # first candidate = true next item
        cand_items = item.numpy()
        pos = int((order == 0).nonzero(as_tuple=True)[0].item())  # where candidate#0 (true) landed
        ranks.append(pos)
        if batch_idx in (0, 1, 2) or pos >= 10:
            top3 = [int(cand_items[i]) for i in order[:3].tolist()]
            examples.append((batch_idx, gt_item, pos, top3, float(scores[0]), float(scores[order[0]])))

ranks = np.array(ranks)
hr10 = float((ranks < 10).mean())
ndcg10 = float(np.mean([1.0 / math.log2(r + 2) if r < 10 else 0.0 for r in ranks]))
print(f"\n=== Reproduced metrics for NeuMF (should match the training log) ===")
print(f"HR@10   = {hr10:.4f}")
print(f"NDCG@10 = {ndcg10:.4f}")
print(f"\nrank distribution: top1={int((ranks==0).sum())}, top3={int((ranks<3).sum())}, top10={int((ranks<10).sum())}, beyond10={int((ranks>=10).sum())}, worst={int(ranks.max())}")

print("\n=== Worked examples (per user) ===")
for b, gt, pos, top3, s_gt, s_top in examples[:5]:
    print(f"user#{b}: true next item = {gt}, its score = {s_gt:.4f}, #{1} ranked candidate score = {s_top:.4f}")
    print(f"   -> true item ranked #{pos+1} out of 100 | HR@10 = {1 if pos<10 else 0} | NDCG@10 = {1/math.log2(pos+2) if pos<10 else 0:.4f}")
    print(f"   -> top-3 recommended items: {top3}")
