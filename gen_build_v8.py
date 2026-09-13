#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Generate build_v8.py from build_v7.py:
- remove bottom voice-over bars (content goes into the slides themselves)
- replace the literature-map slide with a proper '研究背景与研究问题' slide
- enrich selection-reason card with the technology-lineage position
"""
src = open(r"C:\research-training\repro-ncf\build_v7.py", encoding="utf-8").read()
src = src.replace("NCF复现与暑期学习汇报_徐胜钦_v7.pptx", "NCF复现与暑期学习汇报_徐胜钦_v8.pptx")

# 1) drop voice-over bar calls
old_call = '''    # voice-over prompt bar (carry the spoken cue)
    if si in VOICE:
        add_voice_bar(sl, VOICE[si])
'''
assert old_call in src
src = src.replace(old_call, "", 1)

# 2) replace literature-map slide with background & problem slide
start = src.index("# ---------- S3b literature map (new) ----------")
end = src.index("# ---------- S4 core ----------")
new_slide = '''# ---------- S3b background & research problem ----------
s = new_slide("column-2-centered")
set_title(s, "研究背景与研究问题:这篇论文要解决什么")
fill_ph(s, 15, [
    [("研究背景与任务", 15, True, ACCENT)],
    [("", 5, False, DARK)],
    [("• 推荐系统两类范式:基于内容过滤 与 协同过滤(CF)", 11.5, False, DARK)],
    [("• 协同过滤的核心:从用户-物品交互中学习偏好", 11.5, False, DARK)],
    [("• 主流做法 矩阵分解(MF):用户与物品映射为隐向量,用内积表示交互", 11.5, False, DARK)],
    [("• 本文任务:隐式反馈下的 Top-N 推荐(点击 / 购买等行为,而非显式评分)", 11.5, False, DARK)],
    [("• 技术演进位置:NCF 承接 Wide&Deep(2016),启发 SASRec / LightGCN 等后续工作", 11.5, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)
fill_ph(s, 16, [
    [("已有方法的不足 → 论文的问题", 15, True, ACCENT)],
    [("", 5, False, DARK)],
    [("• 内积表达能力受限:本质是线性模型,难以拟合复杂交互", 11.5, False, DARK)],
    [("• 隐向量各维度被假设相互独立,交互建模受到约束", 11.5, False, DARK)],
    [("• NCF 之前的工作多是 MF 的变体,或引入辅助信息,交互仍靠内积", 11.5, False, DARK)],
    [("", 4, False, DARK)],
    [("论文主张与贡献", 15, True, ACCENT)],
    [("• 用神经网络学习交互函数 f(user, item),替代内积", 11.5, False, DARK)],
    [("• 提出 GMF / MLP / NeuMF 三种结构,实验超越 BPR、eALS 等方法", 11.5, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)
add_textbox(s, 0.7, 6.15, 11.9, 0.42, [
    [("核心论点:", 12, True, ACCENT),
     ("把内积替换为神经网络,能提升协同过滤的表达能力与推荐效果", 12, False, DARK)],
], align=PP_ALIGN.CENTER)

'''
src = src[:start] + new_slide + src[end:]

# 3) selection-reason card: fold lineage in explicitly
src = src.replace('("选题理由", ["代码规模小,单卡可完整复现", "与序列 / GNN 推荐一脉相承", "为后续科研建立最小闭环"]),',
                  '("选题理由", ["内积→神经网络的思路起点,方法代表性最强", "与 SASRec / LightGCN 一脉相承", "代码规模小,单卡可完整复现,适合建立最小科研闭环"]),')

open(r"C:\research-training\repro-ncf\build_v8.py", "w", encoding="utf-8").write(src)
print("build_v8.py generated")
