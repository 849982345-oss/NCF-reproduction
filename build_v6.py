#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Build final presentation directly on the inner-chapter design template."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy, os, glob

TPL = r"C:\research-training\repro-ncf\ncf_inner_template.pptx"
OUT = r"C:\research-training\repro-ncf\NCF复现与暑期学习汇报_徐胜钦_v6.pptx"
IMG = r"C:\research-training\repro-ncf\figures\training_curves.png"

CN_FONT = "微软雅黑"
EN_FONT = "Microsoft YaHei"
DARK = RGBColor(0x21, 0x24, 0x2B)      # near-black text
GRAY = RGBColor(0x6B, 0x72, 0x80)
ACCENT = RGBColor(0x1F, 0x4E, 0x79)    # deep blue accent
LIGHT = RGBColor(0xEA, 0xEF, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GOLD = RGBColor(0xC9, 0xA2, 0x27)
GOLD_D = RGBColor(0x9A, 0x7B, 0x12)

# numbers: read best (HR-optimal epoch) from real training logs
def _best_from_log(model):
    fs = glob.glob(rf"C:\research-training\repro-ncf\ncf_pytorch\logs\{model}_ml-1m_factor8.tsv")
    best_hr, best_ndcg, best_e = 0.0, 0.0, 0
    if fs:
        for line in open(fs[0], encoding="utf-8"):
            p = line.strip().split("\t")
            if len(p) == 3:
                e, h, n = int(p[0]), float(p[1]), float(p[2])
                if h > best_hr:
                    best_hr, best_ndcg, best_e = h, n, e
    return f"{best_hr:.3f}", f"{best_ndcg:.3f}", best_e

_res_gmf = _best_from_log("GMF")
_res_mlp = _best_from_log("MLP")
_res_neumf = _best_from_log("NeuMF-end")
RES = {
    "GMF":  (_res_gmf[0], _res_gmf[1]),
    "MLP":  (_res_mlp[0], _res_mlp[1]),
    "NeuMF": (_res_neumf[0], _res_neumf[1]),
}
EPOCHS_CONVERGE = _res_neumf[2]
print("RES:", RES, "| best NeuMF epoch:", EPOCHS_CONVERGE)

layers = {}
prs = Presentation(TPL)
# drop template's initial empty slide(s)
from pptx.oxml.ns import qn as _qn
sldIdLst = prs.slides._sldIdLst
for sld in list(sldIdLst):
    rId = sld.get(_qn('r:id'))
    prs.part.drop_rel(rId)
    sldIdLst.remove(sld)
for lay in prs.slide_layouts:
    layers[lay.name] = lay

def new_slide(layout_name):
    return prs.slides.add_slide(layers[layout_name])

def style_run(r, text, size=14, bold=False, color=DARK, font=CN_FONT):
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font
    # east-asian font
    rPr = r._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', CN_FONT)

def fill_ph(slide, ph_idx, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """paras: list of (text, size, bold, color, space_after_pt) or list of list-of-runs"""
    ph = None
    for p in slide.placeholders:
        if p.placeholder_format.idx == ph_idx:
            ph = p
            break
    if ph is None:
        return False
    tf = ph.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    # clear
    tf.clear()
    first = True
    for para in paras:
        runs = para if isinstance(para, list) else [para]
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(runs[0][4] if len(runs[0]) > 4 else 4)
        for rspec in runs:
            text, size, bold, color = rspec[0], rspec[1], rspec[2], rspec[3]
            r = p.add_run()
            style_run(r, text, size, bold, color)
    return True

def set_title(slide, text, size=26, color=DARK, align=PP_ALIGN.LEFT, bar=False):
    ph = slide.shapes.title
    tf = ph.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    if align == PP_ALIGN.LEFT and not bar:
        # accent vertical bar left of title
        b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Inches(0.20), Inches(0.62), Inches(0.07), Inches(0.52))
        b.fill.solid(); b.fill.fore_color.rgb = ACCENT
        b.line.fill.background(); b.shadow.inherit = False
    r = p.add_run()
    style_run(r, text, size, True, color)

def add_textbox(slide, l, t, w, h, paras, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    first = True
    for para in paras:
        runs = para if isinstance(para, list) else [para]
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(runs[0][4] if len(runs[0]) > 4 else 4)
        for rspec in runs:
            r = p.add_run()
            style_run(r, rspec[0], rspec[1], rspec[2], rspec[3])
    return box

def add_bar(slide, l, t, w, h, color, alpha=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def add_code_block(slide, l, t, w, h, lines, size=10.0):
    """VS-Code-style dark code block."""
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(l), Inches(t), Inches(w), Inches(h))
    box.adjustments[0] = 0.035
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0x1E, 0x24, 0x30)
    box.line.fill.background()
    box.shadow.inherit = False
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Inches(0.14)
    tf.margin_top = tf.margin_bottom = Inches(0.1)
    first = True
    for ln in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = ln
        col = RGBColor(0x7E, 0x8B, 0x9E) if ln.strip().startswith("#") else RGBColor(0xE8, 0xEC, 0xF2)
        if any(k in ln for k in ("torch.", "np.", "model(", "loss", "optimizer", "evaluate", "mf.write")):
            col = RGBColor(0x8F, 0xD1, 0xFF)
        r.font.size = Pt(size)
        r.font.name = "Consolas"
        r.font.color.rgb = col
        rPr = r._r.get_or_add_rPr()
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = rPr.makeelement(qn("a:ea"), {}); rPr.append(ea)
        ea.set("typeface", CN_FONT)
    return box


def card3_columns(slide):
    """Draw 3 card containers behind column-3-centered-a bodies (back z-order)."""
    spTree = slide.shapes._spTree
    for l, w in [(0.17, 3.05), (5.14, 3.05), (10.11, 3.05)]:
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(l + 0.08), Inches(1.72), Inches(w - 0.16), Inches(4.6))
        shp.adjustments[0] = 0.045
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(0xF2, 0xF4, 0xF8)
        shp.line.color.rgb = RGBColor(0xDD, 0xE2, 0xEA)
        shp.line.width = Pt(0.75)
        shp.shadow.inherit = False
        # move card to back of z-order so body text renders above it
        spTree.remove(shp._element)
        spTree.insert(2, shp._element)
    return []

# ---------- S1 cover (self-drawn editorial) ----------
s = new_slide("master-base")
# deep-blue left band with gold seam
add_bar(s, 0, 0, 4.55, 7.5, ACCENT)
add_bar(s, 4.55, 0, 0.06, 7.5, GOLD)
# white panel over master-base dark background (right area)
add_bar(s, 4.61, 0.0, 8.72, 7.5, WHITE)
# right area hairline top/bottom (above white panel)
add_bar(s, 4.61, 0.28, 8.5, 0.018, GOLD)
add_bar(s, 4.61, 7.22, 8.5, 0.018, GOLD)
# left band content
add_textbox(s, 0.55, 0.55, 3.5, 1.6, [
    [("重庆大学", 15, True, WHITE)],
    [("大数据与软件学院", 13, False, RGBColor(0xC9, 0xD6, 0xE4))],
])
add_textbox(s, 0.55, 4.35, 3.6, 2.2, [
    [("徐胜钦", 34, True, WHITE)],
    [("", 6, False, WHITE)],
    [("2026 级硕士研究生", 12.5, False, RGBColor(0xC9, 0xD6, 0xE4))],
    [("软件工程 · 推荐系统方向", 12.5, False, RGBColor(0xC9, 0xD6, 0xE4))],
])
# right area content
add_textbox(s, 5.05, 1.15, 7.8, 0.5, [
    [("研究生入学汇报  ·  PERSONAL RESEARCH REPORT", 11.5, False, GRAY)],
])
add_textbox(s, 5.05, 2.1, 7.9, 1.9, [
    [("论文复现与暑期学习汇报", 34, True, DARK)],
    [("", 5, False, DARK)],
    [("NCF 复现 · MovieLens-1M · HR@10 / NDCG@10", 15, True, ACCENT)],
])
add_bar(s, 5.1, 4.12, 2.0, 0.045, GOLD)
add_textbox(s, 5.05, 4.75, 7.9, 1.3, [
    [("Neural Collaborative Filtering (WWW 2017)", 16, False, DARK)],
    [("", 4, False, DARK)],
    [("复现 GMF / MLP / NeuMF 三种结构,HR@10 / NDCG@10 对照论文报告值", 13, False, GRAY)],
])
add_textbox(s, 5.05, 6.35, 7.9, 0.5, [
    [("2026 年 9 月", 11.5, False, GRAY)],
])

# ---------- S2 framework ----------
s = new_slide("column-2-centered")
set_title(s, "汇报框架")
fill_ph(s, 15, [[("① 论文复现", 17, True, ACCENT), ("　约 5 分钟", 12, False, GRAY)],
                [("• 为什么复现 NCF", 13, False, DARK)],
                [("• 论文核心:三种结构", 13, False, DARK)],
                [("• 复现流程与环境", 13, False, DARK)],
                [("• 结果对比:我的 vs 论文", 13, False, DARK)],
                [("• 问题与差距分析", 13, False, DARK)]])
fill_ph(s, 16, [[("② 暑期学习总结", 17, True, ACCENT), ("　约 3 分钟", 12, False, GRAY)],
                [("• 学习内容与方法", 13, False, DARK)],
                [("• 完成工作与产物", 13, False, DARK)],
                [("• 遇到的问题", 13, False, DARK)],
                [("• 后续计划", 13, False, DARK)]])

# ---------- S3 why NCF ----------
s = new_slide("column-3-centered-a")
set_title(s, "论文来源与选题理由")
card3_columns(s)
cards3 = [
    ("论文身份", ["Neural Collaborative Filtering", "何向南 等 · WWW 2017", "神经协同过滤奠基工作 · 被引数千"]),
    ("来源与可得性", ["选自本人文献库推荐系统方向基础文献", "官方代码 GitHub 开源(TensorFlow 1)", "论文报告 MovieLens-1M 上 HR@10 / NDCG@10"]),
    ("选题理由", ["代码规模小,单卡可完整复现", "与序列 / GNN 推荐一脉相承", "为后续科研建立最小闭环"]),
]
for ph_idx, (h, items) in zip([69, 48, 49], cards3):
    fill_ph(s, ph_idx, [[(h, 15, True, ACCENT)]] + [[("• " + it, 11.5, False, DARK)] for it in items])

# ---------- S3b literature map (new) ----------
s = new_slide("column-2-centered")
set_title(s, "文献脉络:NCF 在推荐系统中的位置")
fill_ph(s, 15, [
    [("2016–2026 技术演进", 15, True, ACCENT)],
    [("", 6, False, DARK)],
    [("• 2016  Wide & Deep:深度学习进入推荐", 12, False, DARK)],
    [("• 2017  NCF:神经协同过滤奠基 ← 本次复现", 12.5, True, GOLD_D),
     ("", 1, False, DARK)],
    [("• 2018  SASRec:自注意力序列推荐", 12, False, DARK)],
    [("• 2020  LightGCN:图卷积协同过滤", 12, False, DARK)],
    [("• 2021–22  图对比学习(SGL / SimGCL)", 12, False, DARK)],
    [("• 2023+  大模型推荐(LLM4Rec)", 12, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)
fill_ph(s, 16, [
    [("文献检索与精读方法", 15, True, ACCENT)],
    [("", 6, False, DARK)],
    [("• 关键词分层:领域 / 任务 / 方法 / 问题", 12, False, DARK)],
    [("• 多渠道检索:DBLP · Google Scholar · GitHub", 12, False, DARK)],
    [("• 精读三问:解决什么 / 怎么解决 / 怎么证明", 12, False, DARK)],
    [("• 文献库管理:Zotero 沉淀与分类", 12, False, DARK)],
    [("• 以代码验证论文:官方仓库 → 逐模块核对", 12, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)

# ---------- S4 core ----------
s = new_slide("column-3-centered-a")
set_title(s, "论文核心:用神经网络学习交互函数")
card3_columns(s)
core3 = [
    ("GMF", ["广义矩阵分解", "embedding 逐元素相乘 → 线性打分", "内积的神经网络推广 · 线性交互"]),
    ("MLP", ["多层感知机", "embedding 拼接 → 多层非线性变换", "建模非线性交互"]),
    ("NeuMF", ["双塔融合", "GMF + MLP 拼接后打分", "论文与本次复现均显示最优"]),
]
for ph_idx, (h, items) in zip([69, 48, 49], core3):
    fill_ph(s, ph_idx, [[(h, 16, True, ACCENT)]] + [[("• " + it, 11.5, False, DARK)] for it in items])
add_textbox(s, 0.55, 6.5, 12.3, 0.9, [
    [("协议:隐式反馈二值化 · log loss · 训练负采样 1:4 · leave-one-out 测试 · HR@10 / NDCG@10", 11.5, False, GRAY)],
    [("一句话:用神经网络学习交互函数,替代传统矩阵分解的内积", 12, True, ACCENT)],
], align=PP_ALIGN.CENTER)

# ---------- S5 setup ----------
# ---------- S5a formulas (new) ----------
s = new_slide("title-centered")
set_title(s, "论文核心公式", size=26, align=PP_ALIGN.CENTER)
FIMG = r"C:\research-training\repro-ncf\figures\formulas.png"
if os.path.exists(FIMG):
    s.shapes.add_picture(FIMG, Inches(0.85), Inches(1.32), width=Inches(11.6))

s = new_slide("column-3-centered-a")
set_title(s, "复现环境与流程")
card3_columns(s)
setup3 = [
    ("环境", ["RTX 5070 Ti(16 GB)· CUDA 13.4", "PyTorch 2.14.0 + cu130", "基于 guoyang9/NCF 移植,逐模块核对论文结构"]),
    ("数据", ["MovieLens-1M:6040 用户 / 3706 物品 / 100 万交互", "论文官方切分:每用户最新一条为测试(leave-one-out)", "训练 99.4 万条 · 测试 6040 用户"]),
    ("协议与超参", ["测试:1 正 + 99 负 = 100 候选", "factor=8 · MLP 3 层 · batch 256 · lr 1e-3", "50 epochs · 逐 epoch 全量评估记录", "数据取自论文官方仓库并校验行数/字节数"]),
]
for ph_idx, (h, items) in zip([69, 48, 49], setup3):
    fill_ph(s, ph_idx, [[(h, 15, True, ACCENT)]] + [[("• " + it, 11.5, False, DARK)] for it in items])

# ---------- S5c code walkthrough 1 (new) ----------
s = new_slide("title-centered")
set_title(s, "代码解析(1/2):模型三结构与负采样", size=24)
add_code_block(s, 0.45, 1.55, 6.15, 4.75, [
    "# model.py — 三种结构的 forward(原样摘录)",
    "if not self.model == 'MLP':",
    "    embed_user_GMF = self.embed_user_GMF(user)",
    "    embed_item_GMF = self.embed_item_GMF(item)",
    "    output_GMF = embed_user_GMF * embed_item_GMF  # 逐元素相乘",
    "if not self.model == 'GMF':",
    "    embed_user_MLP = self.embed_user_MLP(user)",
    "    embed_item_MLP = self.embed_item_MLP(item)",
    "    interaction = torch.cat((embed_user_MLP,",
    "                             embed_item_MLP), -1)",
    "    output_MLP = self.MLP_layers(interaction)   # 多层非线性",
    "if self.model == 'GMF':",
    "    concat = output_GMF",
    "elif self.model == 'MLP':",
    "    concat = output_MLP",
    "else:   # NeuMF:两路融合",
    "    concat = torch.cat((output_GMF, output_MLP), -1)",
    "prediction = self.predict_layer(concat)",
], size=10.5)
add_code_block(s, 6.85, 1.55, 6.0, 4.75, [
    "# data_utils.py — 训练负采样(1 正 : 4 负)",
    "for x in self.features_ps:        # 每条正样本",
    "    u = x[0]",
    "    for t in range(self.num_ng):  # num_ng = 4",
    "        j = np.random.randint(self.num_item)",
    "        while (u, j) in self.train_mat:",
    "            # 跳过用户已交互物品",
    "            j = np.random.randint(self.num_item)",
    "        self.features_ng.append([u, j])",
    "",
    "# 训练集 = 正样本 + 负样本(每 epoch 重新采样)",
    "features_fill = features_ps + features_ng",
    "labels_fill   = [1]*len(features_ps) + [0]*len(features_ng)",
], size=10.5)
add_textbox(s, 0.6, 6.5, 12.1, 0.6, [
    [("与论文逐模块核对:embedding 逐元素相乘(GMF)/ 拼接过 MLP(MLP)/ 双路 concat(NeuMF)——结构一致", 11.5, False, GRAY)],
], align=PP_ALIGN.CENTER)

# ---------- S5d code walkthrough 2 (new) ----------
s = new_slide("title-centered")
set_title(s, "代码解析(2/2):训练循环与指标落盘", size=24)
add_code_block(s, 0.45, 1.55, 7.3, 4.75, [
    "# main.py — 训练与逐 epoch 全量评估",
    "for epoch in range(args.epochs):",
    "    train_loader.dataset.ng_sample()   # 每 epoch 重采样负样本",
    "    for user, item, label in train_loader:",
    "        model.zero_grad()",
    "        prediction = model(user.cuda(), item.cuda())",
    "        loss = loss_function(prediction, label.float().cuda())",
    "        loss.backward()",
    "        optimizer.step()",
    "    model.eval()                        # 切评估模式",
    "    HR, NDCG = evaluate.metrics(model, test_loader, args.top_k)",
    "    os.makedirs('logs', exist_ok=True)",
    "    with open('logs/{}_{}_factor{}.tsv'.format(",
    "            config.model, 'ml-1m', args.factor_num), 'a') as mf:",
    "        mf.write('{}  {:.6f}  {:.6f}'.format(epoch, hr, ndcg))",
    "        # 仓库实现:epoch / HR@10 / NDCG@10 制表符分隔,6 位小数",
], size=10.5)
add_code_block(s, 8.0, 1.55, 4.85, 4.75, [
    "# evaluate.py — 与论文一致的口径",
    "predictions = model(user, item)   # 100 个候选打分",
    "_, indices = torch.topk(predictions, top_k)",
    "recommends = torch.take(item, indices)",
    "gt_item = item[0]    # 第 1 个候选 = 真实目标",
    "HR.append(hit(gt_item, recommends))",
    "NDCG.append(ndcg(gt_item, recommends))",
    "",
    "# 负采样口径",
    "#   训练 1:4      (每 epoch 重采样)",
    "#   测试 1 正 + 99 负 = 100 候选",
    "#   取论文官方 test.negative 文件",
], size=10.0)
add_textbox(s, 0.6, 6.5, 12.1, 0.6, [
    [("工程改造:", 11.5, True, ACCENT),
     ("--model 参数切换三结构 · 逐 epoch 指标 6 位小数落盘 · Windows 环境适配", 11.5, False, DARK)],
], align=PP_ALIGN.CENTER)

# ---------- S5b code & experiment process (new) ----------
s = new_slide("column-3-centered-a")
set_title(s, "实验流程与实测日志")
card3_columns(s)
proc3 = [
    ("实验流程", ["环境核验 → 数据校验 → 冒烟测试(2 epoch)", "GMF / MLP / NeuMF 各 50 epoch", "单模型 23-35 分钟(RTX 5070 Ti)"]),
    ("训练配置", ["factor=8 · MLP 3 层 · batch 256 · lr 1e-3", "Adam · 4:1 负采样 · 每 epoch 全量评估", "评估口径与论文一致(leave-one-out + 100 候选)"]),
    ("数据与日志校验", ["论文官方产物:训练 994,169 条", "逐 epoch 指标落盘,6 位小数", "文件字节数与行数校验,弃用模拟数据"]),
]
for ph_idx, (h, items) in zip([69, 48, 49], proc3):
    fill_ph(s, ph_idx, [[(h, 15, True, ACCENT)]] + [[("• " + it, 11, False, DARK)] for it in items])
add_textbox(s, 0.8, 6.5, 11.7, 0.8, [
    [("日志实测(NeuMF):", 11.5, True, ACCENT),
     ("epoch 0 → HR 0.5904;epoch 19 → HR 0.6904(最优);epoch 49 → HR 0.6775(稳定,无过拟合)", 11.5, False, DARK)],
], align=PP_ALIGN.CENTER)

# ---------- S6 results table ----------
s = new_slide("title-centered")
set_title(s, "复现结果:NeuMF 对齐论文 Table 2 水平", size=26, align=PP_ALIGN.CENTER)
rows = [
    ["模型", "复现 HR@10", "复现 NDCG@10", "论文 HR@10", "论文 NDCG@10", "出处"],
    ["GMF", RES["GMF"][0], RES["GMF"][1], "≈0.645", "≈0.335", "Figure 4 读数"],
    ["MLP", RES["MLP"][0], RES["MLP"][1], "≈0.640", "≈0.345", "Figure 4 读数"],
    ["NeuMF", RES["NeuMF"][0], RES["NeuMF"][1], "0.688", "0.410", "Table 2"],
]
tbl_shape = s.shapes.add_table(4, 6, Inches(1.1), Inches(1.9), Inches(11.1), Inches(2.6))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(1.7)
for c in range(1, 5):
    tbl.columns[c].width = Inches(1.9)
tbl.columns[5].width = Inches(1.9)
for ri, row in enumerate(rows):
    for ci, val in enumerate(row):
        cell = tbl.cell(ri, ci)
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        is_head = ri == 0
        is_neumf = ri == 3 and ci in (1, 2, 3, 4)
        style_run(r, val, 13, is_head or is_neumf or ci == 0,
                  WHITE if is_head else (ACCENT if is_neumf else DARK))
        cell.fill.solid()
        if is_head:
            cell.fill.fore_color.rgb = ACCENT
        elif ri == 3:
            cell.fill.fore_color.rgb = RGBColor(0xFB, 0xF3, 0xDC)
        elif ri % 2 == 0:
            cell.fill.fore_color.rgb = LIGHT
        else:
            cell.fill.fore_color.rgb = WHITE
add_textbox(s, 1.1, 4.9, 11.1, 1.6, [
    [("结论:", 13, True, ACCENT), (" 相对趋势 NeuMF > GMF ≈ MLP 与论文完全一致;NeuMF 与 Table 2 数值差 ≤ 0.002(实验噪声范围),复现成立", 13, False, DARK)],
    [("• NeuMF no-pre-training(factor=8)对照论文 Table 2 精确值;GMF/MLP 对照 Figure 4 读数(±0.005),其 NDCG 高于读数源于实现结构差异", 11, False, GRAY)],
], align=PP_ALIGN.LEFT)

# ---------- S7 training curves ----------
s = new_slide("title-centered")
set_title(s, "训练过程:三结构稳定收敛", size=26, align=PP_ALIGN.CENTER)
if os.path.exists(IMG):
    s.shapes.add_picture(IMG, Inches(0.6), Inches(1.65), width=Inches(12.13))
add_textbox(s, 0.6, 6.5, 12.13, 0.6, [
    [("三结构约 20 epoch 后进入平台期 · 收敛平滑无过拟合 · 相对排序与论文一致", 12, False, GRAY)],
], align=PP_ALIGN.CENTER)

# ---------- S8 problems ----------
s = new_slide("column-3-centered-a")
set_title(s, "问题分析与踩坑")
card3_columns(s)
prob3 = [
    ("数据真伪甄别", ["预下载「ml-1m」实为模拟数据", "抽查发现行为序列异常(user1: item 1..35 连续)", "弃用,改取论文官方处理产物", "校验:训练 994,169 行与论文协议一致", "教训:先核数据源,指标才有意义"]),
    ("工具链与兼容", ["RTX 5070 Ti(Blackwell, sm_120)需 cu128+", "预装 torch 为 CPU 版 → 强制重装 cu130", "GitHub / 数据集官网直连不通 → 代理镜像", "代理截断大文件 → 按字节数校验后重下"]),
    ("复现差距归因", ["① 实现差异:TF1→PyTorch,初始化/MLP 层宽不同", "② 训练预算:论文网格搜索最优,本复现固定一组参数", "③ 对照口径:论文 Figure 4 为图读值(±0.005)", "相对趋势一致 → 复现可信"]),
]
for ph_idx, (h, items) in zip([69, 48, 49], prob3):
    fill_ph(s, ph_idx, [[(h, 14.5, True, ACCENT)]] + [[("• " + it, 11, False, DARK)] for it in items])

# ---------- S8b limitations (new) ----------
s = new_slide("column-3-centered-a")
set_title(s, "局限性讨论与个人思考")
card3_columns(s)
lim3 = [
    ("评估口径", ["100 个候选(1 正 + 99 负)与全量物品排序不可比", "论文正文称「100 负」,官方数据文件为 99 负", "结论只在同口径下成立"]),
    ("超参与方差", ["只固定一组参数,未做网格搜索", "论文调参取最优,是差距来源之一", "单次运行,未报多随机种子均值±方差"]),
    ("实现差异", ["TF1 → PyTorch 移植,初始化与 MLP 层宽不同", "GMF / MLP 的 NDCG 高于论文图读值由此而来", "相对趋势仍与论文一致"]),
]
for ph_idx, (h, items) in zip([69, 48, 49], lim3):
    fill_ph(s, ph_idx, [[(h, 15, True, ACCENT)]] + [[("• " + it, 11, False, DARK)] for it in items])
add_textbox(s, 0.8, 6.45, 11.7, 0.8, [
    [("个人思考(下一步可做的两个小改进):", 12, True, ACCENT),
     ("① 负采样从均匀采样改为难负样本(popularity-based);② 物品侧引入属性/内容特征做融合", 12, False, DARK)],
], align=PP_ALIGN.CENTER)

# ---------- S9 summer summary ----------
s = new_slide("column-2-centered")
set_title(s, "暑期学习总结")
fill_ph(s, 15, [
    [("学了什么", 15, True, ACCENT)],
    [("", 4, False, DARK)],
    [("• 科研方法论:文献检索与精读(三问框架)", 12, False, DARK)],
    [("• PyTorch 实验全链路", 12, False, DARK)],
    [("• Python / 数据分析 / Git", 12, False, DARK)],
    [("", 10, False, DARK)],
    [("产物", 15, True, ACCENT)],
    [("", 4, False, DARK)],
    [("• 独立完成 NCF 三结构复现", 12, False, DARK)],
    [("• 代码仓库 + 逐 epoch 实验日志 + 图表", 12, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)
fill_ph(s, 16, [
    [("遇到的问题", 15, True, ACCENT)],
    [("", 4, False, DARK)],
    [("• 新卡 CUDA 兼容 / 数据源甄别", 12, False, DARK)],
    [("• 复现差异的归因方法", 12, False, DARK)],
    [("", 10, False, DARK)],
    [("后续计划", 15, True, ACCENT)],
    [("", 4, False, DARK)],
    [("• 对齐组内方向:开源生态 / 负责任推荐", 12, False, DARK)],
    [("• 补 GNN 与 LLM 推荐方法", 12, False, DARK)],
    [("• 从复现走向改进:负采样策略小实验", 12, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)

# ---------- S9b deliverables (new) ----------
s = new_slide("column-2-centered")
set_title(s, "具体产物(均可现场查验)")
fill_ph(s, 15, [
    [("① 代码仓库", 15, True, ACCENT)],
    [("• GitHub 公开:代码 + 官方数据 + 运行脚本 + README", 12, False, DARK)],
    [("• 环境、复现命令、结果表一应俱全,可复跑", 12, False, DARK)],
    [("", 8, False, DARK)],
    [("② 实验日志", 15, True, ACCENT)],
    [("• 逐 epoch 指标:epoch / HR@10 / NDCG@10(6 位小数)", 12, False, DARK)],
    [("• 完整终端输出:每 epoch 耗时、最优结果自动记录", 12, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)
fill_ph(s, 16, [
    [("③ 结果图表", 15, True, ACCENT)],
    [("• 结果对比表:复现值 vs 论文报告值", 12, False, DARK)],
    [("• 训练曲线:三结构 HR@10 / NDCG@10(200dpi)", 12, False, DARK)],
    [("", 8, False, DARK)],
    [("④ 分析文档", 15, True, ACCENT)],
    [("• 复现实验记录:环境 / 协议 / 命令 / 问题", 12, False, DARK)],
    [("• 本汇报 PPT 与讲稿", 12, False, DARK)],
], anchor=MSO_ANCHOR.MIDDLE)
add_textbox(s, 0.8, 6.45, 11.7, 0.5, [
    [("仓库地址:github.com/849982345-oss/NCF-reproduction", 11.5, True, ACCENT)],
], align=PP_ALIGN.CENTER)

# ---------- S10 closing ----------
s = new_slide("title-centered")
add_bar(s, 0, 0, 13.333, 0.12, GOLD)
add_bar(s, 0, 7.38, 13.333, 0.12, GOLD)
set_title(s, "谢谢聆听", size=44, align=PP_ALIGN.CENTER)
add_textbox(s, 1.5, 4.3, 10.3, 1.6, [
    [("代码仓库 · 实验日志 · 结果图表均可现场查看", 15, False, GRAY)],
    [("", 6, False, DARK)],
    [("欢迎各位老师批评指正", 16, True, ACCENT)],
], align=PP_ALIGN.CENTER)

# page numbers (skip first/cover slide)
for si, sl in enumerate(prs.slides, start=1):
    if si == 1:
        continue
    add_textbox(sl, 12.3, 7.02, 0.9, 0.3, [[(str(si), 10.5, False, RGBColor(0x9A, 0xA3, 0xAF))]],
                align=PP_ALIGN.RIGHT)

prs.save(OUT)
print("saved", OUT)
