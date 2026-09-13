#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Generate build_v6.py from build_v5.py: insert 5 new slides (literature map, formulas,
code walkthrough x2, limitations), and retitle the process slide."""
import io

src = open(r"C:\research-training\repro-ncf\build_v5.py", encoding="utf-8").read()

# output name
src = src.replace("NCF复现与暑期学习汇报_徐胜钦_v5.pptx", "NCF复现与暑期学习汇报_徐胜钦_v6.pptx")

# ---- add helper: monospace code block ----
helper_anchor = "def card3_columns(slide):"
code_helper = '''def add_code_block(slide, l, t, w, h, lines, size=10.0):
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


'''
assert helper_anchor in src
src = src.replace(helper_anchor, code_helper + helper_anchor, 1)

# ---------------- NEW SLIDE 1: literature map (before S4 core) ----------------
anchor_s4 = "# ---------- S4 core ----------"
lit_slide = '''# ---------- S3b literature map (new) ----------
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

'''
assert anchor_s4 in src
src = src.replace(anchor_s4, lit_slide + anchor_s4, 1)

# ---------------- NEW SLIDE 2: formulas (before S5b) ----------------
anchor_s5b = "# ---------- S5b code & experiment process (new) ----------"
formula_slide = '''# ---------- S5a formulas (new) ----------
s = new_slide("title-centered")
set_title(s, "论文核心公式", size=26, align=PP_ALIGN.CENTER)
FIMG = r"C:\\research-training\\repro-ncf\\figures\\formulas.png"
if os.path.exists(FIMG):
    s.shapes.add_picture(FIMG, Inches(0.75), Inches(1.5), width=Inches(11.8))
add_textbox(s, 0.8, 6.6, 11.7, 0.6, [
    [("交互函数从内积推广到神经网络:线性(GMF)+ 非线性(MLP)双路融合,用二分类损失训练", 12, True, ACCENT)],
], align=PP_ALIGN.CENTER)

'''
anchor_s5 = 's = new_slide("column-3-centered-a")\nset_title(s, "复现环境与流程")'
assert anchor_s5 in src
src = src.replace(anchor_s5, formula_slide + anchor_s5, 1)

# ---------------- NEW SLIDES 3&4: code walkthrough (before S5b) ----------------
code_slides = '''# ---------- S5c code walkthrough 1 (new) ----------
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

'''
src = src.replace(anchor_s5b, code_slides + anchor_s5b, 1)

# ---------------- retitle & simplify the process slide ----------------
src = src.replace('set_title(s, "代码与实验过程")', 'set_title(s, "实验流程与实测日志")')
src = src.replace('''proc3 = [
    ("数据准备", ["论文官方处理产物(非第三方预处理)", "训练 994,169 条 · 测试 6040 用户 × 100 候选", "校验行数与字节数,弃用模拟数据"]),
    ("代码核对与改造", ["逐模块核对:结构 / 损失 / 负采样 / 评估", "改造 1:--model 参数一键切换三结构", "改造 2:逐 epoch 指标落盘(6 位小数)", "改造 3:Windows 环境适配"]),
    ("实验流程", ["环境核验 → 数据校验 → 冒烟测试(2 epoch)", "GMF / MLP / NeuMF 各 50 epoch", "单模型 23-35 分钟(RTX 5070 Ti)"]),
]''',
'''proc3 = [
    ("实验流程", ["环境核验 → 数据校验 → 冒烟测试(2 epoch)", "GMF / MLP / NeuMF 各 50 epoch", "单模型 23-35 分钟(RTX 5070 Ti)"]),
    ("训练配置", ["factor=8 · MLP 3 层 · batch 256 · lr 1e-3", "Adam · 4:1 负采样 · 每 epoch 全量评估", "评估口径与论文一致(leave-one-out + 100 候选)"]),
    ("数据与日志校验", ["论文官方产物:训练 994,169 条", "逐 epoch 指标落盘,6 位小数", "文件字节数与行数校验,弃用模拟数据"]),
]''')

# ---------------- NEW SLIDE 5: limitations (before S9) ----------------
anchor_s9 = "# ---------- S9 summer summary ----------"
lim_slide = '''# ---------- S8b limitations (new) ----------
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

'''
assert anchor_s9 in src
src = src.replace(anchor_s9, lim_slide + anchor_s9, 1)

# add GOLD_D color definition (darker gold for text use)
src = src.replace('GOLD = RGBColor(0xC9, 0xA2, 0x27)',
                  'GOLD = RGBColor(0xC9, 0xA2, 0x27)\nGOLD_D = RGBColor(0x9A, 0x7B, 0x12)')

open(r"C:\research-training\repro-ncf\build_v6.py", "w", encoding="utf-8").write(src)
print("build_v6.py generated")
