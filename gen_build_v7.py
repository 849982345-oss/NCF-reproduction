#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Generate build_v7.py from build_v6.py: voice-over prompt bar per page, footer with repo URL,
speaker notes (talk script), remove leftover bottom lines that would collide."""
src = open(r"C:\research-training\repro-ncf\build_v6.py", encoding="utf-8").read()

src = src.replace("NCF复现与暑期学习汇报_徐胜钦_v6.pptx", "NCF复现与暑期学习汇报_徐胜钦_v7.pptx")

# ---- 1. helper: voice-over bar ----
helper_anchor = "def card3_columns(slide):"
voice_helper = '''def add_voice_bar(slide, text):
    """Light prompt bar at page bottom carrying the spoken cue line."""
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(0.55), Inches(6.60), Inches(12.23), Inches(0.42))
    bar.adjustments[0] = 0.25
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(0xEA, 0xF1, 0xF9)
    bar.line.fill.background()
    bar.shadow.inherit = False
    tf = bar.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.16); tf.margin_right = Inches(0.1)
    tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r1 = p.add_run(); style_run(r1, "口播  ", 10, True, ACCENT)
    r2 = p.add_run(); style_run(r2, text, 11, False, DARK)
    return bar


'''
assert helper_anchor in src
src = src.replace(helper_anchor, voice_helper + helper_anchor, 1)

# ---- 2. drop colliding bottom lines ----
drop = [
'''add_textbox(s, 0.6, 6.5, 12.1, 0.6, [
    [("与论文逐模块核对:embedding 逐元素相乘(GMF)/ 拼接过 MLP(MLP)/ 双路 concat(NeuMF)——结构一致", 11.5, False, GRAY)],
], align=PP_ALIGN.CENTER)
''',
'''add_textbox(s, 0.6, 6.5, 12.1, 0.6, [
    [("工程改造:", 11.5, True, ACCENT),
     ("--model 参数切换三结构 · 逐 epoch 指标 6 位小数落盘 · Windows 环境适配", 11.5, False, DARK)],
], align=PP_ALIGN.CENTER)
''',
'''add_textbox(s, 0.8, 6.5, 11.7, 0.8, [
    [("日志实测(NeuMF):", 11.5, True, ACCENT),
     ("epoch 0 → HR 0.5904;epoch 19 → HR 0.6904(最优);epoch 49 → HR 0.6775(稳定,无过拟合)", 11.5, False, DARK)],
], align=PP_ALIGN.CENTER)
''',
]
for d in drop:
    assert d in src, d[:60]
    src = src.replace(d, "", 1)

# move curves-slide caption up, and limitations "personal thought" line up
src = src.replace('add_textbox(s, 0.6, 6.5, 12.13, 0.6, [\n    [("三结构约 20 epoch 后进入平台期 · 收敛平滑无过拟合 · 相对排序与论文一致", 12, False, GRAY)],',
                  'add_textbox(s, 0.6, 6.22, 12.13, 0.34, [\n    [("三结构约 20 epoch 后进入平台期 · 收敛平滑无过拟合 · 相对排序与论文一致", 11, False, GRAY)],')
src = src.replace('add_textbox(s, 0.8, 6.45, 11.7, 0.8, [', 'add_textbox(s, 0.8, 6.22, 11.7, 0.34, [')
src = src.replace('[("个人思考(下一步可做的两个小改进):", 12, True, ACCENT),\n     ("① 负采样从均匀采样改为难负样本(popularity-based);② 物品侧引入属性/内容特征做融合", 12, False, DARK)]',
                  '[("个人思考:", 10.5, True, ACCENT),\n     ("① 负采样改为难负样本 ② 物品侧引入属性融合", 10.5, False, DARK)]')

# ---- 3. footer with repo URL + page number, voice bar, and speaker notes ----
old_pager = '''# page numbers (skip first/cover slide)
for si, sl in enumerate(prs.slides, start=1):
    if si == 1:
        continue
    add_textbox(sl, 12.3, 7.02, 0.9, 0.3, [[(str(si), 10.5, False, RGBColor(0x9A, 0xA3, 0xAF))]],
                align=PP_ALIGN.RIGHT)'''
new_pager = '''# ---- voice-over bars, footer (repo URL + page no.), speaker notes ----
VOICE = {
 2: "我是 2026 级徐胜钦,跨专业到软件工程,方向推荐系统;今天汇报:一篇论文复现 + 假期学习总结。",
 3: "复现的是 WWW 2017 的 NCF——神经协同过滤的奠基工作,官方开源、有权威对照数字。",
 4: "NCF 承接 Wide&Deep,启发了后面的 SASRec 与 LightGCN——这是我文献调研的脉络。",
 5: "论文核心:用神经网络替代内积建模交互——GMF 线性、MLP 非线性、NeuMF 双路融合,复现中 NeuMF 最好。",
 6: "三种结构的预测式与二分类损失;下面两行是 HR@10 与 NDCG@10 的定义式。",
 7: "环境 RTX 5070 Ti + PyTorch cu130;数据用论文官方处理产物,切分与评估协议和论文一致。",
 8: "这是仓库真实代码:三结构 forward 与 1:4 负采样,与论文逐模块核对过。",
 9: "训练循环五步,每个 epoch 结束后对 6040 个测试用户做全量评估,指标 6 位小数落盘。",
 10: "流程:冒烟测试确认链路后,三结构各跑 50 个 epoch,单模型 23 到 35 分钟。",
 11: "NeuMF 复现 0.690 / 0.409,论文 0.688 / 0.410,差 ≤ 0.002;相对趋势完全一致。",
 12: "三个模型都稳定收敛,约 20 个 epoch 进入平台期,曲线平滑、无过拟合。",
 13: "三个真实的坑:数据真伪甄别、新卡 CUDA 兼容、代理下载截断;差距归因分三层。",
 14: "如实说明三点局限:评估口径、超参与方差、实现差异;下一步想试难负采样的小改进。",
 15: "假期主线是打通最小科研闭环:从读懂论文到跑通实验再到结果对比。",
 16: "四类产物都可以现场查验:代码仓库、实验日志、结果图表、分析文档。",
 17: "汇报结束,欢迎各位老师提问;仓库地址在页脚,全部材料都可现场查看。",
}
from pptx.util import Pt as _Pt
for si, sl in enumerate(prs.slides, start=1):
    if si == 1:
        continue
    # voice-over prompt bar (carry the spoken cue)
    if si in VOICE:
        add_voice_bar(sl, VOICE[si])
    # footer: repo URL (left) + page number (right)
    add_textbox(sl, 0.55, 7.06, 7.5, 0.28,
                [[("github.com/849982345-oss/NCF-reproduction", 8.5, False, RGBColor(0xA8, 0xB0, 0xBC))]])
    add_textbox(sl, 12.15, 7.06, 0.65, 0.28,
                [[(str(si), 9.5, False, RGBColor(0xA8, 0xB0, 0xBC))]], align=PP_ALIGN.RIGHT)

# speaker notes = talk-script paragraphs (view in Presenter View)
NOTES = {
 2: "各位老师好,我是 2026 级硕士研究生徐胜钦,本科新能源科学与工程,跨专业到软件工程,研究方向是推荐系统。下面汇报两部分:一篇论文的复现情况,和假期学习总结。",
 3: "我复现的是 Neural Collaborative Filtering,何向南等作者,发表于 WWW 2017。它是用深度学习做协同过滤这条线的奠基工作,被引数千次;官方代码 GitHub 开源(TensorFlow 1);论文报告了 MovieLens-1M 上的 HR@10 和 NDCG@10,有明确可对照的数字。",
 4: "选它之前我做了文献调研:NCF 承接 2016 年的 Wide & Deep,之后 2018 年出现 SASRec 把自注意力引入序列推荐,2020 年 LightGCN 把图卷积简化到极致,再往后是对比学习和大模型推荐。NCF 正好在链条的起点位置,适合作为第一篇复现。",
 5: "论文要解决的核心问题是:协同过滤里怎么建模用户和物品的交互。传统矩阵分解用内积,表达能力受限;论文用神经网络学习交互函数,提出三种结构:GMF 把 embedding 逐元素相乘再线性打分,建模线性交互;MLP 把 embedding 拼接后过多层感知机,建模非线性;NeuMF 融合两路。训练用隐式反馈二分类、log loss、1 正比 4 负采样;评估用 leave-one-out,每个用户 100 个候选,算 HR@10 和 NDCG@10。",
 6: "这是核心公式:GMF 的预测式、MLP 的前向、NeuMF 的融合式,以及训练用的二分类 log loss。下面两个是评估指标的定义式——HR@10 是命中率,NDCG@10 还考虑了排名位置。",
 7: "环境是本机 RTX 5070 Ti 16G,PyTorch 2.14 加 cu130。数据用的是论文官方仓库处理好的 MovieLens-1M:训练 99.4 万条,测试是 6040 个用户、每个用户 1 个真实目标加 99 个随机负样本,和论文协议完全一致。超参固定 factor=8、batch 256、学习率 1e-3。",
 8: "这是仓库里的真实代码,左边是 model.py 里三种结构的 forward:可以看到 GMF 的逐元素相乘、MLP 的拼接进多层网络、NeuMF 的双路 cat 融合。右边是负采样:每条正样本采 4 条负样本,并且用 while 循环跳过用户已经交互过的物品。这些逻辑都和论文逐模块核对过。",
 9: "这是训练循环:清梯度、前向、算 loss、反向、更新参数五步;每个 epoch 结束后切换到评估模式,对全部 6040 个测试用户做全量评估,然后把指标以 6 位小数写进日志文件。这是我自己加的三处工程改造之一,另外两处是 --model 参数一键切换三种结构、以及 Windows 环境适配。",
 10: "实验流程是:环境核验、数据校验、先用 2 个 epoch 冒烟测试确认整条链路通,然后三种结构各跑 50 个 epoch,每个 epoch 都在全量测试用户上评估。三个模型在 5070 Ti 上各跑了 23 到 35 分钟。",
 11: "这是复现结果对照论文:NeuMF 是我复现到最好的——HR@10 是 0.690,NDCG@10 是 0.409;论文 Table 2 报告的是 0.688 和 0.410,两项差距都在 0.002 以内,复现成立。GMF 复现 0.640,论文图里读数约 0.645;MLP 是 0.675。关键相对趋势 NeuMF 最好、GMF 和 MLP 接近,和论文结论完全一致。",
 12: "这是三个模型 50 个 epoch 的训练曲线:都在 20 个 epoch 左右进入平台期,曲线平滑,没有过拟合迹象,相对排序也保持和论文一致。",
 13: "复现过程中有三个值得说的坑:第一,数据真伪——我最初拿到一份预处理数据,抽查发现行为序列异常,判断是模拟数据直接弃用,改从论文官方仓库取数并校验行数;第二,显卡兼容——5070 Ti 是新架构,需要 cu128 以上的 PyTorch,机器上预装的是 CPU 版,我卸载后重装了 cu130;第三,网络问题——GitHub 和数据集官网直连不通,用代理镜像拿到代码和数据,代理还会截断大文件,我按字节数校验后重新下载。结果差距我归因三层:实现差异、训练预算、对照口径。",
 14: "局限我也如实说:第一,评估口径是 100 个候选,和全量物品排序不可比,而且论文正文说 100 个负样本、官方数据文件实际是 99 个;第二,我只固定了一组超参,没有做网格搜索,也是差距来源之一;第三,单次运行没有报多随机种子的方差。个人思考:下一步可以先试难负采样,再考虑物品侧引入属性特征。",
 15: "假期学习围绕一条主线:能不能独立完成一次论文复现。学习内容包括科研方法论——文献检索和精读的三问框架、PyTorch 的实验全链路、Python 与数据分析基础、Git 工作流。完成的工作是系统梳理了推荐系统基础,并独立完成 NCF 从读到跑通到对比的全过程。遇到的问题:跨专业基础差距,我的策略是先跑通最小闭环。后续计划分三步:对齐组内方向、从复现走向改进、长期做出自己的第一篇工作。",
 16: "产物一共四类,都可以现场查验:第一是代码仓库,已经推上 GitHub,包含代码、官方数据和运行脚本;第二是实验日志,每个 epoch 的指标 6 位小数落盘,完整终端输出也保留;第三是结果图表,包括对比表和训练曲线;第四是分析文档和这份 PPT。",
 17: "我的汇报到这里,代码仓库、实验日志和结果图表都可以现场查看,欢迎各位老师批评指正。",
}
for si, sl in enumerate(prs.slides, start=1):
    if si in NOTES:
        sl.notes_slide.notes_text_frame.text = NOTES[si]'''
assert old_pager in src
src = src.replace(old_pager, new_pager, 1)

open(r"C:\research-training\repro-ncf\build_v7.py", "w", encoding="utf-8").write(src)
print("build_v7.py generated")
