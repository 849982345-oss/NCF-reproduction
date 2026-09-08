# Slide 1: NCF 论文复现与暑期学习汇报

**Layout: title-cover**

## Neural Collaborative Filtering (WWW 2017) 复现 · MovieLens-1M

### 徐胜钦 · 重庆大学 大数据与软件学院 · 2026级硕士研究生
2026年9月

---

# Slide 2: 汇报框架
**Layout: column-2-centered**

**Visual: cards-2**

[Card 1: 论文复现 · 约5分钟]
- 为什么复现 NCF
- 论文核心:三种结构
- 复现流程与环境
- 结果对比(我的 vs 论文)
- 问题与差距分析

[Card 2: 暑期学习总结 · 约3分钟]
- 学习内容与方法
- 完成工作与产物
- 遇到的问题
- 后续计划

---

# Slide 3: 为什么复现 NCF

**Visual: cards-3**

[Card 1: 领域代表作]
- 神经协同过滤奠基工作
- WWW 2017,被引数千
- 理解「深度学习如何建模 user-item 交互」的入口

[Card 2: 结果可对照]
- 论文报告 MovieLens-1M 上 HR@10 / NDCG@10
- 代码规模小(数百行),单卡可完整复现
- 有公开权威数字检验复现质量

[Card 3: 为科研打底]
- 官方为 TensorFlow 1,本复现为 PyTorch
- 与序列推荐、GNN 推荐一脉相承
- 对齐后续组内方向(开源生态 / 负责任推荐)

---

# Slide 4: 论文核心——用神经网络学习交互函数

**Visual: comparison-3**

[Column 1: GMF]
- 广义矩阵分解:embedding 逐元素相乘 → 线性打分
- 内积的神经网络推广,建模线性交互

[Column 2: MLP]
- embedding 拼接 → 多层感知机
- 建模非线性交互

[Column 3: NeuMF]
- GMF + MLP 双塔融合后打分
- 论文与本次复现均显示最优

**训练与评估协议:隐式反馈二值化 · log loss · 训练负采样 1:4 · leave-one-out 测试 · HR@10 / NDCG@10**

---

# Slide 5: 复现环境与流程

**Visual: cards-3**

[Card 1: 环境]
- RTX 5070 Ti(16GB)· CUDA 13.4
- PyTorch 2.14.0 + cu130
- 基于 guoyang9/NCF PyTorch 移植,逐模块核对论文结构

[Card 2: 数据]
- MovieLens-1M:6040 用户 / 3706 物品 / 100 万交互
- 论文官方切分:每用户最新一条为测试
- 训练 99.4 万条 / 测试 6040 用户

[Card 3: 协议与超参]
- 测试:1 正 + 99 负 = 100 候选(官方文件)
- factor=8 · MLP 3 层 · batch 256 · lr 1e-3 · 50 epochs
- 每 epoch 全量评估并记录指标

---

# Slide 6: 复现结果——NeuMF 达到并超过论文水平
**Layout: content-centered-a**

**Visual: table**

| 模型 | 复现 HR@10 | 复现 NDCG@10 | 论文 HR@10 | 论文 NDCG@10 | 出处 |
|------|-----------|-------------|-----------|-------------|------|
| GMF | 0.640 | 0.368 | ≈0.645 | ≈0.335 | Figure 4 |
| MLP | 0.675 | 0.396 | ≈0.640 | ≈0.345 | Figure 4 |
| **NeuMF** | **0.690** | **0.412** | **0.688** | **0.410** | Table 2 |
| 结论 | 相对趋势 NeuMF > GMF ≈ MLP 与论文完全一致;绝对差距 < 0.03,源于实现与超参差异 | | | | |

---

# Slide 7: 训练过程——三种结构均稳定收敛
**Layout: content-image-right-a**

**Visual: table**

[Image: training-curves]

左:HR@10 随 epoch 变化(约 20 epoch 进入平台期);右:NDCG@10 随 epoch 变化

---

# Slide 8: 问题分析与踩坑

**Visual: cards-3**

[Card 1: 数据真伪]
- 预下载「ml-1m」实为模拟数据(user 行为序列异常)
- 弃用,改取论文官方处理产物
- 教训:先核数据来源与协议,指标才有意义

[Card 2: 工具链]
- RTX 5070 Ti(sm_120)需 cu128+;系统预装 CPU 版 torch,强制重装 cu130
- GitHub / grouplens 直连不通,代理镜像获取代码与数据

[Card 3: 复现差距归因]
- TF→PyTorch 实现差异 + 训练预算/调参策略不同
- 论文 Figure 数值为图读近似
- 相对趋势一致 → 复现可信

---

# Slide 9: 暑期学习总结
**Layout: content-centered-a**

**Visual: bullets**

- 学习内容:科研方法论(文献检索与精读·三问框架)、PyTorch 实验全链路、Python / 数据分析、Git
- 完成工作:系统入门推荐系统;独立完成 NCF 三结构复现(GMF / MLP / NeuMF)
- 具体产物:完整代码仓库(README·日志·逐 epoch 指标)· 结果图表 · 本汇报
- 遇到的问题:新卡 CUDA 兼容 · 数据源甄别 · 论文数字与复现差异的归因方法
- 后续计划:对齐导师方向(开源生态 / 多智能体 / 负责任推荐);补 GNN 与 LLM 推荐方法;从复现走向改进(负采样策略小实验)

---

# Slide 10: 谢谢聆听

**Visual: hero-statement**

代码仓库、实验日志与结果图表均可现场查看,欢迎老师批评指正
