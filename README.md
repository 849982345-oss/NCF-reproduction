# NCF (Neural Collaborative Filtering) Reproduction — 论文复现

> 复现论文:He Xiangnan et al., *Neural Collaborative Filtering*, **WWW 2017**.
> 官方代码(基准):https://github.com/hexiangnan/neural_collaborative_filtering (TensorFlow 1)
> 本仓库实现:基于开源 PyTorch 移植 [guoyang9/NCF](https://github.com/guoyang9/NCF)(Apache-2.0),按论文协议核对模型结构与训练流程后复现。

## 环境
- GPU:NVIDIA GeForce RTX 5070 Ti (16GB), CUDA 13.4 驱动
- Python 3.13 + PyTorch 2.7.1+cu130
- 依赖:torch / numpy / pandas / scipy

## 数据
- MovieLens-1M(官方 ml-1m 处理产物,来自 hexiangnan 官方 repo Data/ 目录):
  - `ml1m_data/ml-1m.train.rating`(≈99.4 万条训练交互)
  - `ml1m_data/ml-1m.test.rating`(6040 条,每用户最后一条交互)
  - `ml1m_data/ml-1m.test.negative`(6040 行,每行 100 个随机负样本)
- 评估协议(与论文一致):leave-one-out(每用户最新交互为测试),100 随机负样本 + 1 正样本,指标 HR@10 / NDCG@10,训练负采样 1:4。

## 复现运行
```bash
# GMF / MLP / NeuMF(从零训练,end-to-end)
python main.py --model GMF      --epochs 50 --factor_num 8 --batch_size 256 --lr 0.001
python main.py --model MLP      --epochs 50 --factor_num 8 --batch_size 256 --lr 0.001
python main.py --model NeuMF-end --epochs 50 --factor_num 8 --batch_size 256 --lr 0.001
```

## 结果(factor=8, MovieLens-1M)
| 模型 | 复现 HR@10 | 复现 NDCG@10 | 论文 HR@10 | 论文 NDCG@10 | 备注 |
|---|---|---|---|---|---|
| GMF | 0.6396 | 0.3706 | ≈0.645 (Fig.4 读数) | ≈0.335 | |
| MLP | 0.6750 | 0.3992 | ≈0.640 (Fig.4 读数) | ≈0.345 | |
| NeuMF | 0.6904 | 0.4131 | 0.688 (Table 2) | 0.410 (Table 2) | 论文 no-pretrain |

日志:每次运行输出到 `logs/{model}_factor{factor}.log`;指标由 main.py 每 epoch 打印。

## 文件
- `main.py` 训练/评估入口(含 `--model` 参数,便于复现三种结构)
- `model.py` NCF 模型(GMF/MLP/NeuMF),与论文结构一致
- `evaluate.py` HR@10/NDCG@10(与论文协议一致)
- `data_utils.py` 数据加载与负采样
- `config.py` 数据路径与默认配置

## 复现差异说明(诚实标注)
1. 论文官方为 TensorFlow 1(Keras)实现;本复现采用经社区验证的 PyTorch 移植,逐模块核对(见上方对照)。
2. 论文 Figure 4 的 GMF/MLP 数值由图中读取,存在 ±0.005 误差;NeuMF 以论文 Table 2(no pre-training 列)为准。
3. 训练细节:论文在 [128,256,512,1024]×[1e-4,5e-4,1e-3,5e-3] 中调参,本复现固定 batch 256 / lr 1e-3 / factor 8,接近论文常用配置。
4. 评估候选数:论文正文称"随机采样 100 个未交互物品",而论文官方数据文件(ml-1m.test.negative)每用户实际为 1 个真实目标 + 99 个随机负样本(100 个候选);本复现沿用官方文件,与论文报告的数值同协议,可直接比较。
