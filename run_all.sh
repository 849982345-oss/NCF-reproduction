#!/bin/bash
# Full reproduction run: GMF / MLP / NeuMF-end on MovieLens-1M, factor=8
cd /c/research-training/repro-ncf/ncf_pytorch || exit 1
PY=/c/Users/Administrator/AppData/Local/Programs/Python/Python313/python.exe
mkdir -p ../console_logs
for m in GMF MLP NeuMF-end; do
  echo "===== $(date '+%H:%M:%S') START $m ====="
  $PY main.py --model "$m" --epochs 50 --factor_num 8 --batch_size 256 --lr 0.001 --gpu 0 \
    > "../console_logs/${m}.log" 2>&1
  echo "===== $(date '+%H:%M:%S') END $m (exit $?) ====="
done
echo "ALL_DONE $(date '+%H:%M:%S')"
