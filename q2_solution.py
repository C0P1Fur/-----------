# -*- coding: utf-8 -*-
"""C题问题2：历史滚动预测 + 残差分位数备用 + 日前MILP + 实时应急购电"""
from pathlib import Path
import numpy as np
from microgrid_core import (
    read_attachment1, read_attachment2, build_load_pv_predictions,
    reserve_from_residuals, solve_nominal, execute_segment, write_result2
)

# ==================== 1. 数据读取 ====================
print(f"数据读取", flush=True)
BASE = Path(__file__).resolve().parent
ATT1 = BASE / "附件1.xlsx"
ATT2 = BASE / "附件2.xlsx"
TEMPLATE = BASE / "result2.xlsx"
OUTPUT = BASE / "result2out.xlsx"
price, init_load, init_pv = read_attachment1(str(ATT1))
load, pv = read_attachment2(str(ATT2))

# ==================== 2. 参数初始化 / 因果预测 ====================
print(f"参数初始化", flush=True)
pred_load, pred_pv = build_load_pv_predictions(load, pv, init_load, init_pv)
residual = (load - pv) - (pred_load - pred_pv)
D, N = load.shape
q_all = np.zeros((D, N)); ch_all = np.zeros((D, N)); dis_all = np.zeros((D, N))
em_all = np.zeros((D, N)); E_all = np.zeros((D, N+1)); daily_cost = np.zeros(D)

# ==================== 3. 模型求解 + 逐时执行 ====================
e_start = 6000.0
for d in range(D):
    print(f"正在计算第 {d+1}/{D} 天...", flush=True)
    reserve = reserve_from_residuals(residual, d, 0.5)
    q, _, _, _, _ = solve_nominal(pred_load[d] - pred_pv[d], reserve, price, e_start)
    ch, dis, em, _, E = execute_segment(q, load[d], pv[d], e_start)
    q_all[d], ch_all[d], dis_all[d], em_all[d], E_all[d] = q, ch, dis, em, E
    daily_cost[d] = float((price*q).sum() + 5.0*(price*em).sum())
    e_start = E[-1]

# ==================== 4. 结果写入 ====================
write_result2(str(TEMPLATE), str(OUTPUT), q_all, ch_all, dis_all, em_all, E_all, daily_cost)
sl = slice(31, None)  # 2月1日至12月31日
print(f"常规购电量: {q_all[sl].sum():.2f} kWh")
print(f"紧急购电量: {em_all[sl].sum():.2f} kWh")
print(f"总费用: {daily_cost[sl].sum():.2f} 元")
print(f"已按模板写出: {OUTPUT}")
