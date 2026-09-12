# -*- coding: utf-8 -*-
"""C题问题3：附件3光伏预报校正 + 06/12/18日内重调度（Excel 使用 openpyxl）。"""
from pathlib import Path
import numpy as np
from microgrid_core import (
    read_attachment1, read_attachment2, read_attachment3,
    build_load_pv_predictions, build_corrected_pv_forecasts,
    build_version_residuals, pv_forecast_10min, reserve_version,
    load_feedback, solve_nominal, solve_adjust, execute_segment, write_result3
)

# ==================== 1. 数据读取 ====================
BASE = Path(__file__).resolve().parent
ATT1 = BASE / "附件1.xlsx"
ATT2 = BASE / "附件2.xlsx"
ATT3 = BASE / "附件3.xlsx"
TEMPLATE = BASE / "result3.xlsx"
OUTPUT = BASE / "result3out.xlsx"
price, init_load, init_pv = read_attachment1(str(ATT1))
load, pv = read_attachment2(str(ATT2))
F = read_attachment3(str(ATT3))

# ==================== 2. 参数初始化 / 预测校正 ====================
pred_load, pred_pv = build_load_pv_predictions(load, pv, init_load, init_pv)
corrected = build_corrected_pv_forecasts(F, pv)
version_resid = build_version_residuals(pred_load, load, pv, corrected)
D, N = load.shape
q0_all = np.zeros((D, N)); q_all = np.zeros((D, N))
ch_all = np.zeros((D, N)); dis_all = np.zeros((D, N)); em_all = np.zeros((D, N))
E_all = np.zeros((D, N+1)); daily_cost = np.zeros(D)

# ==================== 3. 00时计划 + 06/12/18滚动调整 + 真实执行 ====================
e_start = 6000.0
starts = [0, 36, 72, 108]
ends = [36, 72, 108, 144]
for d in range(D):
    print(f"正在计算第 {d+1}/{D} 天...", flush=True)
    g0 = pv_forecast_10min(corrected, pv, d, 0)
    q0, _, _, _, _ = solve_nominal(
        pred_load[d] - g0, reserve_version(version_resid, d, 0), price, e_start
    )
    q0_all[d] = q0
    q_day = q0.copy()
    e_cur = e_start
    E_all[d, 0] = e_start

    for k, (s, e) in enumerate(zip(starts, ends)):
        if k > 0:
            g = pv_forecast_10min(corrected, pv, d, k)
            lp = load_feedback(pred_load, load, d, s)
            q_day[s:] = solve_adjust(
                q0[s:], lp - g, reserve_version(version_resid, d, k), price[s:], e_cur
            )
        ch, dis, em, _, E = execute_segment(q_day[s:e], load[d, s:e], pv[d, s:e], e_cur)
        ch_all[d, s:e] = ch; dis_all[d, s:e] = dis; em_all[d, s:e] = em
        E_all[d, s:e+1] = E
        e_cur = E[-1]

    q_all[d] = q_day
    inc = np.maximum(q_day - q0, 0.0)
    dec = np.maximum(q0 - q_day, 0.0)
    daily_cost[d] = float(
        (price*q0).sum() + 1.5*(price*inc).sum() - 0.5*(price*dec).sum()
        + 5.0*(price*em_all[d]).sum()
    )
    e_start = e_cur

# ==================== 4. 结果写入 ====================
settlement_price = np.tile(price, (D, 1))
write_result3(str(TEMPLATE), str(OUTPUT), q0_all, q_all, ch_all, dis_all,
              em_all, E_all, daily_cost, settlement_price)
sl = slice(31, None)
print(f"最终常规购电量: {q_all[sl].sum():.2f} kWh")
print(f"紧急购电量: {em_all[sl].sum():.2f} kWh")
print(f"总费用: {daily_cost[sl].sum():.2f} 元")
print(f"已按模板写出: {OUTPUT}")
