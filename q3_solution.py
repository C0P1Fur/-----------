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
# 附件和模板从脚本所在目录读取，输出使用带out后缀的文件名
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
# 生成逐日因果预测，每日训练仅使用此前历史且初期用典型日和历史均值启动
pred_load, pred_pv = build_load_pv_predictions(load, pv, init_load, init_pv)
# 同发布版本与同提前量的历史偏差用于收缩校正当前光伏预报
corrected = build_corrected_pv_forecasts(F, pv)
version_resid = build_version_residuals(pred_load, load, pv, corrected)
D, N = load.shape
print(D, N)
# 分别保存原日前计划与最终常规购电量，以便计算净调整费用
q0_all = np.zeros((D, N)); q_all = np.zeros((D, N))
ch_all = np.zeros((D, N)); dis_all = np.zeros((D, N)); em_all = np.zeros((D, N))
E_all = np.zeros((D, N+1)); daily_cost = np.zeros(D)

# ==================== 3. 00时计划 + 06/12/18滚动调整 + 真实执行 ====================
e_start = 6000.0
# 四个执行段从00时与06时及12时和18时开始，每段六小时
starts = [0, 36, 72, 108]
ends = [36, 72, 108, 144]
for d in range(D):
    print(f"正在计算第 {d+1}/{D} 天...", flush=True)
    g0 = pv_forecast_10min(corrected, pv, d, 0)
    q0, _, _, _, _ = solve_nominal(
        pred_load[d] - g0, reserve_version(version_resid, d, 0), price, e_start
    )
    q0_all[d] = q0
    # 复制原计划供后续更新，保留q0作为各次调整及最终结算的基准
    q_day = q0.copy()
    e_cur = e_start
    E_all[d, 0] = e_start

    for k, (s, e) in enumerate(zip(starts, ends)):
        if k > 0:
            g = pv_forecast_10min(corrected, pv, d, k)
            lp = load_feedback(pred_load, load, d, s)
            # 仅重算尚未执行的时段，并以当前真实储电量作为新的规划初值
            q_day[s:] = solve_adjust(
                q0[s:], lp - g, reserve_version(version_resid, d, k), price[s:], e_cur
            )
        # 执行当前六小时段时先充放电并对剩余缺口应急补购，下一段承接末端状态
        ch, dis, em, _, E = execute_segment(q_day[s:e], load[d, s:e], pv[d, s:e], e_cur)
        ch_all[d, s:e] = ch; dis_all[d, s:e] = dis; em_all[d, s:e] = em
        # 状态包含分段首尾边界，因此其切片比电量切片多一个元素
        E_all[d, s:e+1] = E
        e_cur = E[-1]

    q_all[d] = q_day
    # 增减量均相对原日前计划取正部分，最终结算不重复累计中间版本交易
    inc = np.maximum(q_day - q0, 0.0)
    dec = np.maximum(q0 - q_day, 0.0)
    # 按原计划费加1.5倍增购费减0.5倍减购净返还再加五倍应急费结算
    daily_cost[d] = float(
        (price*q0).sum() + 1.5*(price*inc).sum() - 0.5*(price*dec).sum()
        + 5.0*(price*em_all[d]).sum()
    )
    # 将最后一个执行分段的实际储电量传给下一天
    e_start = e_cur
    print(f"当前储电量：",e_start)

# ==================== 4. 结果写入 ====================
# 把相同分时电价复制到每个日期以满足统一导出函数的二维输入
settlement_price = np.tile(price, (D, 1))
write_result3(str(TEMPLATE), str(OUTPUT), q0_all, q_all, ch_all, dis_all,
              em_all, E_all, daily_cost, settlement_price)
# 只汇总二月一日至十二月三十一日，1月用于启动预测和状态预热
sl = slice(31, None)
print(f"最终常规购电量: {q_all[sl].sum():.2f} kWh")
print(f"紧急购电量: {em_all[sl].sum():.2f} kWh")
print(f"总费用: {daily_cost[sl].sum():.2f} 元")
emergency_cost = float(5.0 * (settlement_price[sl] * em_all[sl]).sum())
print(f"紧急购电总费用: {emergency_cost:.2f} 元")
print(f"已按模板写出: {OUTPUT}")
