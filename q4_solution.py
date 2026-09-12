# -*- coding: utf-8 -*-
"""C题问题4：波动电价下分别重算问题2（4-2）和问题3（4-3）（Excel 使用 openpyxl）。"""
from pathlib import Path
import numpy as np
from microgrid_core import (
    read_attachment1, read_attachment2, read_attachment3, read_attachment4,
    build_load_pv_predictions, build_price_predictions, build_corrected_pv_forecasts,
    build_version_residuals, reserve_from_residuals, reserve_version,
    pv_forecast_10min, load_feedback, price_feedback,
    solve_nominal, solve_adjust, execute_segment, write_result2, write_result3
)

# ==================== 1. 数据读取 ====================
# 附件和模板从脚本所在目录读取，输出使用带out后缀的文件名
BASE = Path(__file__).resolve().parent
ATT1 = BASE / "附件1.xlsx"
ATT2 = BASE / "附件2.xlsx"
ATT3 = BASE / "附件3.xlsx"
ATT4 = BASE / "附件4.xlsx"
TPL42 = BASE / "result4-2.xlsx"
TPL43 = BASE / "result4-3.xlsx"
OUT42 = BASE / "result4-2out.xlsx"
OUT43 = BASE / "result4-3out.xlsx"
fixed_price, init_load, init_pv = read_attachment1(str(ATT1))
load, pv = read_attachment2(str(ATT2))
F = read_attachment3(str(ATT3))
actual_price = read_attachment4(str(ATT4))

# ==================== 2. 公共预测 ====================
# 生成逐日因果预测，每日训练仅使用此前历史且初期用典型日和历史均值启动
pred_load, pred_pv = build_load_pv_predictions(load, pv, init_load, init_pv)
# 未来电价使用历史滚动预测，附件4真实价格只用于历史反馈和事后结算
pred_price = build_price_predictions(actual_price, fixed_price)
# 净负荷残差为实际减预测，之后只从当前日期之前的残差配置备用
residual = (load - pv) - (pred_load - pred_pv)
D, N = load.shape

# ==================== 3A. 问题4-2：波动电价下的日前策略 ====================
# 分别保存问题4日前策略的计划和实际充放电及应急电量
q42 = np.zeros((D, N)); ch42 = np.zeros((D, N)); dis42 = np.zeros((D, N)); em42 = np.zeros((D, N))
E42 = np.zeros((D, N+1)); cost42 = np.zeros(D)
e_start = 6000.0
for d in range(D):
    # 使用固定0.5分位数配置非负裕量，此处未执行候选参数择优
    reserve = reserve_from_residuals(residual, d, 0.5)
    q, _, _, _, _ = solve_nominal(pred_load[d] - pred_pv[d], reserve, pred_price[d], e_start)
    ch, dis, em, _, E = execute_segment(q, load[d], pv[d], e_start)
    q42[d], ch42[d], dis42[d], em42[d], E42[d] = q, ch, dis, em, E
    # 按实际交付时段电价计算问题4日前策略的常规和应急费用
    cost42[d] = float((actual_price[d]*q).sum() + 5.0*(actual_price[d]*em).sum())
    # 将当天真实日末储电量承接到下一天，不重新设为6000千瓦时
    e_start = E[-1]
write_result2(str(TPL42), str(OUT42), q42, ch42, dis42, em42, E42, cost42)

# ==================== 3B. 问题4-3：波动电价 + 日内重调度 ====================
# 同发布版本与同提前量的历史偏差用于收缩校正当前光伏预报
corrected = build_corrected_pv_forecasts(F, pv)
version_resid = build_version_residuals(pred_load, load, pv, corrected)
q0 = np.zeros((D, N)); q43 = np.zeros((D, N)); ch43 = np.zeros((D, N)); dis43 = np.zeros((D, N)); em43 = np.zeros((D, N))
E43 = np.zeros((D, N+1)); cost43 = np.zeros(D)
e_start = 6000.0
# 每段使用左闭右开下标，更新后只执行当前六小时再接收下一版本信息
starts = [0, 36, 72, 108]; ends = [36, 72, 108, 144]
for d in range(D):
    print(f"正在计算第 {d+1}/{D} 天...", flush=True)
    g0 = pv_forecast_10min(corrected, pv, d, 0)
    base, _, _, _, _ = solve_nominal(
        pred_load[d] - g0, reserve_version(version_resid, d, 0), pred_price[d], e_start
    )
    q0[d] = base
    # 保留原日前计划副本并以当天真实储能状态启动日内执行
    q_day = base.copy(); e_cur = e_start; E43[d, 0] = e_start
    for k, (s, e) in enumerate(zip(starts, ends)):
        if k > 0:
            g = pv_forecast_10min(corrected, pv, d, k)
            lp = load_feedback(pred_load, load, d, s)
            pp = price_feedback(pred_price, actual_price, d, s)
            # 剩余购电按更新的负荷与光伏及预测电价优化，增减仍相对原日前base
            q_day[s:] = solve_adjust(base[s:], lp-g, reserve_version(version_resid, d, k), pp, e_cur)
        # 执行当前六小时段时先充放电并对剩余缺口应急补购，下一段承接末端状态
        ch, dis, em, _, E = execute_segment(q_day[s:e], load[d, s:e], pv[d, s:e], e_cur)
        ch43[d, s:e] = ch; dis43[d, s:e] = dis; em43[d, s:e] = em
        E43[d, s:e+1] = E; e_cur = E[-1]
    q43[d] = q_day
    inc = np.maximum(q_day-base, 0.0); dec = np.maximum(base-q_day, 0.0)
    p = actual_price[d]
    # 按真实电价结算原计划与最终计划的净差额，并计入五倍应急费用
    cost43[d] = float((p*base).sum() + 1.5*(p*inc).sum() - 0.5*(p*dec).sum() + 5.0*(p*em43[d]).sum())
    # 将最后一个执行分段的实际储电量传给下一天
    e_start = e_cur
write_result3(str(TPL43), str(OUT43), q0, q43, ch43, dis43, em43, E43, cost43, actual_price)

# ==================== 4. 汇总输出 ====================
# 只汇总二月一日至十二月三十一日，1月用于启动预测和状态预热
sl = slice(31, None)
print(f"问题4-2总费用: {cost42[sl].sum():.2f} 元；紧急购电量: {em42[sl].sum():.2f} kWh")
print(f"问题4-3总费用: {cost43[sl].sum():.2f} 元；紧急购电量: {em43[sl].sum():.2f} kWh")
print(f"已写出: {OUT42}")
print(f"已写出: {OUT43}")
