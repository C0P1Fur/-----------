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
# 附件和模板从脚本所在目录读取，输出使用带out后缀的文件名
BASE = Path(__file__).resolve().parent
ATT1 = BASE / "附件1.xlsx"
ATT2 = BASE / "附件2.xlsx"
TEMPLATE = BASE / "result2.xlsx"
OUTPUT = BASE / "result2out.xlsx"
price, init_load, init_pv = read_attachment1(str(ATT1))
load, pv = read_attachment2(str(ATT2))

# ==================== 2. 参数初始化 / 因果预测 ====================
print(f"参数初始化", flush=True)
# 生成逐日因果预测，每日训练仅使用此前历史且初期用典型日和历史均值启动
pred_load, pred_pv = build_load_pv_predictions(load, pv, init_load, init_pv)
# 净负荷残差为实际减预测，之后只从当前日期之前的残差配置备用
residual = (load - pv) - (pred_load - pred_pv)
D, N = load.shape
# 用日期和时段两维保存计划及实际执行量，储能状态另需N加1个边界
q_all = np.zeros((D, N)); ch_all = np.zeros((D, N)); dis_all = np.zeros((D, N))
em_all = np.zeros((D, N)); E_all = np.zeros((D, N+1)); daily_cost = np.zeros(D)

# ==================== 3. 模型求解 + 逐时执行 ====================
e_start = 6000.0
for d in range(D):
    print(f"正在计算第 {d+1}/{D} 天...", flush=True)
    # 使用固定0.5分位数配置非负裕量，此处未执行候选参数择优
    reserve = reserve_from_residuals(residual, d, 0.5)
    # 只保留名义购电计划，实际充放电由下一步真实供需执行重新确定
    q, _, _, _, _ = solve_nominal(pred_load[d] - pred_pv[d], reserve, price, e_start)
    ch, dis, em, _, E = execute_segment(q, load[d], pv[d], e_start)
    q_all[d], ch_all[d], dis_all[d], em_all[d], E_all[d] = q, ch, dis, em, E
    # 常规计划按原价结算，应急补购按五倍价结算且未使用计划电量仍收费
    daily_cost[d] = float((price*q).sum() + 5.0*(price*em).sum())
    # 将当天真实日末储电量承接到下一天，不重新设为6000千瓦时
    e_start = E[-1]

# ==================== 4. 结果写入 ====================
write_result2(str(TEMPLATE), str(OUTPUT), q_all, ch_all, dis_all, em_all, E_all, daily_cost)
# 只汇总二月一日至十二月三十一日，1月用于启动预测和状态预热
sl = slice(31, None)  # 2月1日至12月31日
print(f"常规购电量: {q_all[sl].sum():.2f} kWh")
print(f"紧急购电量: {em_all[sl].sum():.2f} kWh")
print(f"总费用: {daily_cost[sl].sum():.2f} 元")
print(f"已按模板写出: {OUTPUT}")
