# -*- coding: utf-8 -*-
"""C题问题1：确定性典型日 MILP（Excel 使用 openpyxl）。"""
from pathlib import Path
import numpy as np
from microgrid_core import read_attachment1, solve_nominal, write_result1

# ==================== 1. 数据读取 ====================
BASE = Path(__file__).resolve().parent
ATT1 = BASE / "附件1.xlsx"
TEMPLATE = BASE / "result1.xlsx"
OUTPUT = BASE / "result1out.xlsx"
price, load, pv = read_attachment1(str(ATT1))

# ==================== 2. 模型求解 ====================
# 典型日无预测误差，备用裕量为0；0:00和24:00储电量均固定6000 kWh。
q, charge, discharge, E, objective = solve_nominal(
    load - pv, np.zeros(144), price, e_start=6000.0, e_terminal=6000.0
)

# ==================== 3. 结果输出 ====================
write_result1(str(TEMPLATE), str(OUTPUT), q, charge, discharge, E)
print(f"全天购电量: {q.sum():.2f} kWh")
print(f"全天购电费: {objective:.2f} 元")
print(f"0:00 / 24:00 储电量: {E[0]:.2f} / {E[-1]:.2f} kWh")
print(f"已按模板写出: {OUTPUT}")
