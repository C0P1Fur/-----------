# -*- coding: utf-8 -*-
"""微网与外部电网电力调控策略——公共函数"""
from __future__ import annotations

import math
from datetime import date, datetime
from pathlib import Path
from typing import Any

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel
from scipy.optimize import milp, LinearConstraint, Bounds
from scipy.sparse import lil_matrix

DT = 1.0 / 6.0
ETA_C = 0.9
ETA_D = 0.9
E_MIN = 1200.0
E_MAX = 10800.0
P_MAX_E = 5000.0 * DT  # 每10分钟最大充/放电量 kWh
RIDGE_LAMBDA = 10.0
N_SLOT = 144
DAYS = 365
BASE_DATE = date(2025, 1, 1)
RELEASE_HOURS = np.array([0, 6, 12, 18], dtype=int)


# ============================================================================
# Excel 数据读取
# ============================================================================
# 读取附件1 返回price, load, pv
def read_attachment1(path: str):
    print(f"[读取] 正在读取附件1：{path}", flush=True)

    wb = load_workbook(
        path,
        read_only=True,
        data_only=True,
        keep_links=False
    )

    ws = wb["Sheet1"]

    data = []
    for row in ws.iter_rows(
        min_row=2,
        max_row=145,
        min_col=2,
        max_col=4,
        values_only=True
    ):
        data.append(row)

    wb.close()

    arr = np.asarray(data, dtype=float)

    price = arr[:, 0]
    load = arr[:, 1]
    pv = arr[:, 2]

    print("[读取] 附件1读取完成", flush=True)

    return price, load, pv


# 读取附件2 返回load(act), pv(act)
def read_attachment2(path: str):

    print(f"[读取] 正在打开附件2：{path}", flush=True)

    wb = load_workbook(
        path,
        read_only=True,
        data_only=True,
        keep_links=False
    )

    print("[读取] Excel 已打开，正在读取小区负载...", flush=True)

    ws_load = wb["小区负载"]

    load_data = []

    for row in ws_load.iter_rows(
        min_row=2,
        max_row=366,
        min_col=2,
        max_col=145,
        values_only=True
    ):
        load_data.append(row)

    print("[读取] 小区负载读取完成，正在读取光伏...", flush=True)

    ws_pv = wb["光伏发电实际功率"]

    pv_data = []

    for row in ws_pv.iter_rows(
        min_row=2,
        max_row=366,
        min_col=2,
        max_col=145,
        values_only=True
    ):
        pv_data.append(row)

    wb.close()

    print("[读取] Excel 文件已关闭，正在转换 NumPy 数组...", flush=True)

    load = np.asarray(load_data, dtype=float)
    pv = np.asarray(pv_data, dtype=float)

    print(
        f"[读取] 附件2读取完成："
        f"load.shape={load.shape}, "
        f"pv.shape={pv.shape}",
        flush=True
    )

    return load, pv


"""
读取附件3
返回：
    forecast: ndarray, shape = (365, 4, 24)

    forecast[d, k, h]
    d = 0~364，对应 2025-01-01 ~ 2025-12-31
    k = 0,1,2,3，对应 0:00、6:00、12:00、18:00
    h = 0~23，对应未来1~24小时
"""
def read_attachment3(path: str):

    print(f"[读取] 正在读取附件3：{path}", flush=True)

    wb = load_workbook(
        path,
        read_only=True,
        data_only=True,
        keep_links=False
    )

    ws = wb["Sheet1"]

    # 365天 × 每天4次发布 × 未来24小时
    forecast = np.zeros((365, 4, 24), dtype=float)

    row_index = 0

    for row in ws.iter_rows(
        min_row=2,
        max_row=1461,
        min_col=3,
        max_col=26,
        values_only=True
    ):
        day = row_index // 4
        release = row_index % 4

        # None 安全处理
        values = [
            0.0 if v is None else float(v)
            for v in row
        ]

        forecast[day, release, :] = values

        row_index += 1

    wb.close()

    if row_index != 1460:
        raise ValueError(
            f"附件3读取行数异常：实际 {row_index} 行，预期 1460 行"
        )

    print(
        f"[读取] 附件3完成，forecast.shape={forecast.shape}",
        flush=True
    )

    return forecast

"""
读取附件4
返回：
    price[d, t]
    d = 0~364
    t = 0~143
"""
def read_attachment4(path: str):

    print(f"[读取] 正在读取附件4：{path}", flush=True)

    wb = load_workbook(
        path,
        read_only=True,
        data_only=True,
        keep_links=False
    )

    ws = wb["Sheet1"]

    price_data = []

    for row in ws.iter_rows(
        min_row=2,
        max_row=366,
        min_col=2,
        max_col=145,
        values_only=True
    ):
        values = [
            np.nan if v is None else float(v)
            for v in row
        ]

        price_data.append(values)

    wb.close()

    price = np.asarray(price_data, dtype=float)

    if price.shape != (365, 144):
        raise ValueError(
            f"附件4维度异常：实际 {price.shape}，预期 (365, 144)"
        )

    if np.isnan(price).any():
        raise ValueError("附件4中存在空白电价数据")

    print(
        f"[读取] 附件4完成，price.shape={price.shape}",
        flush=True
    )

    return price

# ============================================================================
# MILP 优化与真实执行
# ============================================================================
"""决策变量顺序[q,c,d,E,y] 返回计划购电量"""
def solve_nominal(net_pred_kw: np.ndarray, reserve_kw: np.ndarray,
                  price: np.ndarray, e_start: float, e_terminal: float = 6000.0):
    n = len(net_pred_kw) # 获得单决策变量数量
    nv = 5*n + 1 # 总变量的数量
    iq, ic, idd, iE, iy = 0, n, 2*n, 3*n, 4*n + 1 # 购电量 充电量 放电量 储电量 状态变量在数组中的位置

    # 输出
    obj = np.zeros(nv)
    obj[:n] = price
    lb = np.zeros(nv)
    ub = np.full(nv, np.inf)
    ub[ic:ic+n] = P_MAX_E
    ub[idd:idd+n] = P_MAX_E
    lb[iE:iE+n+1] = E_MIN
    ub[iE:iE+n+1] = E_MAX
    ub[iy:iy+n] = 1.0
    integrality = np.zeros(nv)
    integrality[iy:iy+n] = 1

    A = lil_matrix((4*n + 2, nv))
    lo = np.full(4*n + 2, -np.inf)
    hi = np.full(4*n + 2, np.inf)
    r = 0
    for t in range(n):
        A[r, iE+t+1] = 1
        A[r, iE+t] = -1
        A[r, ic+t] = -ETA_C
        A[r, idd+t] = 1 / ETA_D
        lo[r] = hi[r] = 0
        r += 1
    A[r, iE] = 1; lo[r] = hi[r] = e_start; r += 1
    A[r, iE+n] = 1; lo[r] = hi[r] = e_terminal; r += 1

    req = (net_pred_kw + reserve_kw) * DT
    for t in range(n):
        A[r, t] = 1
        A[r, ic+t] = -1
        A[r, idd+t] = 1
        lo[r] = req[t]
        r += 1
    for t in range(n):
        A[r, ic+t] = 1; A[r, iy+t] = -P_MAX_E; hi[r] = 0; r += 1
        A[r, idd+t] = 1; A[r, iy+t] = P_MAX_E; hi[r] = P_MAX_E; r += 1

    res = milp(
        obj,
        integrality=integrality,
        bounds=Bounds(lb, ub),
        constraints=LinearConstraint(A.tocsr(), lo, hi),
        options={"time_limit": 15, "mip_rel_gap": 1e-7},
    )
    if not res.success:
        raise RuntimeError(f"MILP求解失败: {res.message}")
    x = res.x
    return x[:n], x[ic:ic+n], x[idd:idd+n], x[iE:iE+n+1], float(res.fun)

"""考虑计划购电量和实际购电量的MILP"""
def solve_adjust(q0: np.ndarray, net_pred_kw: np.ndarray, reserve_kw: np.ndarray,
                 price: np.ndarray, e_start: float):
    n = len(q0)
    iq, iu, iv, ic, idd, iE, iy = 0, n, 2*n, 3*n, 4*n, 5*n, 6*n + 1
    nv = 7*n + 1
    obj = np.zeros(nv)
    obj[iu:iu+n] = 1.5 * price
    obj[iv:iv+n] = -0.5 * price
    lb = np.zeros(nv)
    ub = np.full(nv, np.inf)
    ub[iv:iv+n] = q0
    ub[ic:ic+n] = P_MAX_E
    ub[idd:idd+n] = P_MAX_E
    lb[iE:iE+n+1] = E_MIN
    ub[iE:iE+n+1] = E_MAX
    ub[iy:iy+n] = 1.0
    integrality = np.zeros(nv)
    integrality[iy:iy+n] = 1

    A = lil_matrix((5*n + 2, nv))
    lo = np.full(5*n + 2, -np.inf)
    hi = np.full(5*n + 2, np.inf)
    r = 0
    for t in range(n):
        A[r, iq+t] = 1; A[r, iu+t] = -1; A[r, iv+t] = 1
        lo[r] = hi[r] = q0[t]
        r += 1
    for t in range(n):
        A[r, iE+t+1] = 1; A[r, iE+t] = -1
        A[r, ic+t] = -ETA_C; A[r, idd+t] = 1/ETA_D
        lo[r] = hi[r] = 0
        r += 1
    A[r, iE] = 1; lo[r] = hi[r] = e_start; r += 1
    A[r, iE+n] = 1; lo[r] = hi[r] = 6000.0; r += 1
    req = (net_pred_kw + reserve_kw) * DT
    for t in range(n):
        A[r, iq+t] = 1; A[r, ic+t] = -1; A[r, idd+t] = 1
        lo[r] = req[t]
        r += 1
    for t in range(n):
        A[r, ic+t] = 1; A[r, iy+t] = -P_MAX_E; hi[r] = 0; r += 1
        A[r, idd+t] = 1; A[r, iy+t] = P_MAX_E; hi[r] = P_MAX_E; r += 1

    res = milp(
        obj,
        integrality=integrality,
        bounds=Bounds(lb, ub),
        constraints=LinearConstraint(A.tocsr(), lo, hi),
        options={"time_limit": 15, "mip_rel_gap": 1e-7},
    )
    if not res.success:
        raise RuntimeError(f"调整MILP求解失败: {res.message}")
    return res.x[:n]


def execute_segment(q: np.ndarray, load: np.ndarray, pv: np.ndarray, e_start: float):
    """真实执行：盈余先充电，缺口先放电，仍不足则紧急购电。"""
    n = len(q)
    ch = np.zeros(n); dis = np.zeros(n); emergency = np.zeros(n); waste = np.zeros(n)
    E = np.zeros(n + 1); E[0] = e_start
    for t in range(n):
        balance = q[t] + (pv[t] - load[t]) * DT
        if balance >= 0:
            ch[t] = max(0.0, min(balance, P_MAX_E, (E_MAX - E[t]) / ETA_C))
            waste[t] = max(0.0, balance - ch[t])
        else:
            dis[t] = max(0.0, min(-balance, P_MAX_E, (E[t] - E_MIN) * ETA_D))
            emergency[t] = max(0.0, -balance - dis[t])
        E[t+1] = E[t] + ETA_C * ch[t] - dis[t] / ETA_D
    return ch, dis, emergency, waste, E

# ============================================================================
# 预测模型
# ============================================================================
"""滚动岭回归 12维特征 60天窗口。"""
def build_load_pv_predictions(load: np.ndarray, pv: np.ndarray,
                              init_load: np.ndarray, init_pv: np.ndarray):
    D, N = load.shape
    feat = np.zeros((D, N, 12), dtype=float)
    t = np.arange(N)
    theta = 2 * np.pi * t / N
    for d in range(7, D):
        feat[d, :, 0] = load[d - 1]
        feat[d, :, 1] = pv[d - 1]
        feat[d, :, 2] = load[d - 7]
        feat[d, :, 3] = pv[d - 7]
        feat[d, :, 4] = load[d - 7:d].mean(axis=0)
        feat[d, :, 5] = pv[d - 7:d].mean(axis=0)
        feat[d, :, 6] = np.sin(theta)
        feat[d, :, 7] = np.cos(theta)
        feat[d, :, 8] = np.sin(2 * theta)
        feat[d, :, 9] = np.cos(2 * theta)
        feat[d, :, 10] = math.sin(2 * np.pi * (d % 7) / 7)
        feat[d, :, 11] = math.cos(2 * np.pi * (d % 7) / 7)

    pred_l = np.zeros_like(load)
    pred_g = np.zeros_like(pv)
    pred_l[0] = init_load
    pred_g[0] = init_pv

    for d in range(1, 15):
        s = max(0, d - 7)
        pred_l[d] = load[s:d].mean(axis=0)
        pred_g[d] = pv[s:d].mean(axis=0)

    for d in range(15, D):
        s = max(7, d - 60)
        X = feat[s:d].reshape(-1, 12)
        Y = np.column_stack([load[s:d].reshape(-1), pv[s:d].reshape(-1)])
        mu = X.mean(axis=0)
        sd = X.std(axis=0)
        sd[sd == 0] = 1.0
        Z = (X - mu) / sd
        ym = Y.mean(axis=0)
        beta = np.linalg.solve(
            Z.T @ Z + RIDGE_LAMBDA * np.eye(12),
            Z.T @ (Y - ym),
        )
        P = ym + ((feat[d] - mu) / sd) @ beta
        pred_l[d] = np.maximum(P[:, 0], 0.0)
        pred_g[d] = np.maximum(P[:, 1], 0.0)
        pred_g[d][np.all(pv[d - 7:d] == 0, axis=0)] = 0.0
    return pred_l, pred_g

"""滚动岭回归电价预测 9维"""
def build_price_predictions(price: np.ndarray, init_price: np.ndarray):
    D, N = price.shape
    feat = np.zeros((D, N, 9), dtype=float)
    t = np.arange(N)
    theta = 2 * np.pi * t / N
    for d in range(7, D):
        feat[d, :, 0] = price[d - 1]
        feat[d, :, 1] = price[d - 7]
        feat[d, :, 2] = price[d - 7:d].mean(axis=0)
        feat[d, :, 3] = np.sin(theta)
        feat[d, :, 4] = np.cos(theta)
        feat[d, :, 5] = np.sin(2 * theta)
        feat[d, :, 6] = np.cos(2 * theta)
        feat[d, :, 7] = math.sin(2 * np.pi * (d % 7) / 7)
        feat[d, :, 8] = math.cos(2 * np.pi * (d % 7) / 7)

    pred = np.zeros_like(price)
    pred[0] = init_price
    for d in range(1, 15):
        s = max(0, d - 7)
        pred[d] = price[s:d].mean(axis=0)
    for d in range(15, D):
        s = max(7, d - 60)
        X = feat[s:d].reshape(-1, 9)
        y = price[s:d].reshape(-1)
        mu = X.mean(axis=0)
        sd = X.std(axis=0)
        sd[sd == 0] = 1.0
        Z = (X - mu) / sd
        ym = y.mean()
        beta = np.linalg.solve(
            Z.T @ Z + RIDGE_LAMBDA * np.eye(9), Z.T @ (y - ym)
        )
        pred[d] = np.maximum(ym + ((feat[d] - mu) / sd) @ beta, 0.001)
    return pred

'''裕量计算'''
def reserve_from_residuals(residuals: np.ndarray, day: int, quantile: float = 0.5):
    if day == 0:
        return np.zeros(residuals.shape[1], dtype=float)
    hist = residuals[max(0, day - 28):day]
    return np.maximum(np.quantile(hist, quantile, axis=0), 0.0)

# ============================================================================
# 问题3/4的日内光伏预测与反馈
# ============================================================================
def build_corrected_pv_forecasts(F: np.ndarray, pv_actual: np.ndarray):
    """附件3：同版本/同提前量28天历史偏差 + n/(n+7)收缩校正。"""
    D = F.shape[0]
    ferr = np.full_like(F, np.nan, dtype=float)

    def actual_at_hour(day: int, hour_abs: int):
        total = day * 24 + hour_abs
        target_day = (total - 1) // 24
        local_hour = total - target_day * 24
        idx = int(local_hour * 6 - 1)
        if target_day < 0 or target_day >= D:
            return np.nan
        return pv_actual[target_day, idx]

    for d in range(D):
        for k, rh in enumerate(RELEASE_HOURS):
            for h in range(1, 25):
                a = actual_at_hour(d, int(rh + h))
                if np.isfinite(a):
                    ferr[d, k, h-1] = F[d, k, h-1] - a

    corrected = np.zeros_like(F)
    for d in range(D):
        for k in range(4):
            for h in range(24):
                e = ferr[max(0, d-28):d, k, h]
                e = e[np.isfinite(e)]
                n = len(e)
                bias = float(e.mean()) if n else 0.0
                shrink = n / (n + 7.0) if n else 0.0
                corrected[d, k, h] = max(F[d, k, h] - shrink * bias, 0.0)
    return corrected


def pv_forecast_10min(corrected: np.ndarray, pv_actual: np.ndarray, day: int, version: int):
    s = int(RELEASE_HOURS[version] * 6)
    if day == 0 and s == 0:
        anchor = 0.0
    elif s == 0:
        anchor = pv_actual[day-1, 143]
    else:
        anchor = pv_actual[day, s-1]
    knots = np.r_[anchor, corrected[day, version]]
    out = np.zeros(N_SLOT - s)
    for j, t in enumerate(range(s, N_SLOT)):
        minutes = (t - s + 1) * 10
        h0 = minutes // 60
        if minutes % 60 == 0:
            out[j] = knots[h0]
        else:
            f = (minutes % 60) / 60.0
            out[j] = (1-f) * knots[h0] + f * knots[h0+1]
    return out


def build_version_residuals(pred_load: np.ndarray, load: np.ndarray, pv: np.ndarray,
                            corrected: np.ndarray):
    out = np.full((4, DAYS, N_SLOT), np.nan)
    for k, rh in enumerate(RELEASE_HOURS):
        s = int(rh * 6)
        for d in range(DAYS):
            g = pv_forecast_10min(corrected, pv, d, k)
            out[k, d, s:] = (load[d, s:] - pv[d, s:]) - (pred_load[d, s:] - g)
    return out


def reserve_version(version_residuals: np.ndarray, day: int, version: int):
    s = int(RELEASE_HOURS[version] * 6)
    if day == 0:
        return np.zeros(N_SLOT - s)
    hist = version_residuals[version, max(0, day-28):day, s:]
    return np.maximum(np.quantile(hist, 0.5, axis=0), 0.0)


def load_feedback(pred_load: np.ndarray, load: np.ndarray, day: int, start: int):
    if start == 0:
        return pred_load[day].copy()
    bias = float(np.mean(load[day, start-12:start] - pred_load[day, start-12:start]))
    t = np.arange(start, N_SLOT)
    return np.maximum(pred_load[day, start:] + bias * np.exp(-(t-start)/36.0), 0.0)


def price_feedback(pred_price: np.ndarray, actual_price: np.ndarray, day: int, start: int):
    if start == 0:
        return pred_price[day].copy()
    bias = float(np.mean(actual_price[day, start-12:start] - pred_price[day, start-12:start]))
    t = np.arange(start, N_SLOT)
    return np.maximum(pred_price[day, start:] + bias * np.exp(-(t-start)/36.0), 0.001)


# ============================================================================
# result 模板写入
# ============================================================================
def positive_runs(arr: np.ndarray, eps: float = 1e-7):
    runs = []
    i = 0
    while i < len(arr):
        if arr[i] > eps:
            j = i
            total = 0.0
            while j < len(arr) and arr[j] > eps:
                total += float(arr[j])
                j += 1
            runs.append((i, j, total))
            i = j
        else:
            i += 1
    return runs


def _clock(slot_boundary: int):
    """10分钟边界序号转为 HH:MM；144 显示为 24:00。"""
    m = slot_boundary * 10
    if m == 24 * 60:
        return "24:00"
    return f"{m // 60}:{m % 60:02d}"


def _excel_date_to_day_index(value: Any, epoch=None):
    """把模板日期单元格转换为相对 2025-01-01 的日序号；非日期返回 None。"""
    if value is None or value == "⁝":
        return None
    if isinstance(value, datetime):
        return (value.date() - BASE_DATE).days
    if isinstance(value, date):
        return (value - BASE_DATE).days
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            d = from_excel(value, epoch=epoch)
            if isinstance(d, datetime):
                d = d.date()
            return (d - BASE_DATE).days
        except Exception:
            return int(round(float(value) - 45658.0))
    if isinstance(value, str):
        text = value.strip().replace(".", "-").replace("/", "-")
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
            try:
                return (datetime.strptime(text, fmt).date() - BASE_DATE).days
            except ValueError:
                pass
    return None


def _fill_charge_sheet(wb, ch: np.ndarray, dis: np.ndarray, E: np.ndarray):
    ws = wb["充放电量"]
    blocks = {
        "0:00-4:00": 0,
        "4:00-8:00": 24,
        "8:00-12:00": 48,
        "12:00-16:00": 72,
        "16:00-20:00": 96,
        "20:00-24:00": 120,
    }
    current_day = None
    # 按模板现有行扫描；只填模板已经列出的指定日期，保留原格式。
    for r in range(2, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a == "⁝":
            current_day = None
        else:
            idx = _excel_date_to_day_index(a, wb.epoch)
            if idx is not None:
                current_day = idx
        if current_day is None or not (0 <= current_day < ch.shape[0]):
            continue

        period = ws.cell(r, 2).value
        if period in blocks:
            s = blocks[period]
            ws.cell(r, 3).value = float(ch[current_day, s:s+24].sum())
            ws.cell(r, 4).value = float(dis[current_day, s:s+24].sum())

        marker = ws.cell(r, 5).value
        marker_text = str(marker).strip() if marker is not None else ""
        if marker == 0 or marker_text == "0:00":
            ws.cell(r, 6).value = float(E[current_day, 0])
        elif marker_text == "24:00":
            ws.cell(r, 6).value = float(E[current_day, -1])


def _fill_emergency_sheet(wb, emergency: np.ndarray):
    ws = wb["紧急购电量"]
    r = 2
    max_row = ws.max_row
    while r <= max_row:
        day = _excel_date_to_day_index(ws.cell(r, 1).value, wb.epoch)
        if day is None or not (0 <= day < emergency.shape[0]):
            r += 1
            continue

        # 当前日期在模板中占用到下一条日期之前的所有空白行。
        rr = r + 1
        while rr <= max_row and ws.cell(rr, 1).value is None:
            rr += 1
        capacity = rr - r

        # 先清理模板中的旧示例值，避免残留。
        for row in range(r, r + capacity):
            ws.cell(row, 2).value = None
            ws.cell(row, 3).value = None

        runs = positive_runs(emergency[day])
        for j, (s, e, amount) in enumerate(runs[:capacity]):
            ws.cell(r + j, 2).value = f"{_clock(s)}-{_clock(e)}"
            ws.cell(r + j, 3).value = float(amount)
        r = rr


def _write_matrix(ws, start_row: int, start_col: int, matrix):
    for i, row in enumerate(matrix, start=start_row):
        for j, value in enumerate(row, start=start_col):
            ws.cell(i, j).value = value


def write_result1(template_path: str | Path, output_path: str | Path,
                  q, ch, dis, E):
    wb = load_workbook(template_path)
    ws_q = wb["计划购电量"]
    for i, value in enumerate(q, start=2):
        ws_q.cell(i, 2).value = float(value)

    ws_cd = wb["充放电量"]
    for b in range(6):
        s = 24 * b
        ws_cd.cell(2 + b, 2).value = float(ch[s:s+24].sum())
        ws_cd.cell(2 + b, 3).value = float(dis[s:s+24].sum())
    ws_cd.cell(2, 5).value = float(E[0])
    ws_cd.cell(3, 5).value = float(E[-1])

    wb.save(output_path)
    wb.close()


def write_result2(template_path: str | Path, output_path: str | Path,
                  q, ch, dis, emergency, E, daily_cost):
    wb = load_workbook(template_path)
    ws = wb["计划购电量"]
    rows = [
        [float(x) for x in q[d]] + [float(q[d].sum()), float(daily_cost[d])]
        for d in range(31, 365)
    ]
    _write_matrix(ws, 2, 2, rows)
    _fill_charge_sheet(wb, ch, dis, E)
    _fill_emergency_sheet(wb, emergency)
    wb.save(output_path)
    wb.close()


def write_result3(template_path: str | Path, output_path: str | Path,
                  q0, q, ch, dis, emergency, E, daily_cost, settlement_price):
    wb = load_workbook(template_path)
    plan_rows = []
    adjust_rows = []
    for d in range(31, 365):
        plan_rows.append(
            [float(x) for x in q0[d]]
            + [float(q0[d].sum()), float((settlement_price[d] * q0[d]).sum())]
        )
        adjust_rows.append(
            [float(x) for x in q[d]]
            + [float(q[d].sum()), float(daily_cost[d])]
        )

    _write_matrix(wb["计划购电量"], 2, 2, plan_rows)
    _write_matrix(wb["调整购电量"], 2, 2, adjust_rows)
    _fill_charge_sheet(wb, ch, dis, E)
    _fill_emergency_sheet(wb, emergency)
    wb.save(output_path)
    wb.close()
