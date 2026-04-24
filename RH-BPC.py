# -*- coding: utf-8 -*-
import os
import math
import time
import numpy as np
import pandas as pd
import gurobipy as gp
from gurobipy import GRB
from collections import defaultdict
from itertools import combinations


def safe_save_table(df: pd.DataFrame, xlsx_path: str):
    base, _ = os.path.splitext(xlsx_path)
    try:
        import openpyxl  # noqa: F401
        df.to_excel(xlsx_path, index=False)
        print(f"[save] 已保存 Excel: {xlsx_path}")
    except Exception:
        try:
            import xlsxwriter  # noqa: F401
            df.to_excel(xlsx_path, index=False, engine="xlsxwriter")
            print(f"[save] 已保存 Excel(xlsxwriter): {xlsx_path}")
        except Exception:
            csv_path = base + ".csv"
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            print(f"[save] 未检测到 Excel 引擎，已改存 CSV: {csv_path}")


BASE_DIR = r"F:/Onedrive/PlatEMO-2EVRPLDTW/Instance"
Target_DIR = r"D:\OneDrive-CSU\OneDrive - csu.edu.cn\2E-VRP\RH-BCP\Experiment\RH-BCP"
TIME_LIMIT = 3600
MIP_GAP = 0.0
OUTPUT_LOG = 1

# If True: force full-pool exact benchmarking on small instances
EXACT_SOLVE_MODE = True

ROUTE_MAX_STOPS_EXACT = 10 ** 9
ENUM_TIME_CAP = None
ENUM_NODE_CAP = None

PRICING_ROUNDS = 0
PRICING_BEAM_WIDTH = 50
PRICING_MAX_STOPS = 7
PRICING_ADD_PER_S = 10
RC_EPS = 1e-6

MAX_TRUCKS_MULT = 3

USE_DUMMY_COLUMNS = True
DUMMY_COST = 1e6

# In exact benchmarking mode, dummy columns must be disabled; otherwise the model may use
# dummy coverage at a penalty, which is not a true full-model optimum.
if 'EXACT_SOLVE_MODE' in globals() and EXACT_SOLVE_MODE:
    USE_DUMMY_COLUMNS = False


# ===== Stage-0 exactness control =====
# To GUARANTEE stage-0 global optimality under the implemented master model,
# we must:
#   1) enumerate ALL feasible second-echelon route orders for the active stage-0 customers;
#   2) avoid unsafe subset-based route aggregation/dominance pruning, because loading-queue
#      constraints depend on route ordering / first-customer timing;
#   3) solve the final stage-0 MIP to proven optimality (no time limit, zero gap).
STAGE0_REQUIRE_GLOBAL_OPT = True
STAGE0_LP_TIME_LIMIT = None   # None => no LP time limit for stage 0
STAGE0_MIP_TIME_LIMIT = None  # None => no MIP time limit for stage 0

# ========= 成本与物理参数 =========
Q1 = 100.0
Q2 = 10
h1, h2 = 12.5, 3.36
c_v, c_u = 0.1286, 0.001

DR = {'W': 5.0, 'f_s': 10.0, 'E': 550.0, 'n': 4, 'rho': 1.225, 'zeta': 100.0 * math.pi, 'g': 9.81}

# ========= Satellite loading-queue (handling time) =========
# Adds a single loading dock per satellite.
# - Each selected drone sortie must be loaded before takeoff.
# - Loading time is proportional to the number of DELIVERY packages on that sortie.
# - Loading operations are sequential per satellite, but TAKEOFF does NOT block the dock by default.
#   (Dock can keep loading while a drone waits on the ground for takeoff time.)
#
# Queue time variables (hours):
#   tB[s,r]: time when loading of route (s,r) is finished.
#   tD[s,r]: actual takeoff time of route (s,r), with tD >= tB.
#
ENABLE_LOADING_QUEUE = True
LOADING_TIME_PER_DELIV_UNIT_MIN = 1.0  # minutes per delivery package loaded onto the drone
QUEUE_BLOCK_BY_PREV_TAKEOFF = False    # implemented as "False-mode" in this script
M_TIME = 1e4                           # Big-M for time constraints (hours)




def read_instance(csv_path):
    df0 = pd.read_csv(csv_path)
    try:
        _ = df0.iloc[:, 0:2].to_numpy(dtype=float)
        df = df0
    except Exception:
        df = pd.read_csv(csv_path, skiprows=1)

    coords = df.iloc[:, 0:2].to_numpy(dtype=float)
    demands_raw = df.iloc[:, 4].to_numpy(dtype=float)
    attributes = df.iloc[:, 6].astype(str).to_numpy()
    cust_type_raw = df.iloc[:, 9].to_numpy()

    # ---- 节点类型划分：depot / sat / cust ----
    idx_depot = np.where(attributes == "depot")[0].tolist()
    idx_sat = np.where(attributes == "sat")[0].tolist()
    idx_cust = np.where(attributes == "cust")[0].tolist()

    if len(idx_depot) < 1 or len(idx_sat) < 1 or len(idx_cust) < 1:
        raise RuntimeError("实例必须包含 >=1 depot, >=1 sat, >=1 customer")

    coord_dep = coords[idx_depot, :]
    coord_sat = coords[idx_sat, :]
    coord_cus = coords[idx_cust, :]
    all_coords = np.vstack([coord_dep, coord_sat, coord_cus])

    num_dep = len(idx_depot)
    num_sat = len(idx_sat)
    num_cus = len(idx_cust)
    num_nodes = num_dep + num_sat + num_cus

    # ---- 需求方向与量（送货 -1 / 取件 +1）----
    demand_cus = (demands_raw[idx_cust] / 10.0).astype(float)

    def _to_type(x):
        try:
            v = float(x)
            if v == -1: return -1
            if v == 1:  return 1
        except Exception:
            s = str(x).strip().lower()
            if s in ["delivery", "deliver", "d", "-1"]: return -1
            if s in ["pickup", "pick", "p", "+1", "1"]:  return 1
        return -1

    cust_type = np.array([_to_type(x) for x in cust_type_raw[idx_cust]], dtype=int)
    Dd = np.where(cust_type == -1, np.rint(demand_cus), 0.0).astype(int)
    Dp = np.where(cust_type == 1, np.rint(demand_cus), 0.0).astype(int)

    # ---- 索引映射 ----
    P = list(range(0, num_dep))  # 仓库：0..|P|-1
    S = list(range(num_dep, num_dep + num_sat))  # 卫星：|P|..|P|+|S|-1
    Z = list(range(num_dep + num_sat, num_nodes))  # 客户

    def cust_local_to_node(k):
        return num_dep + num_sat + k

    # ---- 距离矩阵（单位保持和你原代码一致：/10）----
    dist = np.zeros((num_nodes, num_nodes))
    for i in range(num_nodes):
        for j in range(num_nodes):
            if i != j:
                dist[i, j] = math.hypot(all_coords[i, 0] - all_coords[j, 0],
                                        all_coords[i, 1] - all_coords[j, 1])
    dist /= 10.0

    # ---- 时间窗与服务时间：如果实例里有则读，没有则给宽松默认 ----
    tw_start = np.zeros(num_cus, dtype=float)
    tw_end = np.full(num_cus, float('inf'), dtype=float)
    service = np.zeros(num_cus, dtype=float)

    # 统一把列名转成小写+去空格，方便匹配
    colmap = {c.lower().strip(): i for i, c in enumerate(df.columns)}

    # 尝试一些常见命名：这里包含你当前实例的 "start TW" / "end TW" / "service time"
    for name in ["start tw", "start_tw", "ready", "earliest", "e"]:
        key = name.lower()
        if key in colmap:
            tw_start = df.iloc[idx_cust, colmap[key]].to_numpy(dtype=float)
            break

    for name in ["end tw", "end_tw", "due", "latest", "l"]:
        key = name.lower()
        if key in colmap:
            tw_end = df.iloc[idx_cust, colmap[key]].to_numpy(dtype=float)
            break

    for name in ["service time", "servicetime", "service"]:
        key = name.lower()
        if key in colmap:
            service = df.iloc[idx_cust, colmap[key]].to_numpy(dtype=float)
            break

    # 卡车弧集：P∪S，禁 P->P
    nodes_truck = list(range(num_dep + num_sat))
    A_truck = [(i, j) for i in nodes_truck for j in nodes_truck if i != j and not (i in P and j in P)]

    return {
        "coords_all": all_coords,
        "num_dep": num_dep, "num_sat": num_sat, "num_cus": num_cus, "num_nodes": num_nodes,
        "P": P, "S": S, "Z": Z,
        "Dd": Dd, "Dp": Dp,
        "dist": dist,
        "A_truck": A_truck,
        "cust_local_to_node": cust_local_to_node,
        "tw_start": tw_start,
        "tw_end": tw_end,
        "service": service
    }


_const = (DR['g'] ** 3) / math.sqrt(2.0 * DR['rho'] * DR['zeta'] * DR['n'])
_const = (DR['g'] ** 3) / math.sqrt(2.0 * DR['rho'] * DR['zeta'] * DR['n'])


def energy_segment(load, d_ij):
    """飞行段能耗（Wh）。load 为当前载荷（kg 或与需求同量纲），d_ij 为距离（与 dist 同量纲）。"""
    t_hours = d_ij / max(1e-9, DR['f_s'])
    power_W = (DR['W'] + load) ** 1.5 * _const
    return power_W * t_hours


def hover_energy(load, wait_h):
    """悬停等待能耗（Wh）。wait_h 与时间窗/飞行时间使用同一时间单位（通常是小时）。"""
    if wait_h <= 1e-12:
        return 0.0
    power_W = (DR['W'] + load) ** 1.5 * _const
    return power_W * wait_h


def route_feasible_energy(s, seq, inst):
    """
    检查给定卫星 s 出发、访问客户序列 seq 的无人机子回路是否可行，
    并返回 (可行性, 总能耗E_total, 起飞载重start_load)。

    关键规则（与你的需求一致）：
    1) 能耗 = 各飞行段能耗 +（除首客户外）时间窗等待悬停能耗（按当段到达时载荷计算）；
    2) 首个客户不等待：允许延后起飞，使到达首客户时间 = max(飞行时间, 首客户时间窗起点)；
    3) 最后一段返回“最近卫星”（按距离最小）。
    """
    P, S, Z = inst["P"], inst["S"], inst["Z"]
    num_dep, num_sat = inst["num_dep"], inst["num_sat"]
    Dd, Dp = inst["Dd"], inst["Dp"]
    dist = inst["dist"]

    tw_start = inst.get("tw_start", None)
    tw_end = inst.get("tw_end", None)
    service = inst.get("service", None)
    has_tw = (tw_start is not None) and (tw_end is not None)

    if not seq:
        return False, 0.0, 0

    # 起飞载重 = 该子回路的总送货量
    start_load = 0
    for z in seq:
        zloc = z - (num_dep + num_sat)
        start_load += int(Dd[zloc])
    if start_load > Q2:
        return False, 0.0, 0

    load = float(start_load)
    E_total = 0.0

    # ---------- 首客户：延后起飞（不产生等待悬停能耗） ----------
    prev = s
    z0 = seq[0]
    zloc0 = z0 - (num_dep + num_sat)
    d0 = float(dist[prev, z0])
    t_fly0 = d0 / max(1e-9, DR['f_s'])
    if has_tw:
        e0 = float(tw_start[zloc0]) / 60.0  # minutes -> hours
        l0 = float(tw_end[zloc0]) / 60.0  # minutes -> hours
        # 若直接飞行都超过最晚到达，则不可行
        if t_fly0 > l0 + 1e-9:
            return False, 0.0, 0
        t = max(t_fly0, e0)  # 通过延后起飞消除首客户等待
    else:
        t = t_fly0

    # 飞行能耗：从卫星到首客户
    E_total += energy_segment(load, d0)
    if E_total > DR['E'] + 1e-9:
        return False, 0.0, 0

    # 首客户服务
    if service is not None and len(service) > 0:
        t += float(service[zloc0]) / 60.0  # minutes -> hours

    # 载重更新：送货减少，取件增加
    load = load - float(Dd[zloc0]) + float(Dp[zloc0])
    if load < -1e-9 or load > Q2 + 1e-9:
        return False, 0.0, 0

    prev = z0

    # ---------- 后续客户：时间窗等待产生悬停能耗 ----------
    for z in seq[1:]:
        zloc = z - (num_dep + num_sat)
        dij = float(dist[prev, z])
        t_fly = dij / max(1e-9, DR['f_s'])
        t += t_fly

        if has_tw:
            e = float(tw_start[zloc]) / 60.0  # minutes -> hours
            l = float(tw_end[zloc]) / 60.0  # minutes -> hours
            if t > l + 1e-9:
                return False, 0.0, 0
            if t < e - 1e-9:
                wait_h = e - t
                # 等待发生在服务前，载荷为“到达该客户时”的载荷
                E_total += hover_energy(load, wait_h)
                if E_total > DR['E'] + 1e-9:
                    return False, 0.0, 0
                t = e

        # 服务时间
        if service is not None and len(service) > 0:
            t += float(service[zloc]) / 60.0  # minutes -> hours

        # 飞行能耗（prev->z），按“起飞前载荷”（即到达前载荷）计算
        E_total += energy_segment(load, dij)
        if E_total > DR['E'] + 1e-9:
            return False, 0.0, 0

        # 载重更新
        load = load - float(Dd[zloc]) + float(Dp[zloc])
        if load < -1e-9 or load > Q2 + 1e-9:
            return False, 0.0, 0

        prev = z

    # ---------- 返航：返回最近卫星 ----------
    best_sat = None
    best_d = float("inf")
    # 返航：返回“最后一个客户”最近的卫星（按距离最小）
    # 注意：这里用 prev（即最后一个客户）计算距离，而不是用起飞卫星 s
    for s2 in S:
        d_back = float(dist[prev, s2])
        if d_back < best_d:
            best_d = d_back
            best_sat = s2

    if best_sat is None:
        return False, None, None
    # 返航飞行能耗
    E_total += energy_segment(load, best_d)
    if E_total > DR['E'] + 1e-9:
        return False, 0.0, 0

    return True, float(E_total), int(start_load)


def _route_first_tw_start_h(start_sat, seq_customers, inst):
    if not seq_customers:
        return 0.0
    tw_start = inst.get("tw_start", None)
    if tw_start is None:
        return 0.0
    z0 = int(seq_customers[0])
    zloc0 = _zloc_from_node(z0, inst)
    return float(tw_start[zloc0]) / 60.0


def _route_takeoff_min_no_hover_first_h(start_sat, seq_customers, inst):
    if not seq_customers:
        return 0.0
    tw_start = inst.get("tw_start", None)
    if tw_start is None:
        return 0.0
    dist = inst["dist"]
    z0 = int(seq_customers[0])
    zloc0 = _zloc_from_node(z0, inst)
    e0 = float(tw_start[zloc0]) / 60.0
    fly0 = float(dist[int(start_sat), z0]) / max(1e-9, DR['f_s'])
    return max(0.0, e0 - fly0)


def _route_time_feasible_given_takeoff(start_sat, seq_customers, inst, takeoff_h):
    tw_start = inst.get("tw_start", None)
    tw_end = inst.get("tw_end", None)
    service = inst.get("service", None)
    if tw_start is None or tw_end is None:
        return True
    dist = inst["dist"]

    t = float(takeoff_h)
    prev = int(start_sat)
    for z in seq_customers:
        z = int(z)
        zloc = _zloc_from_node(z, inst)
        t += float(dist[prev, z]) / max(1e-9, DR['f_s'])
        e = float(tw_start[zloc]) / 60.0
        l = float(tw_end[zloc]) / 60.0
        if math.isfinite(l) and t > l + 1e-9:
            return False
        if t < e - 1e-9:
            t = e
        if service is not None and len(service) > 0:
            t += float(service[zloc]) / 60.0
        prev = z
    return True


def _route_takeoff_latest_h(start_sat, seq_customers, inst):
    tw_end = inst.get("tw_end", None)
    if tw_end is None or not seq_customers:
        return float("inf")

    dist = inst["dist"]
    service = inst.get("service", None)

    ub = float("inf")
    prev = int(start_sat)
    acc = 0.0
    for z in seq_customers:
        z = int(z)
        zloc = _zloc_from_node(z, inst)
        acc += float(dist[prev, z]) / max(1e-9, DR['f_s'])
        l = float(tw_end[zloc]) / 60.0
        if math.isfinite(l):
            ub = min(ub, l - acc)
        if service is not None and len(service) > 0:
            acc += float(service[zloc]) / 60.0
        prev = z

    if not math.isfinite(ub):
        return float("inf")

    hi = max(0.0, float(ub))
    lo = 0.0
    if _route_time_feasible_given_takeoff(start_sat, seq_customers, inst, hi):
        return hi
    for _ in range(40):
        mid = 0.5 * (lo + hi)
        if _route_time_feasible_given_takeoff(start_sat, seq_customers, inst, mid):
            lo = mid
        else:
            hi = mid
    return lo


def _add_loading_queue_constraints(m, inst, routes_by_s, gamma, deliv_mat):
    if (not ENABLE_LOADING_QUEUE) or float(LOADING_TIME_PER_DELIV_UNIT_MIN) <= 1e-12:
        return

    S = inst["S"]
    tw_start = inst.get("tw_start", None)

    dep_max = {}
    dep_min_no_hover = {}
    prio_key = {}
    for s in S:
        for r in range(len(routes_by_s[s])):
            seq = routes_by_s[s][r]
            dep_max[(s, r)] = _route_takeoff_latest_h(int(s), seq, inst)
            dep_min_no_hover[(s, r)] = _route_takeoff_min_no_hover_first_h(int(s), seq, inst)
            prio_key[(s, r)] = _route_first_tw_start_h(int(s), seq, inst) if tw_start is not None else 0.0

    tB = {}
    tD = {}
    for s in S:
        for r in range(len(routes_by_s[s])):
            tB[(s, r)] = m.addVar(lb=0.0, ub=M_TIME, vtype=GRB.CONTINUOUS, name=f"tB[{s},{r}]")
            tD[(s, r)] = m.addVar(lb=0.0, ub=M_TIME, vtype=GRB.CONTINUOUS, name=f"tD[{s},{r}]")
    m.update()

    for s in S:
        for r in range(len(routes_by_s[s])):
            m.addConstr(tB[(s, r)] <= M_TIME * gamma[(s, r)], name=f"tB_ub[{s},{r}]")
            m.addConstr(tD[(s, r)] <= M_TIME * gamma[(s, r)], name=f"tD_ub[{s},{r}]")
            m.addConstr(tD[(s, r)] >= tB[(s, r)], name=f"tD_ge_tB[{s},{r}]")

            mn = float(dep_min_no_hover[(s, r)])
            if mn > 1e-12:
                m.addConstr(tD[(s, r)] >= mn - M_TIME * (1 - gamma[(s, r)]), name=f"tD_minNoHover[{s},{r}]")

            mx = dep_max[(s, r)]
            if math.isfinite(mx):
                m.addConstr(tD[(s, r)] <= float(mx) + M_TIME * (1 - gamma[(s, r)]), name=f"tD_max[{s},{r}]")

    for s in S:
        order = list(range(len(routes_by_s[s])))
        order.sort(key=lambda r: (prio_key[(s, r)], r))
        prefix_expr = 0.0
        for r in order:
            load_h = (float(LOADING_TIME_PER_DELIV_UNIT_MIN) * float(deliv_mat[(s, r)])) / 60.0
            prefix_expr = prefix_expr + load_h * gamma[(s, r)]
            m.addConstr(tB[(s, r)] >= prefix_expr - M_TIME * (1 - gamma[(s, r)]), name=f"loadB_lb[{s},{r}]")
            m.addConstr(tB[(s, r)] <= prefix_expr + M_TIME * (1 - gamma[(s, r)]), name=f"loadB_ub[{s},{r}]")

def build_pool_exact(inst):
    """
    生成“全列池（full pool）”：
    - 不限制 route 长度（上限设为客户数），DFS 枚举所有可行序列；
    - 对每个 coverage set 仅保留能耗最小的那条序列；
    - 再做一次简单的支配性过滤（子集+能耗不优的剔除）。
    """
    P, S, Z = inst["P"], inst["S"], inst["Z"]
    dist = inst["dist"]
    num_dep, num_sat = inst["num_dep"], inst["num_sat"]
    num_cus = inst["num_cus"]

    max_stops = min(num_cus, ROUTE_MAX_STOPS_EXACT if ROUTE_MAX_STOPS_EXACT is not None else num_cus)

    routes_by_s = {s: [] for s in S}
    energy_by_sr, cover_mat, deliv_mat = {}, {}, {}

    for s in S:
        Z_sorted = sorted(Z, key=lambda z: dist[s, z])

        # memo: (tuple(path)) -> (ok, E, ds)
        memo = {}

        def eval_path(path):
            key = tuple(path)
            if key in memo:
                return memo[key]
            ok, E, ds = route_feasible_energy(s, list(path), inst)
            memo[key] = (ok, E, ds)
            return memo[key]

        # best per subset (bitmask over customers 0..num_cus-1)
        best = {}  # mask -> (E, path_tuple, ds)

        def mask_of(path):
            m = 0
            for z in path:
                zloc = z - (num_dep + num_sat)
                m |= (1 << zloc)
            return m

        def dfs(path):
            if len(path) > 0:
                ok, E, ds = eval_path(path)
                if ok:
                    msk = mask_of(path)
                    cur = best.get(msk, None)
                    if (cur is None) or (E < cur[0] - 1e-12):
                        best[msk] = (E, tuple(path), ds)

            if len(path) >= max_stops:
                return

            used = set(path)
            # 按离卫星距离排序扩展（更容易先找到可行长路）
            for z in Z_sorted:
                if z in used:
                    continue
                ok, _, _ = eval_path(tuple(path) + (z,))
                if ok:
                    dfs(tuple(path) + (z,))

        dfs(tuple())

        # 转为列表并做简单支配过滤：如果 coverage_i ⊂ coverage_j 且 E_i >= E_j 则删 i
        cov_list = []
        for msk, (E, path, ds) in best.items():
            cov = {i for i in range(num_cus) if (msk >> i) & 1}
            cov_list.append((cov, path, float(E), int(ds)))
        cov_list.sort(key=lambda x: (len(x[0]), x[2]))

        final = []
        for i, (cov_i, path_i, E_i, ds_i) in enumerate(cov_list):
            dominated = False
            for j, (cov_j, path_j, E_j, ds_j) in enumerate(cov_list):
                if i == j:
                    continue
                if cov_i and cov_i.issubset(cov_j) and cov_i != cov_j and E_i >= E_j - 1e-12:
                    dominated = True
                    break
            if not dominated:
                final.append((path_i, E_i, ds_i))

        for r, (path, E, ds) in enumerate(final):
            routes_by_s[s].append(list(path))
            energy_by_sr[(s, r)] = float(E)
            cover_mat[(s, r)] = set(z - (num_dep + num_sat) for z in path)
            deliv_mat[(s, r)] = int(ds)

    return routes_by_s, energy_by_sr, cover_mat, deliv_mat



def build_pool_exact_global(inst):
    # Enumerate ALL feasible second-echelon customer sequences for each satellite.
    # Unlike build_pool_exact(), this routine does NOT collapse routes by coverage subset
    # and does NOT apply subset-dominance pruning.
    # This is required for stage-0 global optimality when loading-queue constraints are
    # active, because different customer orders with the same coverage subset may induce
    # different takeoff windows / queue interactions in the master problem.
    P, S, Z = inst["P"], inst["S"], inst["Z"]
    dist = inst["dist"]
    num_dep, num_sat = inst["num_dep"], inst["num_sat"]
    num_cus = inst["num_cus"]

    max_stops = min(num_cus, ROUTE_MAX_STOPS_EXACT if ROUTE_MAX_STOPS_EXACT is not None else num_cus)

    routes_by_s = {s: [] for s in S}
    energy_by_sr, cover_mat, deliv_mat = {}, {}, {}

    for s in S:
        Z_sorted = sorted(Z, key=lambda z: dist[s, z])
        memo = {}
        seen_routes = set()

        def eval_path(path):
            key = tuple(path)
            if key in memo:
                return memo[key]
            ok, E, ds = route_feasible_energy(s, list(path), inst)
            memo[key] = (ok, E, ds)
            return memo[key]

        def dfs(path):
            if len(path) > 0:
                ok, E, ds = eval_path(path)
                if ok:
                    tpath = tuple(path)
                    if tpath not in seen_routes:
                        seen_routes.add(tpath)
                        r = len(routes_by_s[s])
                        routes_by_s[s].append(list(tpath))
                        energy_by_sr[(s, r)] = float(E)
                        cover_mat[(s, r)] = set(z - (num_dep + num_sat) for z in tpath)
                        deliv_mat[(s, r)] = int(ds)
                else:
                    return

            if len(path) >= max_stops:
                return

            used = set(path)
            for z in Z_sorted:
                if z in used:
                    continue
                ok, _, _ = eval_path(tuple(path) + (z,))
                if ok:
                    dfs(tuple(path) + (z,))

        dfs(tuple())

    return routes_by_s, energy_by_sr, cover_mat, deliv_mat

def warn_uncovered_customers(inst, cover_mat, routes_by_s):
    num_cus = inst["num_cus"]
    uncovered = []
    for zloc in range(num_cus):
        appears = any(zloc in cover_mat[(s, r)]
                      for s in inst["S"] for r in range(len(routes_by_s[s])))
        if not appears:
            uncovered.append(zloc)
    if uncovered:
        print("[WARN] 列池不完整：",
              [u + 1 for u in uncovered])  # 1-based
    return uncovered


def build_master(inst, routes_by_s, energy_by_sr, cover_mat, deliv_mat, as_lp=False, w1=1.0, w2=1.0):
    P, S = inst["P"], inst["S"]
    dist = inst["dist"]
    A_truck = inst["A_truck"]
    Dd, Dp = inst["Dd"], inst["Dp"]
    num_cus = inst["num_cus"]

    total_dem = float((inst["Dd"] + inst["Dp"]).sum())
    max_trucks = max(1, math.ceil(max(1e-9, total_dem) / Q1) * 2)
    max_trucks = min(max_trucks, len(inst["S"]))  # 也可直接 = len(S)
    V = list(range(max_trucks))

    m = gp.Model("2E_master_exact")
    m.Params.OutputFlag = OUTPUT_LOG
    V = list(range(max_trucks))
    VTYPE = GRB.CONTINUOUS if as_lp else GRB.BINARY

    x = m.addVars(A_truck, V, lb=0, ub=1, vtype=VTYPE, name="x")
    tau = m.addVars(V, lb=0, ub=1, vtype=VTYPE, name="tau")
    u_tr = m.addVars(S, V, lb=0.0, ub=len(S), vtype=GRB.CONTINUOUS, name="u_tr")
    t_unload = m.addVars(S, V, lb=0.0, ub=Q1, vtype=GRB.CONTINUOUS, name="t_unload")
    y_dep = m.addVars(P, V, vtype=VTYPE, name="y_dep")

    gamma = {}
    for s in S:
        for r in range(len(routes_by_s[s])):
            gamma[(s, r)] = m.addVar(lb=0, ub=1, vtype=VTYPE, name=f"gamma[{s},{r}]")
    m.update()

    # ---- Satellite loading-queue constraints (optional) ----
    _add_loading_queue_constraints(m, inst, routes_by_s, gamma, deliv_mat)

    # ---------------- Weighted-sum scalarization (two objectives) ----------------
    # obj1 (fixed):    h1 * (#trucks used) + h2 * (#drone routes used)
    # obj2 (variable): truck travel cost + drone energy cost
    #
    # Gurobi solves:   min  w1*obj1 + w2*obj2
    #
    # NOTE: Sweeping (w1,w2) may produce repeated optimal solutions (same (obj1,obj2))
    # but it is still a standard way to generate multiple Pareto-relevant points.

    # Dummy columns (LP only): allow cover constraints == 1 even if the pool is incomplete.
    # Penalty is added to the objective directly to discourage dummy usage.
    var_dum = {}
    dummy_penalty = 0.0
    if USE_DUMMY_COLUMNS and as_lp:
        for zloc in range(num_cus):
            v = m.addVar(lb=0, ub=1, vtype=GRB.CONTINUOUS, name=f"gamma_dum[{zloc}]")
            var_dum[zloc] = v
            dummy_penalty += DUMMY_COST * v
        m.update()

    obj1_expr = gp.quicksum(h1 * tau[v] for v in V) \
                + gp.quicksum(h2 * gamma[(s, r)] for s in S for r in range(len(routes_by_s[s])))

    obj2_expr = gp.quicksum(c_v * dist[i, j] * x[i, j, v] for (i, j) in A_truck for v in V) \
                + gp.quicksum(c_u * energy_by_sr[(s, r)] * gamma[(s, r)]
                              for s in S for r in range(len(routes_by_s[s])))

    obj_weighted = float(w1) * obj1_expr + float(w2) * obj2_expr + dummy_penalty
    m.setObjective(obj_weighted, GRB.MINIMIZE)

    for v in V:
        m.addConstr(gp.quicksum(x[i, j, v] for i in P for j in S if (i, j) in A_truck) == tau[v], name=f"dep_v{v}")
        m.addConstr(gp.quicksum(x[i, j, v] for i in S for j in P if (i, j) in A_truck) == tau[v], name=f"ret_v{v}")

    for v in V:
        m.addConstr(gp.quicksum(y_dep[p, v] for p in P) == tau[v], name=f"one_dep_v{v}")
        for p in P:
            m.addConstr(gp.quicksum(x[p, s, v] for s in S if (p, s) in A_truck) == y_dep[p, v],
                        name=f"start_at_p{p}_v{v}")
            m.addConstr(gp.quicksum(x[s, p, v] for s in S if (s, p) in A_truck) == y_dep[p, v],
                        name=f"return_to_p{p}_v{v}")

    for v in V:
        for s in S:
            m.addConstr(gp.quicksum(x[i, s, v] for i in (P + S) if (i, s) in A_truck) ==
                        gp.quicksum(x[s, j, v] for j in (P + S) if (s, j) in A_truck),
                        name=f"flow_s{s}_v{v}")

    # MTZ on S
    if len(S) >= 2:
        T = len(S)
        for v in V:
            for i in S:
                for j in S:
                    if i != j and (i, j) in A_truck:
                        m.addConstr(u_tr[i, v] - u_tr[j, v] + T * x[i, j, v] <= T - 1,
                                    name=f"mtz_s{i}_s{j}_v{v}")

    for v in V:
        for (i, j) in A_truck:
            m.addConstr(x[i, j, v] <= tau[v], name=f"use_x_tau_{i}_{j}_v{v}")

    for s in S:
        visit_s = gp.quicksum(x[i, s, v] for v in V for i in (P + S) if (i, s) in A_truck)
        for r in range(len(routes_by_s[s])):
            m.addConstr(gamma[(s, r)] <= visit_s, name=f"route_visit_link_s{s}_r{r}")

    cover_constr = {}

    for zloc in range(num_cus):
        expr = gp.quicksum(gamma[(s, r)] for s in S for r in range(len(routes_by_s[s]))
                           if zloc in cover_mat[(s, r)])
        if USE_DUMMY_COLUMNS and as_lp:
            expr = expr + var_dum[zloc]
        cc = m.addConstr(expr == 1, name=f"cover_z{zloc}")
        cover_constr[zloc] = cc

    for s in S:
        for v in V:
            m.addConstr(t_unload[s, v] <= Q1 * gp.quicksum(x[i, s, v] for i in (P + S) if (i, s) in A_truck),
                        name=f"unload_visit_s{s}_v{v}")
    for v in V:
        m.addConstr(gp.quicksum(t_unload[s, v] for s in S) <= Q1 * tau[v], name=f"unload_cap_v{v}")
    supply_constr = {}
    for s in S:
        rhs = gp.quicksum(deliv_mat[(s, r)] * gamma[(s, r)] for r in range(len(routes_by_s[s])))
        sc = m.addConstr(gp.quicksum(t_unload[s, v] for v in V) >= rhs, name=f"supply_s{s}")
        supply_constr[s] = sc

    for v in range(len(V) - 1):
        m.addConstr(tau[v] >= tau[v + 1], name=f"sym_tau_{v}")

    m._V = V
    m._inst = inst
    m._w1 = float(w1)
    m._w2 = float(w2)
    return m, gamma, cover_constr, supply_constr


def pricing_round(m, routes_by_s, energy_by_sr, cover_mat, deliv_mat, cover_constr, supply_constr):
    inst = m._inst
    P, S, Z = inst["P"], inst["S"], inst["Z"]
    num_dep, num_sat = inst["num_dep"], inst["num_sat"]
    dist = inst["dist"]

    try:
        pi = {z: cover_constr[z].Pi for z in cover_constr}
        lam = {s: supply_constr[s].Pi for s in supply_constr}
    except gp.GurobiError:
        return 0

    added = 0
    for s in S:
        cusp = sorted(Z, key=lambda z: dist[s, z])
        seeds = []
        for z in cusp:
            ok, E, ds = route_feasible_energy(s, [z], inst)
            if ok:
                rc = (h2 + c_u * E) - pi.get(z - (num_dep + num_sat), 0.0) - lam.get(s, 0.0) * ds
                seeds.append((-rc, [z], E, ds, rc))
        seeds.sort(reverse=True)
        frontier = seeds[:PRICING_BEAM_WIDTH]
        cand = []
        for _depth in range(2, PRICING_MAX_STOPS + 1):
            new_front = []
            for _, seq, Ecur, dcur, rc_cur in frontier:
                used = set(seq)
                for z in cusp:
                    if z in used: continue
                    ok, E, ds = route_feasible_energy(s, seq + [z], inst)
                    if not ok: continue
                    rc = (h2 + c_u * E) \
                         - sum(pi.get(zz - (num_dep + num_sat), 0.0) for zz in (seq + [z])) \
                         - lam.get(s, 0.0) * ds
                    new_front.append((-rc, seq + [z], E, ds, rc))
            cand += frontier
            new_front.sort(reverse=True)
            frontier = new_front[:PRICING_BEAM_WIDTH]
        cand += frontier

        seen = set(tuple(seq) for seq in routes_by_s[s])
        negs = []
        for _, seq, E, ds, rc in cand:
            if rc < -RC_EPS and tuple(seq) not in seen:
                negs.append((rc, seq, E, ds))
        negs.sort(key=lambda x: x[0])

        for rc, seq, E, ds in negs[:PRICING_ADD_PER_S]:
            r = len(routes_by_s[s])
            routes_by_s[s].append(list(seq))
            energy_by_sr[(s, r)] = E
            cover_mat[(s, r)] = set(z - (num_dep + num_sat) for z in seq)
            deliv_mat[(s, r)] = int(ds)
            var = m.addVar(lb=0, ub=1, vtype=GRB.CONTINUOUS, name=f"gamma[{s},{r}]")
            w1 = getattr(m, '_w1', 1.0)
            w2 = getattr(m, '_w2', 1.0)
            var.Obj = w1 * h2 + w2 * c_u * E
            # 覆盖
            for zloc in cover_mat[(s, r)]:
                m.chgCoeff(m.getConstrByName(f"cover_z{zloc}"), var, 1.0)
            # 供给
            m.chgCoeff(m.getConstrByName(f"supply_s{s}"), var, -ds)
            # 访问联动
            visit_s = gp.quicksum(m.getVarByName(f"x[{i},{s},{v}]")
                                  for v in m._V for i in (inst["P"] + inst["S"]) if (i, s) in inst["A_truck"])
            m.addConstr(var <= visit_s, name=f"route_visit_link_pricing_s{s}_r{r}")
            added += 1

    if added > 0:
        m.update()
    return added


def extract_paths(m, inst, routes_by_s):
    P, S = inst["P"], inst["S"]

    PS = P + S
    ps_pos = {node: idx for idx, node in enumerate(PS)}

    def ps_local_12(node):
        # 仓库 1..|P|, 卫星 |P|+1..|P|+|S|
        return ps_pos[node] + 1

    first_seq = []
    Vidx = sorted({
        int(v.VarName.split('[')[1].split(']')[0])
        for v in m.getVars() if v.VarName.startswith('tau[')
    })

    # ---------- 一层路径提取 ----------
    for v in Vidx:
        tv = m.getVarByName(f"tau[{v}]")
        if tv is None or tv.X < 0.5:
            continue

        # 1) 找起点仓库 p0：看 y_dep[p,v]
        p0 = None
        for p in P:
            yv = m.getVarByName(f"y_dep[{p},{v}]")
            if yv is not None and yv.X > 0.5:
                p0 = p
                break
        if p0 is None:
            continue

        # 2) 找第一个访问的卫星 s1：看 x[p0,s,v]
        s1 = None
        for s in S:
            xv = m.getVarByName(f"x[{p0},{s},{v}]")
            if xv is not None and xv.X > 0.5:
                s1 = s
                break
        if s1 is None:
            continue

        first_seq += [ps_local_12(p0), ps_local_12(s1)]

        # 3) 继续在卫星之间按 x[s_cur,s_next,v] 拓展
        cur = s1
        visitedS = {cur}
        while True:
            nxt = None
            for s in S:
                xv = m.getVarByName(f"x[{cur},{s},{v}]")
                if xv is not None and xv.X > 0.5:
                    nxt = s
                    break
            if nxt is None:
                break
            first_seq.append(ps_local_12(nxt))
            if nxt in visitedS:
                break
            visitedS.add(nxt)
            cur = nxt

    first_str = " ".join(map(str, first_seq)) if first_seq else ""

    # ---------- 二层路径提取（增加“返回最近站点”） ----------
    num_dep, num_sat = inst["num_dep"], inst["num_sat"]
    dist = inst["dist"]

    def sat_local_12(s):
        """卫星局部编号：1..|S|"""
        return (s - num_dep) + 1

    def cus_local_12(z):
        """客户局部编号：|S|+1..|S|+|Z|"""
        return num_sat + (z - (num_dep + num_sat))

    second_list = []
    for s in S:
        for r in range(len(routes_by_s[s])):
            gv = m.getVarByName(f"gamma[{s},{r}]")
            if gv is None or gv.X is None or gv.X <= 0.5:
                continue

            route = routes_by_s[s][r]  # 这是一个客户节点序列（全局编号）

            if len(route) == 0:
                continue

            # 起点卫星
            second_list.append(sat_local_12(s))
            # 依次输出该子回路内的客户
            for z in route:
                second_list.append(cus_local_12(z))

            # ---- 新增：最后一个客户返回最近卫星 ----
            last_z = route[-1]
            best_sat = None
            best_d = float("inf")
            for s2 in S:
                d2 = dist[last_z, s2]
                if d2 < best_d:
                    best_d = d2
                    best_sat = s2

            if best_sat is None:
                best_sat = s  # 极端兜底（理论上不会发生）

            second_list.append(sat_local_12(best_sat))

    second_str = " ".join(map(str, second_list)) if second_list else ""

    return first_str, second_str


def mip_progress_cb(model, where):
    if where == GRB.Callback.MIP:
        pass


def solve_instance(csv_path, w1=1.0, w2=1.0):
    t0 = time.perf_counter()
    inst = read_instance(csv_path)

    # Exact benchmarking: build a (near-)full route pool by exhaustive enumeration.
    # Note: this is only practical for small instances.
    routes_by_s, energy_by_sr, cover_mat, deliv_mat = build_pool_exact(inst)

    # Pool statistics
    pool_cols = sum(len(routes_by_s[s]) for s in routes_by_s)
    print(f"[pool] total drone-route columns = {pool_cols}")

    miss = warn_uncovered_customers(inst, cover_mat, routes_by_s)
    if miss and not USE_DUMMY_COLUMNS:
        return {"file": os.path.basename(csv_path), "status": "Unsolvable"}

    m, gamma, cover_constr, supply_constr = build_master(inst, routes_by_s, energy_by_sr, cover_mat, deliv_mat,
                                                         as_lp=True, w1=w1, w2=w2)
    m.Params.Method = 1
    m.Params.OutputFlag = OUTPUT_LOG
    m.Params.TimeLimit = TIME_LIMIT
    m.optimize()
    if m.Status not in (GRB.OPTIMAL, GRB.SUBOPTIMAL):
        return {"file": os.path.basename(csv_path), "status": "Unsolvable"}
    lp_val = float(m.ObjVal)

    for rnd in range(1, PRICING_ROUNDS + 1):
        added = pricing_round(m, routes_by_s, energy_by_sr, cover_mat, deliv_mat, cover_constr, supply_constr)
        if added == 0: break
        m.optimize()
        if m.Status not in (GRB.OPTIMAL, GRB.SUBOPTIMAL):
            break
        lp_val = float(m.ObjVal)

    for v in m.getVars():
        n = v.VarName
        if (n.startswith("gamma[") or n.startswith("x[") or
                n.startswith("tau[") or n.startswith("y_dep[")):
            v.VType = GRB.BINARY
    m.update()

    m.Params.Presolve = 2
    m.Params.Cuts = 2
    m.Params.Heuristics = 0.2
    m.Params.MIPFocus = 3
    m.Params.MIPGap = 0.0
    m.Params.TimeLimit = 3600
    m.Params.NodefileStart = 0.5
    m.optimize(mip_progress_cb)

    if m.SolCount == 0 or m.Status in (GRB.INFEASIBLE, GRB.INF_OR_UNBD):
        return {"file": os.path.basename(csv_path), "status": "Unsolvable"}

    proven_optimal = (m.Status == GRB.OPTIMAL) and (abs(float(m.ObjBound) - float(m.ObjVal)) <= 1e-6)
    solve_status = int(m.Status)
    solve_status_str = {
        GRB.OPTIMAL: 'OPTIMAL',
        GRB.TIME_LIMIT: 'TIME_LIMIT',
        GRB.INFEASIBLE: 'INFEASIBLE',
        GRB.INF_OR_UNBD: 'INF_OR_UNBD',
        GRB.INTERRUPTED: 'INTERRUPTED',
        GRB.SUBOPTIMAL: 'SUBOPTIMAL',
    }.get(m.Status, str(m.Status))

    print(f"[mip] status={solve_status_str}, ObjVal={m.ObjVal}, ObjBound={m.ObjBound}, "
          f"MIPGap={getattr(m, 'MIPGap', float('nan'))}, proven_optimal={proven_optimal}")

    print("[check] rows, cols, nz =", m.NumConstrs, m.NumVars, m.NumNZs)

    # 计算 obj1 和 obj2
    dist = inst["dist"]
    obj1 = 0.0  # f2：固定成本
    obj2 = 0.0  # f1：行驶成本 + 能耗成本
    for var in m.getVars():
        name = var.VarName
        val = var.X
        if abs(val) < 1e-8:
            continue
        if name.startswith("x["):
            inside = name[2:-1]
            i_str, j_str, v_str = inside.split(",")
            i = int(i_str)
            j = int(j_str)
            obj2 += c_v * dist[i, j] * val
        elif name.startswith("gamma["):
            inside = name[6:-1]
            s_str, r_str = inside.split(",")
            s = int(s_str)
            r = int(r_str)
            E = energy_by_sr.get((s, r), 0.0)
            obj2 += c_u * E * val
            obj1 += h2 * val
        elif name.startswith("tau["):
            v_str = name[4:-1]
            obj1 += h1 * val

    sum_obj = obj2 + obj1
    weighted_obj = float(w1) * obj1 + float(w2) * obj2
    weighted_obj_norm = weighted_obj / max(1e-12, (float(w1) + float(w2)))
    gap = float(m.MIPGap) if hasattr(m, "MIPGap") else max(0.0, (m.ObjBound - m.ObjVal) / max(1e-9, abs(m.ObjVal)))

    first_str, second_str = extract_paths(m, inst, routes_by_s)
    elapsed = time.perf_counter() - t0
    return {
        "file": os.path.basename(csv_path),
        "time": round(elapsed, 2),
        "first": first_str,
        "second": second_str,
        "obj1": round(obj1, 6),
        "obj2": round(obj2, 6),
        "sum_obj": round(sum_obj, 6),
        "w1": float(w1),
        "w2": float(w2),
        "weighted_obj": round(weighted_obj, 6),
        "weighted_obj_norm": round(weighted_obj_norm, 6),
        "lp": round(lp_val, 6),
        "gap": round(gap, 6),
        "status": "OK"
    }


# ======================================================================================
# Dynamic pickup replanning with route prefix-freeze (single-loop per drone)
# Based on the original energy/time-window logic in route_feasible_energy().
#
# Key ideas:
#   - Stage0: plan ONLY deliveries + pickups with release==0 (release=max(0, tw_start-60min)).
#   - Later stages: when new pickups are released, we "freeze" each drone's executed prefix
#     up to the first customer service-completion time >= replan_time, then re-optimize the suffix.
#   - Delivery ownership is fixed per initial drone route (cannot transfer deliveries across drones).
#   - New drone sortie is allowed for uncovered pickups, with a mild extra penalty (PenaltyNew).
#
# NOTE:
#   - This implementation is designed to be most reliable on small instances (e.g., 15 customers),
#     consistent with the script's EXACT pool enumeration.
#   - For large instances, exhaustive suffix enumeration may be slow; see MAX_SUFFIX_ENUM for safety.
# ======================================================================================

def _zloc_from_node(z_node, inst):
    return int(z_node - (inst["num_dep"] + inst["num_sat"]))


def _is_customer(node, inst):
    return node >= (inst["num_dep"] + inst["num_sat"])


def _cus_local_12(z_node, inst):
    """客户局部编号：|S|+1..|S|+|Z|"""
    return inst["num_sat"] + _zloc_from_node(z_node, inst) +1


def _sat_local_12(s_node, inst):
    """卫星局部编号：1..|S|"""
    return (s_node - inst["num_dep"]) + 1


def _format_route_12(start_sat, seq_customers, end_sat, inst):
    out = [_sat_local_12(start_sat, inst)]
    out += [_cus_local_12(z, inst) for z in seq_customers]
    out += [_sat_local_12(end_sat, inst)]
    return " ".join(map(str, out))


def _nearest_sat_from(node, inst):
    S = inst["S"]
    dist = inst["dist"]
    best_sat, best_d = None, float("inf")
    for s2 in S:
        d2 = float(dist[node, s2])
        if d2 < best_d:
            best_d = d2
            best_sat = s2
    if best_sat is None:
        best_sat = S[0]
    return int(best_sat), float(best_d)


def _compute_release_minutes(inst):
    tw_start = inst.get("tw_start", None)
    num_cus = inst["num_cus"]
    rel = np.zeros(num_cus, dtype=float)
    if tw_start is None or len(tw_start) == 0:
        return rel
    for zloc in range(num_cus):
        rel[zloc] = max(0.0, float(tw_start[zloc]) - 60)
    return rel


def _simulate_route_from_state(start_node, start_time_h, start_load, start_energy_used,
                               seq_customers, inst, ground_wait_free=False,
                               remaining_energy_cap=None):
    """
    Simulate a route suffix from an arbitrary state.
    Returns dict with:
      ok, end_time_h, end_load, end_energy_used, end_sat,
      timeline: list of dicts per visited customer (node, t_arrive_h, t_start_h, t_end_h, load_before, load_after, E_used)
    Rules:
      - Time windows enforced (minutes -> hours).
      - If ground_wait_free=True, waiting before first customer can be done by delaying departure with no hover energy.
        (Typical for a new drone waiting on ground at a satellite.)
      - Otherwise, early arrival waits consume hover energy.
      - Energy uses energy_segment(load_before, distance) for flights, hover_energy(load_at_wait, wait_h) for waits.
      - Return to nearest satellite after last customer.
      - remaining_energy_cap: if provided, enforce start_energy_used + additional <= cap (cap is total E, e.g., DR['E']).
    """
    Dd, Dp = inst["Dd"], inst["Dp"]
    dist = inst["dist"]
    tw_start = inst.get("tw_start", None)
    tw_end = inst.get("tw_end", None)
    service = inst.get("service", None)
    has_tw = (tw_start is not None) and (tw_end is not None)

    t = float(start_time_h)
    load = float(start_load)
    E_used = float(start_energy_used)

    timeline = []
    prev = int(start_node)

    if seq_customers is None:
        seq_customers = []

    for idx, z in enumerate(seq_customers):
        z = int(z)
        zloc = _zloc_from_node(z, inst)

        dij = float(dist[prev, z])
        t_fly = dij / max(1e-9, DR['f_s'])

        # Decide departure/arrival time
        if idx == 0 and ground_wait_free and _is_customer(prev, inst) is False:
            # Start from satellite/ground: can delay departure for first customer without hover energy
            t_arrive = t + t_fly
            if has_tw:
                e = float(tw_start[zloc]) / 60.0
                l = float(tw_end[zloc]) / 60.0
                if t_fly > l + 1e-9:
                    return {"ok": False}
                # delay departure so arrival meets earliest
                if t_arrive < e - 1e-9:
                    # push departure later: t becomes (e - t_fly)
                    t = e - t_fly
                    t_arrive = e
                if t_arrive > l + 1e-9:
                    return {"ok": False}
        else:
            # Normal: depart at current t, arrive at t+t_fly
            t_arrive = t + t_fly
            if has_tw:
                e = float(tw_start[zloc]) / 60.0
                l = float(tw_end[zloc]) / 60.0
                if t_arrive > l + 1e-9:
                    return {"ok": False}
                if t_arrive < e - 1e-9:
                    # wait (hover) until e
                    wait_h = e - t_arrive
                    # waiting occurs before service, at arrival load
                    E_used += hover_energy(load, wait_h)
                    if remaining_energy_cap is None:
                        if E_used > DR['E'] + 1e-9:
                            return {"ok": False}
                    else:
                        if E_used > remaining_energy_cap + 1e-9:
                            return {"ok": False}
                    t_arrive = e

        # Flight energy: prev->z with load before update
        E_used += energy_segment(load, dij)
        if remaining_energy_cap is None:
            if E_used > DR['E'] + 1e-9:
                return {"ok": False}
        else:
            if E_used > remaining_energy_cap + 1e-9:
                return {"ok": False}

        # Service start = arrive, service end add service time
        t_start = t_arrive
        t_end = t_start
        if service is not None and len(service) > 0:
            t_end += float(service[zloc]) / 60.0

        # Load update after service at z
        load_before = load
        load = load - float(Dd[zloc]) + float(Dp[zloc])
        if load < -1e-9 or load > Q2 + 1e-9:
            return {"ok": False}

        timeline.append({
            "node": z,
            "t_arrive_h": t_arrive,
            "t_start_h": t_start,
            "t_end_h": t_end,
            "load_before": load_before,
            "load_after": load,
            "E_used": E_used
        })

        # next leg departs at service end
        t = t_end
        prev = z

    # Return to nearest satellite from last visited node (or start_node if no customers)
    end_node = prev
    end_sat, d_back = _nearest_sat_from(end_node, inst)
    # Return flight energy
    E_used += energy_segment(load, d_back)
    if remaining_energy_cap is None:
        if E_used > DR['E'] + 1e-9:
            return {"ok": False}
    else:
        if E_used > remaining_energy_cap + 1e-9:
            return {"ok": False}
    t_ret = t + d_back / max(1e-9, DR['f_s'])

    return {
        "ok": True,
        "end_time_h": float(t_ret),
        "end_load": float(load),
        "end_energy_used": float(E_used),
        "end_sat": int(end_sat),
        "timeline": timeline
    }



def _simulate_full_route_from_sat(start_sat, seq_customers, inst, start_time_h=0.0, ground_wait_free_first=True):
    """
    Simulate a full route from a satellite.
    start_time_h: takeoff time (hours). If you want queue-based takeoff, pass tD here.
    ground_wait_free_first: if True, the drone may delay takeoff to eliminate first-customer waiting (no hover).
    """
    # start_load = sum deliveries in seq
    Dd = inst["Dd"]
    start_load = 0
    for z in seq_customers:
        zloc = _zloc_from_node(int(z), inst)
        start_load += int(Dd[zloc])
    if start_load > Q2:
        return {"ok": False}

    return _simulate_route_from_state(
        start_node=int(start_sat),
        start_time_h=float(start_time_h),
        start_load=float(start_load),
        start_energy_used=0.0,
        seq_customers=[int(z) for z in seq_customers],
        inst=inst,
        ground_wait_free=bool(ground_wait_free_first),
        remaining_energy_cap=DR['E']
    )

def _active_zlocs_stage0(inst, release_min):
    """Deliveries are always active at stage0. Pickups are active iff release==0."""
    Dd = inst["Dd"]
    Dp = inst["Dp"]
    num_cus = inst["num_cus"]
    active = []
    for zloc in range(num_cus):
        if int(Dd[zloc]) > 0:
            active.append(zloc)
        elif int(Dp[zloc]) > 0 and float(release_min[zloc]) <= 1e-9:
            active.append(zloc)
    return sorted(active)


def _make_inst_with_active_Z(inst, active_zlocs):
    """Return a shallow-copied inst with Z restricted to active customers (by global node ids)."""
    inst2 = dict(inst)
    num_dep, num_sat = inst["num_dep"], inst["num_sat"]
    active_nodes = [num_dep + num_sat + int(zloc) for zloc in active_zlocs]
    inst2["Z"] = active_nodes
    return inst2


def warn_uncovered_customers_active(inst, cover_mat, routes_by_s, active_zlocs):
    uncovered = []
    for zloc in active_zlocs:
        appears = any(zloc in cover_mat[(s, r)]
                      for s in inst["S"] for r in range(len(routes_by_s[s])))
        if not appears:
            uncovered.append(zloc)
    if uncovered:
        print("[warn] active customers uncovered in current pool:", uncovered)
    return uncovered


def build_master_with_active(inst, routes_by_s, energy_by_sr, cover_mat, deliv_mat,
                             active_zlocs, as_lp=False, w1=1.0, w2=1.0,
                             penalty_new=0.0):
    """
    Same as build_master(), but only enforces cover constraints for active_zlocs.
    Adds a mild extra penalty for *new pickup-only sorties* (routes whose start_load == 0).
    """
    P, S = inst["P"], inst["S"]
    dist = inst["dist"]
    A_truck = inst["A_truck"]
    num_cus = inst["num_cus"]

    total_dem = float((inst["Dd"] + inst["Dp"]).sum())
    max_trucks = max(1, math.ceil(max(1e-9, total_dem) / Q1) * 2)
    max_trucks = min(max_trucks, len(inst["S"]))
    V = list(range(max_trucks))

    m = gp.Model("2E_master_active")
    m.Params.OutputFlag = OUTPUT_LOG
    VTYPE = GRB.CONTINUOUS if as_lp else GRB.BINARY

    x = m.addVars(A_truck, V, lb=0, ub=1, vtype=VTYPE, name="x")
    tau = m.addVars(V, lb=0, ub=1, vtype=VTYPE, name="tau")
    u_tr = m.addVars(S, V, lb=0.0, ub=len(S), vtype=GRB.CONTINUOUS, name="u_tr")
    t_unload = m.addVars(S, V, lb=0.0, ub=Q1, vtype=GRB.CONTINUOUS, name="t_unload")
    y_dep = m.addVars(P, V, vtype=VTYPE, name="y_dep")

    gamma = {}
    for s in S:
        for r in range(len(routes_by_s[s])):
            gamma[(s, r)] = m.addVar(lb=0, ub=1, vtype=VTYPE, name=f"gamma[{s},{r}]")
    m.update()

    # ---- Satellite loading-queue constraints (optional) ----
    _add_loading_queue_constraints(m, inst, routes_by_s, gamma, deliv_mat)

    # objective
    # base: h1 trucks + h2 drone sorties
    # extra: penalty_new * pickup-only sortie (start_load==0)
    extra_expr = gp.quicksum((float(penalty_new) if int(deliv_mat[(s, r)]) == 0 else 0.0) * gamma[(s, r)]
                             for s in S for r in range(len(routes_by_s[s])))

    obj1_expr = gp.quicksum(h1 * tau[v] for v in V) \
                + gp.quicksum(h2 * gamma[(s, r)] for s in S for r in range(len(routes_by_s[s]))) \
                + extra_expr

    obj2_expr = gp.quicksum(c_v * dist[i, j] * x[i, j, v] for (i, j) in A_truck for v in V) \
                + gp.quicksum(c_u * energy_by_sr[(s, r)] * gamma[(s, r)]
                              for s in S for r in range(len(routes_by_s[s])))

    m.setObjective(float(w1) * obj1_expr + float(w2) * obj2_expr, GRB.MINIMIZE)

    # truck constraints (same)
    for v in V:
        m.addConstr(gp.quicksum(x[i, j, v] for i in P for j in S if (i, j) in A_truck) == tau[v], name=f"dep_v{v}")
        m.addConstr(gp.quicksum(x[i, j, v] for i in S for j in P if (i, j) in A_truck) == tau[v], name=f"ret_v{v}")

    for v in V:
        m.addConstr(gp.quicksum(y_dep[p, v] for p in P) == tau[v], name=f"one_dep_v{v}")
        for p in P:
            m.addConstr(gp.quicksum(x[p, s, v] for s in S if (p, s) in A_truck) == y_dep[p, v],
                        name=f"start_at_p{p}_v{v}")
            m.addConstr(gp.quicksum(x[s, p, v] for s in S if (s, p) in A_truck) == y_dep[p, v],
                        name=f"return_to_p{p}_v{v}")

    for v in V:
        for s in S:
            m.addConstr(gp.quicksum(x[i, s, v] for i in (P + S) if (i, s) in A_truck) ==
                        gp.quicksum(x[s, j, v] for j in (P + S) if (s, j) in A_truck),
                        name=f"flow_s{s}_v{v}")

    if len(S) >= 2:
        T = len(S)
        for v in V:
            for i in S:
                for j in S:
                    if i != j and (i, j) in A_truck:
                        m.addConstr(u_tr[i, v] - u_tr[j, v] + T * x[i, j, v] <= T - 1,
                                    name=f"mtz_s{i}_s{j}_v{v}")

    for v in V:
        for (i, j) in A_truck:
            m.addConstr(x[i, j, v] <= tau[v], name=f"use_x_tau_{i}_{j}_v{v}")

    # link: drone route only if satellite visited
    for s in S:
        visit_s = gp.quicksum(x[i, s, v] for v in V for i in (P + S) if (i, s) in A_truck)
        for r in range(len(routes_by_s[s])):
            m.addConstr(gamma[(s, r)] <= visit_s, name=f"route_visit_link_s{s}_r{r}")

    # cover constraints (active only)
    cover_constr = {}
    for zloc in active_zlocs:
        expr = gp.quicksum(gamma[(s, r)] for s in S for r in range(len(routes_by_s[s]))
                           if zloc in cover_mat[(s, r)])
        cc = m.addConstr(expr == 1, name=f"cover_z{zloc}")
        cover_constr[zloc] = cc

    # unload & supply constraints (same)
    for s in S:
        for v in V:
            m.addConstr(t_unload[s, v] <= Q1 * gp.quicksum(x[i, s, v] for i in (P + S) if (i, s) in A_truck),
                        name=f"unload_visit_s{s}_v{v}")
    for v in V:
        m.addConstr(gp.quicksum(t_unload[s, v] for s in S) <= Q1 * tau[v], name=f"unload_cap_v{v}")
    supply_constr = {}
    for s in S:
        rhs = gp.quicksum(deliv_mat[(s, r)] * gamma[(s, r)] for r in range(len(routes_by_s[s])))
        sc = m.addConstr(gp.quicksum(t_unload[s, v] for v in V) >= rhs, name=f"supply_s{s}")
        supply_constr[s] = sc

    for v in range(len(V) - 1):
        m.addConstr(tau[v] >= tau[v + 1], name=f"sym_tau_{v}")

    return m, gamma, cover_constr, supply_constr



def _get_selected_drone_routes(m, inst, routes_by_s):
    """Extract selected routes as list of dicts: {start_sat, seq, tB, tD} (tB/tD in hours if enabled)."""
    S = inst["S"]
    out = []
    for s in S:
        for r in range(len(routes_by_s[s])):
            gv = m.getVarByName(f"gamma[{s},{r}]")
            if gv is None:
                continue
            if gv.X > 0.5:
                item = {"start_sat": int(s), "seq": [int(z) for z in routes_by_s[s][r]]}
                tb = m.getVarByName(f"tB[{s},{r}]")
                td = m.getVarByName(f"tD[{s},{r}]")
                if tb is not None:
                    item["tB"] = float(tb.X)
                if td is not None:
                    item["tD"] = float(td.X)
                out.append(item)
    return out

def _build_best_suffix_order(start_node, start_time_h, start_load, start_E_used, nodes_set, inst,
                             energy_cap=DR['E'], max_enum=200000, ground_wait_free_first=False):
    """
    Exhaustive DFS to find a minimum-energy feasible visiting order for nodes_set,
    starting from (start_node, start_time_h, start_load, start_E_used).

    Important:
      - If ground_wait_free_first=True and start_node is a satellite/ground node, the first-leg waiting
        can be handled by delaying departure (no hover energy). Subsequent legs always use hover-wait.
      - The returned sim dict contains a FULL timeline for the whole suffix (best_seq), not an empty timeline.

    Returns (ok, best_seq, best_sim_dict).
    Safety: abort search after max_enum node-expansions.
    """
    nodes = [int(z) for z in nodes_set]
    if len(nodes) == 0:
        sim = _simulate_route_from_state(int(start_node), float(start_time_h), float(start_load), float(start_E_used),
                                         [],
                                         inst, ground_wait_free=False, remaining_energy_cap=energy_cap)
        return (bool(sim.get("ok", False)), [], sim)

    dist = inst["dist"]
    nodes_sorted0 = sorted(nodes, key=lambda z: float(dist[int(start_node), z]))

    best = {"E": float("inf"), "seq": None}
    expansions = 0

    cache = {}

    def key_state(prev, rem, t_h, load, E_used):
        # coarse rounding to improve pruning
        t_m = int(round(t_h * 60.0))  # minutes
        load10 = int(round(load * 10.0))  # 0.1
        E10 = int(round(E_used * 10.0))  # 0.1
        return (int(prev), rem, t_m, load10, E10)

    def dfs(prev, t_h, load, E_used, rem_set, seq_prefix):
        nonlocal expansions
        if expansions >= max_enum:
            return
        expansions += 1

        rem_fz = frozenset(rem_set)
        k = key_state(prev, rem_fz, t_h, load, E_used)
        old = cache.get(k, None)
        if old is not None and E_used >= old - 1e-9:
            return
        cache[k] = E_used

        # prune by current used energy
        if E_used >= best["E"] - 1e-9:
            return

        if len(rem_set) == 0:
            # only return-to-nearest-sat cost remains
            sim_ret = _simulate_route_from_state(int(prev), float(t_h), float(load), float(E_used), [],
                                                 inst, ground_wait_free=False, remaining_energy_cap=energy_cap)
            if sim_ret.get("ok", False):
                E_total = float(sim_ret["end_energy_used"])
                if E_total < best["E"] - 1e-9:
                    best["E"] = E_total
                    best["seq"] = list(seq_prefix)
            return

        cand = sorted(list(rem_set), key=lambda z: float(dist[int(prev), z]))
        for z in cand:
            # For the very first leg (from ground/satellite), optionally allow ground-wait-free departure.
            is_first_leg = (len(seq_prefix) == 0) and (int(prev) == int(start_node)) and bool(
                ground_wait_free_first) and (_is_customer(int(prev), inst) is False)
            sim1 = _simulate_route_from_state(int(prev), float(t_h), float(load), float(E_used), [int(z)], inst,
                                              ground_wait_free=bool(is_first_leg), remaining_energy_cap=energy_cap)
            if not sim1.get("ok", False):
                continue

            tl = sim1["timeline"][-1]
            new_prev = int(z)
            new_t = float(tl["t_end_h"])
            new_load = float(tl["load_after"])
            new_E = float(tl["E_used"])

            if new_E >= best["E"] - 1e-9:
                continue

            new_rem = set(rem_set)
            new_rem.remove(int(z))
            dfs(new_prev, new_t, new_load, new_E, new_rem, seq_prefix + [int(z)])

    dfs(int(start_node), float(start_time_h), float(start_load), float(start_E_used), set(nodes_sorted0), [])

    if best["seq"] is None:
        return (False, [], {"ok": False, "expansions": expansions})

    # Re-simulate the FULL suffix to obtain a complete timeline and consistent end info.
    best_sim = _simulate_route_from_state(int(start_node), float(start_time_h), float(start_load), float(start_E_used),
                                          [int(z) for z in best["seq"]], inst,
                                          ground_wait_free=bool(ground_wait_free_first),
                                          remaining_energy_cap=energy_cap)
    if not best_sim.get("ok", False):
        return (False, [], {"ok": False, "expansions": expansions})

    best_sim["expansions"] = expansions
    return (True, best["seq"], best_sim)



def _find_best_new_pickup_only_sortie(remaining_pick_nodes, inst, t_dec_h, max_suffix_enum,
                                     exact_subset_limit=10, max_subset_trials_per_k=2000,
                                     allowed_start_sats=None):
    """
    Find ONE best new pickup-only sortie for the current stage.

    Goal: maximize the number of newly released pickup customers served by this one new drone route;
    tie-break by lower ending energy usage. This is used repeatedly so that multiple new drones can
    be dispatched until all newly released pickups are covered or no further feasible new sortie exists.

    Returns None if no feasible pickup-only route can cover any remaining pickup node.
    If allowed_start_sats is given, new sorties can only depart from those active satellites.
    Otherwise returns a dict with keys:
      start_sat, seq, sim, covered_nodes, end_energy_used
    """
    rem = sorted(set(int(z) for z in remaining_pick_nodes))
    if not rem:
        return None

    if allowed_start_sats is None:
        start_sat_candidates = [int(s) for s in inst['S']]
    else:
        start_sat_candidates = sorted(set(int(s) for s in allowed_start_sats if int(s) in set(inst['S'])))
        if not start_sat_candidates:
            return None

    tw_end = inst.get('tw_end', None)
    dist = inst['dist']

    def _better(cand, best):
        if cand is None:
            return False
        if best is None:
            return True
        if len(cand['covered_nodes']) != len(best['covered_nodes']):
            return len(cand['covered_nodes']) > len(best['covered_nodes'])
        if cand['end_energy_used'] < best['end_energy_used'] - 1e-9:
            return True
        if abs(cand['end_energy_used'] - best['end_energy_used']) <= 1e-9:
            return tuple(cand['seq']) < tuple(best['seq'])
        return False

    def _eval_subset(s0, subset_nodes):
        ok, seq_best, sim_best = _build_best_suffix_order(
            start_node=int(s0),
            start_time_h=float(t_dec_h),
            start_load=0.0,
            start_E_used=0.0,
            nodes_set=set(int(z) for z in subset_nodes),
            inst=inst,
            energy_cap=DR['E'],
            max_enum=max_suffix_enum,
            ground_wait_free_first=True
        )
        if not ok:
            return None
        return {
            'start_sat': int(s0),
            'seq': [int(z) for z in seq_best],
            'sim': sim_best,
            'covered_nodes': set(int(z) for z in subset_nodes),
            'end_energy_used': float(sim_best['end_energy_used'])
        }

    best_global = None

    for s0 in start_sat_candidates:
        # First try: one sortie covers all remaining pickups.
        cand_all = _eval_subset(s0, rem)
        if _better(cand_all, best_global):
            best_global = cand_all
        if cand_all is not None:
            continue

        local_best = None
        n = len(rem)

        # Exact subset search for modest number of remaining pickups.
        if n <= exact_subset_limit:
            for k in range(n - 1, 0, -1):
                trials = 0
                for subset in combinations(rem, k):
                    trials += 1
                    if trials > max_subset_trials_per_k:
                        break
                    cand = _eval_subset(s0, subset)
                    if _better(cand, local_best):
                        local_best = cand
                if local_best is not None and len(local_best['covered_nodes']) == k:
                    break
        else:
            # Greedy fallback for larger batches: build a maximal feasible subset.
            order_pool = []
            order_pool.append(sorted(rem, key=lambda z: float(dist[int(s0), int(z)])))
            if tw_end is not None and len(tw_end) > 0:
                order_pool.append(sorted(rem, key=lambda z: float(tw_end[_zloc_from_node(int(z), inst)])))
                order_pool.append(sorted(rem, key=lambda z: (float(tw_end[_zloc_from_node(int(z), inst)]),
                                                            float(dist[int(s0), int(z)]))))
            else:
                order_pool.append(list(rem))

            seen_orders = set()
            unique_orders = []
            for od in order_pool:
                key = tuple(od)
                if key not in seen_orders:
                    seen_orders.add(key)
                    unique_orders.append(od)

            for order in unique_orders:
                chosen = []
                best_chosen = None
                skipped = []
                for z in order:
                    trial = chosen + [int(z)]
                    cand = _eval_subset(s0, trial)
                    if cand is not None:
                        chosen = trial
                        best_chosen = cand
                    else:
                        skipped.append(int(z))
                # second pass: after route structure changes, retry skipped pickups once
                if best_chosen is not None and skipped:
                    for z in skipped:
                        trial = chosen + [int(z)]
                        cand = _eval_subset(s0, trial)
                        if cand is not None:
                            chosen = trial
                            best_chosen = cand
                if _better(best_chosen, local_best):
                    local_best = best_chosen

        if _better(local_best, best_global):
            best_global = local_best

    return best_global

def solve_instance_dynamic_pickup(csv_path, w1=5.0, w2=5.0, penalty_new_factor=2.0,
                                  max_suffix_enum=200000, verbose=True,
                                  merge_release_window_min=3.0,
                                  lock_unaffected=False,
                                  score_slack_weight=200.0,
                                  score_emargin_weight=1000.0,
                                  score_lmargin_weight=20.0):
    """
    Main entry for your dynamic-pickup requirement.
    Returns a list of per-stage result dicts (each dict is a table row).
    """
    t_all0 = time.perf_counter()
    inst = read_instance(csv_path)

    # releases
    release_min = _compute_release_minutes(inst)

    # scoring helpers (used when choosing which in-route drone should take a newly released pickup)
    tw_end_arr = inst.get("tw_end", None)

    def _min_slack_minutes_from_tl(tl):
        if tl is None or len(tl) == 0:
            return float("inf")
        if tw_end_arr is None or len(tw_end_arr) == 0:
            return float("inf")
        ms = float("inf")
        for rec in tl:
            try:
                node = int(rec.get("node"))
                zloc = _zloc_from_node(node, inst)
                slack = float(tw_end_arr[zloc]) - float(rec.get("t_start_h", 0.0)) * 60.0
                if slack < ms:
                    ms = slack
            except Exception:
                continue
        return float(ms)

    def _load_margin_from_tl(tl, start_load=0.0):
        peak = float(start_load)
        if tl is not None:
            for rec in tl:
                try:
                    peak = max(peak, float(rec.get("load_before", peak)), float(rec.get("load_after", peak)))
                except Exception:
                    continue
        return float(max(0.0, float(Q2) - peak))

    # stage0 active set (deliveries + pickups release==0)
    active0 = _active_zlocs_stage0(inst, release_min)
    inst0 = _make_inst_with_active_Z(inst, active0)

    # Stage0 solve (Gurobi)
    t0 = time.perf_counter()
    # -------- helper: build a consistent one-row output for early failures (include solve time) --------
    def _early_fail_row(stage_id: int, status: str, t_start: float, **extra):
        row = {
            "Instances": os.path.basename(csv_path),
            "Stage": int(stage_id),
            "TriggerRelease(min)": "",
            "ReplanTime(min)": "",
            "NewPickups(zloc)": "",
            "First-echelon": "",
            "Second-echelon": "",
            "DecisionTimes(min)": "",
            "NumSorties": "",
            "Obj1": "",
            "Obj2": "",
            "Sum_obj": "",
            "PenaltyNew": "",
            "NewSortieCreated": "",
            "DroneEnergy(Wh)": "",
            "ReplanSolveTime(s)": round(time.perf_counter() - float(t_start), 3),
            "Status": str(status),
            "Stage0_ObjVal": "",
            "Stage0_ObjBound": "",
            "Stage0_MIPGap": "",
            "Stage0_ProvenOptimal": 0,
        }
        if extra:
            row.update(extra)
        return [row]

    if STAGE0_REQUIRE_GLOBAL_OPT:
        routes_by_s, energy_by_sr, cover_mat, deliv_mat = build_pool_exact_global(inst0)
    else:
        routes_by_s, energy_by_sr, cover_mat, deliv_mat = build_pool_exact(inst0)
    pool_cols = sum(len(routes_by_s[s]) for s in routes_by_s)
    if verbose:
        print(f"[stage0-pool] active customers={len(active0)}, pool_cols={pool_cols}, global_opt={STAGE0_REQUIRE_GLOBAL_OPT}")

    miss = warn_uncovered_customers_active(inst0, cover_mat, routes_by_s, active0)
    if miss:
        return _early_fail_row(0, "Unsolvable(pool)", t0)

    PenaltyNew = float(penalty_new_factor) * float(h2)

    m, gamma, cover_constr, supply_constr = build_master_with_active(
        inst0, routes_by_s, energy_by_sr, cover_mat, deliv_mat,
        active_zlocs=active0, as_lp=True, w1=w1, w2=w2, penalty_new=PenaltyNew
    )
    m.Params.Method = 1
    m.Params.OutputFlag = OUTPUT_LOG
    if STAGE0_LP_TIME_LIMIT is not None:
        m.Params.TimeLimit = float(STAGE0_LP_TIME_LIMIT)
    m.optimize()
    if m.Status != GRB.OPTIMAL:
        return _early_fail_row(0, f"Stage0-LP-NotOptimal({int(m.Status)})", t0)
    lp_val = float(m.ObjVal)

    # binarize
    for v in m.getVars():
        n = v.VarName
        if (n.startswith("gamma[") or n.startswith("x[") or
                n.startswith("tau[") or n.startswith("y_dep[")):
            v.VType = GRB.BINARY
    m.update()

    m.Params.Presolve = 2
    m.Params.Cuts = 2
    m.Params.Heuristics = 0.2
    m.Params.MIPFocus = 3
    m.Params.MIPGap = 0.0
    if STAGE0_MIP_TIME_LIMIT is not None:
        m.Params.TimeLimit = float(STAGE0_MIP_TIME_LIMIT)
    m.optimize(mip_progress_cb)

    if m.SolCount == 0 or m.Status in (GRB.INFEASIBLE, GRB.INF_OR_UNBD):
        return _early_fail_row(0, "Unsolvable(MIP)", t0)
    if m.Status != GRB.OPTIMAL or abs(float(m.ObjBound) - float(m.ObjVal)) > 1e-6:
        return _early_fail_row(0, f"Stage0-NotProvenOptimal({int(m.Status)})", t0,
                               Stage0_ObjVal=float(getattr(m, 'ObjVal', float('nan'))),
                               Stage0_ObjBound=float(getattr(m, 'ObjBound', float('nan'))),
                               Stage0_MIPGap=float(getattr(m, 'MIPGap', float('nan'))),
                               Stage0_ProvenOptimal=0)

    first_str, second_str = extract_paths(m, inst0, routes_by_s)

    # Extract drone entities
    drones = _get_selected_drone_routes(m, inst0, routes_by_s)
    if verbose:
        print(f"[stage0] selected drones={len(drones)}")

    # Format stage0 second-echelon routes with clear separators
    second_routes0 = []
    for d in drones:
        s0 = int(d["start_sat"])
        seq0 = [int(z) for z in d["seq"]]
        end_sat0 = int(_nearest_sat_from(seq0[-1], inst0)[0]) if len(seq0) > 0 else s0
        second_routes0.append(_format_route_12(s0, seq0, end_sat0, inst0))
    second_str_fmt0 = " | ".join(second_routes0)

    # Build per-drone ownership (deliveries fixed) and stage0 simulation timeline
    num_dep, num_sat = inst["num_dep"], inst["num_sat"]
    Dd, Dp = inst["Dd"], inst["Dp"]
    drone_states = []
    for didx, d in enumerate(drones):
        start_sat = int(d["start_sat"])
        seq = [int(z) for z in d["seq"]]
        sim = _simulate_full_route_from_sat(start_sat, seq, inst, start_time_h=float(d.get('tD', 0.0)), ground_wait_free_first=False)  # use full inst for proper zloc arrays
        if not sim.get("ok", False):
            # should not happen if stage0 used same feasibility, but guard
            return _early_fail_row(0, f"Stage0RouteSimFail(d{didx})", t0)

        deliv_owned = set()
        pickup_owned = set()
        for z in seq:
            zloc = _zloc_from_node(z, inst)
            if int(Dd[zloc]) > 0:
                deliv_owned.add(int(z))
            if int(Dp[zloc]) > 0:
                pickup_owned.add(int(z))

        # timeline from sim
        tl = sim["timeline"]
        drone_states.append({
            "id": didx,
            "start_sat": start_sat,
            "prefix_seq": [],  # executed prefix customers (global ids)
            "plan_seq": seq,  # current full planned customers (global ids)
            "timeline": tl,  # list per visited customer in plan_seq
            "deliv_owned": deliv_owned,
            "pickup_assigned": set(pickup_owned),  # pickups already in plan
            "created_stage": 0,
            "end_energy_used": float(sim.get("end_energy_used", 0.0)),
            "end_time_h": float(sim.get("end_time_h", 0.0)),
            "end_sat": int(sim.get("end_sat", start_sat)),
            "last_cus_end_h": float(tl[-1]["t_end_h"]) if tl else 0.0,
        })

    # Pending pickups (release>0)
    pending = []
    for zloc in range(inst["num_cus"]):
        if int(Dp[zloc]) > 0 and float(release_min[zloc]) > 1e-9:
            pending.append((float(release_min[zloc]), int(zloc)))
    pending.sort(key=lambda x: x[0])

    # Global pickup tracking for persistence across stages
    # - active_pick_nodes: all pickup customer nodes whose request has been released (including known-at-start)
    # - served_pick_nodes: pickup customer nodes already completed (t_end <= current decision time)
    active_pick_nodes = set()
    served_pick_nodes = set()
    for _zloc in range(inst["num_cus"]):
        if int(Dp[_zloc]) > 0 and float(release_min[_zloc]) <= 1e-9:
            active_pick_nodes.add(int(inst["num_dep"] + inst["num_sat"] + _zloc))

    # Stage0 objective recompute (same style as original, but on inst0 solution)
    # We will output cumulative objective per stage for drone part + keep stage0 truck+drone obj if needed.
    # For dynamic stages, truck part is kept constant as stage0.
    # Here we reuse the extraction from original solve_instance to compute obj1/obj2.
    dist = inst0["dist"]
    obj1_0 = 0.0
    obj2_0 = 0.0
    truck_fixed0 = 0.0
    truck_var0 = 0.0
    for var in m.getVars():
        name = var.VarName
        val = var.X
        if abs(val) < 1e-8:
            continue
        if name.startswith("x["):
            inside = name[2:-1]
            i_str, j_str, v_str = inside.split(",")
            i = int(i_str);
            j = int(j_str)
            obj2_0 += c_v * dist[i, j] * val
            truck_var0 += c_v * dist[i, j] * val
        elif name.startswith("gamma["):
            inside = name[6:-1]
            s_str, r_str = inside.split(",")
            s = int(s_str);
            r = int(r_str)
            E = energy_by_sr.get((s, r), 0.0)
            obj2_0 += c_u * E * val
            obj1_0 += h2 * val
            if int(deliv_mat[(s, r)]) == 0:
                obj1_0 += PenaltyNew * val
        elif name.startswith("tau["):
            obj1_0 += h1 * val
            truck_fixed0 += h1 * val

    stage_rows = []
    stage_rows.append({
        "Instances": os.path.basename(csv_path),
        "Stage": 0,
        "TriggerRelease(min)": 0,
        "ReplanTime(min)": 0,
        "NewPickups(zloc)": "",
        "First-echelon": first_str,
        "Second-echelon": second_str_fmt0,
        "DecisionTimes(min)": "",
        "NumSorties": len(drone_states),
        "Obj1": round(obj1_0, 6),
        "Obj2": round(obj2_0, 6),
        "Sum_obj": round(obj1_0 + obj2_0, 6),
        "PenaltyNew": round(PenaltyNew, 6),
        "NewSortieCreated": 0,
        "DroneEnergy(Wh)": round(sum(float(d.get("end_energy_used", 0.0)) for d in drone_states), 6),
        "ReplanSolveTime(s)": round(time.perf_counter() - t0, 3),
        "Status": "OK",
        "Stage0_ObjVal": round(float(m.ObjVal), 6),
        "Stage0_ObjBound": round(float(m.ObjBound), 6),
        "Stage0_MIPGap": round(float(getattr(m, "MIPGap", 0.0)), 12),
        "Stage0_ProvenOptimal": 1
    })

    # Helper: whether a global node is a pickup-customer node in this instance
    def _is_pick_node(node):
        try:
            node = int(node)
        except Exception:
            return False
        if not _is_customer(node, inst):
            return False
        zloc = _zloc_from_node(node, inst)
        return int(Dp[zloc]) > 0

    # Helper: first departure time from start_sat to first customer (accounting for ground-wait-free delay)
    def _first_departure_time_h(ds):
        tl = ds.get("timeline", [])
        seq = ds.get("plan_seq", [])
        if (not tl) or (not seq):
            return float("inf")
        first_node = int(seq[0])
        t_arr = float(tl[0].get("t_arrive_h", 0.0))
        t_fly = float(inst["dist"][int(ds["start_sat"]), first_node]) / max(1e-9, DR['f_s'])
        return float(t_arr - t_fly)

    # Helper to find decision point index for a drone given a time t_dec_h
    # Returns:
    #   -1 if the drone has not departed yet at t_dec_h (so no prefix is frozen),
    #   otherwise i such that timeline[i]["t_end_h"] is the first completion >= t_dec_h.
    def _decision_index(dr_state, t_dec_h):
        dep1 = _first_departure_time_h(dr_state)
        if float(t_dec_h) < float(dep1) - 1e-12:
            return -1
        tl = dr_state.get("timeline", [])
        for i, rec in enumerate(tl):
            if float(rec["t_end_h"]) >= float(t_dec_h) - 1e-12:
                return i
        return len(tl) - 1  # already finished all customers -> freeze all

    # Helper: compute earliest global decision time among all drones after a given release time
    def _next_global_decision_time_h(t_rel_h):
        # Global stage decision time: the earliest "replan-eligible" time >= release.
        # - If a drone has not departed yet at t_rel_h, it can replan immediately at t_rel_h.
        # - Otherwise, it can only replan right after completing the next customer (first t_end >= t_rel_h).
        times = []
        for ds in drone_states:
            dep1 = _first_departure_time_h(ds)
            if float(t_rel_h) < float(dep1) - 1e-12:
                times.append(float(t_rel_h))
                continue
            tl = ds.get("timeline", [])
            decided = None
            for rec in tl:
                t_end = float(rec["t_end_h"])
                if t_end >= float(t_rel_h) - 1e-12:
                    decided = t_end
                    break
            if decided is None:
                decided = float(t_rel_h)
            times.append(float(decided))
        if not times:
            return float(t_rel_h)
        return float(min(times))

    # Helper: per-drone next decision time (in minutes) given a stage decision time t_dec_h.
    # If the drone has not departed by t_dec_h, it can replan immediately at t_dec_h.
    # Otherwise, it can replan only after completing the next customer (first t_end >= t_dec_h).
    def _next_decision_time_min(dr_state, t_dec_h):
        dep1 = _first_departure_time_h(dr_state)
        if float(t_dec_h) < float(dep1) - 1e-12:
            return float(t_dec_h) * 60.0
        tl = dr_state.get("timeline", [])
        for rec in tl:
            t_end = float(rec.get("t_end_h", rec.get("t_end", 0.0)))
            if t_end >= float(t_dec_h) - 1e-12:
                return float(t_end) * 60.0
        # If no timeline (or already finished), fall back to last completion time (or current t_dec_h)
        last_end = float(dr_state.get("last_cus_end_h", t_dec_h))
        return float(max(last_end, float(t_dec_h))) * 60.0

    # Dynamic stages loop
    stage_id = 0
    while pending:
        stage_id += 1
        next_rel_min, _ = pending[0]
        # Optional: merge nearby releases to reduce frequent replans (delay planning slightly).
        target_rel_min = float(next_rel_min)
        if float(merge_release_window_min) > 1e-9:
            target_rel_min = float(next_rel_min) + float(merge_release_window_min)
        t_rel_h = float(target_rel_min) / 60.0

        # global decision time: earliest customer-completion >= (possibly delayed) release
        t_dec_h = _next_global_decision_time_h(t_rel_h)
        t_dec_min = int(round(t_dec_h * 60.0))

        # activate all pickups with release <= t_dec
        new_pickups_zloc = []
        t_dec_min_f = float(t_dec_h) * 60.0
        while pending and pending[0][0] <= t_dec_min_f + 1e-9:
            rel_m, zloc = pending.pop(0)
            new_pickups_zloc.append(int(zloc))

        if verbose:
            print(f"[stage{stage_id}] t_dec={t_dec_min} min, activate pickups zloc={new_pickups_zloc}")

        # Add newly activated pickups into global active set
        for _zloc in new_pickups_zloc:
            active_pick_nodes.add(int(inst["num_dep"] + inst["num_sat"] + int(_zloc)))

        # Update served pickups set up to current decision time (completed by t_dec)
        for _ds in drone_states:
            for _rec in _ds.get("timeline", []):
                try:
                    _tend = float(_rec.get("t_end_h", 0.0))
                except Exception:
                    continue
                if _tend <= float(t_dec_h) + 1e-12:
                    _n = int(_rec.get("node"))
                    if _is_pick_node(_n):
                        served_pick_nodes.add(_n)

        t_stage_start = time.perf_counter()

        # Build per-drone frozen prefix and current states at their own decision points >= t_dec
        # Also determine remaining owned deliveries / assigned pickups not yet served.
        per_drone_state = []

        # Active (released & not yet completed) pickup nodes at this stage
        active_unserved_pick = set(active_pick_nodes) - set(served_pick_nodes)

        # Enforce global uniqueness of pickup assignments:
        # - Pickups already committed to some drone in earlier stages remain with that drone (unless already served).
        # - Pickups in a frozen prefix at this decision time cannot be reassigned.
        pickup_owner = {}  # pick_node -> drone_id
        fixed_pick_owner = {}  # pick_node -> drone_id (appears in frozen prefix)

        # First pass: compute each drone's frozen prefix/state and collect fixed pickups
        drone_tmp = []  # list of dicts with computed prefix/state for each ACTIVE drone
        for ds in drone_states:
            tl_all = ds.get("timeline", [])
            # Single-sortie assumption: if the drone has already completed ALL planned customers by t_dec,
            # we do NOT allow it to take new pickups (ignore idle/returning drones; new sorties will be created instead).
            if tl_all:
                last_cus_end = float(tl_all[-1].get("t_end_h", 0.0))
                if last_cus_end <= float(t_dec_h) + 1e-12:
                    continue

            di = _decision_index(ds, t_dec_h)

            if di < 0:
                # Not departed yet at t_dec: no frozen prefix; state stays at start satellite at time t_dec.
                prefix_seq = []
                prefix_tl = []
                cur_node = int(ds["start_sat"])
                cur_time = float(t_dec_h)
                # Guard: if the drone has not departed yet, it cannot take off earlier than the planned Stage-0 takeoff time (includes loading-queue delay)
                cur_time = max(cur_time, float(ds.get("tD", cur_time)))
                cur_E = 0.0
                # Initial load = all deliveries owned by this drone (it is already loaded at departure)
                cur_load = 0.0
                for z in ds.get("deliv_owned", set()):
                    zloc = _zloc_from_node(int(z), inst)
                    cur_load += float(Dd[zloc])
            else:
                # Frozen prefix includes the first customer completion >= t_dec
                prefix_seq = [int(z) for z in ds["plan_seq"][:di + 1]]
                prefix_tl = ds["timeline"][:di + 1]
                last = prefix_tl[-1]
                cur_node = int(last["node"])
                cur_time = float(last["t_end_h"])
                cur_load = float(last["load_after"])
                cur_E = float(last["E_used"])
                if cur_time < float(t_dec_h) - 1e-12:
                    cur_time = float(t_dec_h)

            served_nodes = set(prefix_seq)

            # Pickups in the frozen prefix are fixed to this drone
            for _n in prefix_seq:
                if _is_pick_node(int(_n)):
                    fixed_pick_owner[int(_n)] = int(ds["id"])

            drone_tmp.append({
                "ds": ds,
                "di": di,
                "prefix_seq": prefix_seq,
                "prefix_timeline": prefix_tl,
                "served_nodes": served_nodes,
                "cur_node": cur_node,
                "cur_time_h": cur_time,
                "cur_load": cur_load,
                "cur_E": cur_E,
            })

        # Initialize pickup owner with fixed pickups
        pickup_owner.update(fixed_pick_owner)

        # Second pass: keep previously assigned pickups (active & unserved) with their last owner (stability),
        # but never override fixed prefix ownership.
        for info in drone_tmp:
            ds = info["ds"]
            did = int(ds["id"])
            prev_picks = set(ds.get("pickup_assigned", set())) & set(active_unserved_pick)
            for _n in sorted(prev_picks):
                _n = int(_n)
                if _n in pickup_owner and pickup_owner[_n] != did:
                    # already owned by another drone (likely fixed); skip
                    continue
                pickup_owner.setdefault(_n, did)

        # All owned active pickups are committed and should not be assigned again
        committed_pick_nodes = set(pickup_owner.keys())

        # Build per-drone state using the unique pickup_owner assignment
        for info in drone_tmp:
            ds = info["ds"]
            did = int(ds["id"])
            served_nodes = set(info["served_nodes"])

            rem_deliv = set(ds.get("deliv_owned", set())) - served_nodes

            # Keep only pickups owned by THIS drone (active & unserved), and not already in frozen prefix
            rem_pick = set()
            prev_picks = set(ds.get("pickup_assigned", set())) & set(active_unserved_pick)
            for _n in prev_picks:
                _n = int(_n)
                if pickup_owner.get(_n, None) == did and (_n not in served_nodes):
                    rem_pick.add(_n)

            per_drone_state.append({
                "id": did,
                "created_stage": ds.get("created_stage", 0),
                "start_sat": int(ds["start_sat"]),
                "prefix_seq": info["prefix_seq"],
                "prefix_timeline": info["prefix_timeline"],
                "cur_node": info["cur_node"],
                "cur_time_h": info["cur_time_h"],
                "cur_load": info["cur_load"],
                "cur_E": info["cur_E"],
                "rem_deliv_nodes": rem_deliv,
                "rem_pick_nodes": rem_pick,
                "got_new_pickups": False,
                "old_plan_seq": [int(z) for z in ds.get("plan_seq", [])]
            })
        # DecisionTimes(min): for each drone, when the plan becomes "decided" after t_dec (or done@time if already finished)
        decision_times_parts = []
        for ds in drone_states:
            tl_all = ds.get("timeline", [])
            seq_all = ds.get("plan_seq", [])
            did = ds.get("id", "?")
            if (not tl_all) or (not seq_all):
                decision_times_parts.append(f"d{did}:na")
                continue
            last_cus_end = float(tl_all[-1]["t_end_h"])
            if last_cus_end <= float(t_dec_h) + 1e-12:
                decision_times_parts.append(f"d{did}:done@{int(round(last_cus_end * 60.0))}")
            else:
                di = _decision_index(ds, float(t_dec_h))
                if di < 0:
                    decision_times_parts.append(f"d{did}:{int(round(float(t_dec_h) * 60.0))}")
                else:
                    decision_times_parts.append(f"d{did}:{int(round(float(tl_all[di]['t_end_h']) * 60.0))}")
        decision_times_str = "|".join(decision_times_parts)

        # Build pool of pickups to be assigned at this stage:
        # all released pickups that are not yet completed and not frozen in any drone's committed prefix.
        # Build pool of pickups to be assigned at this stage:
        # all released pickups that are not yet completed and not frozen in any drone's committed prefix.
        # NOTE: use a named key function to avoid IDE "unresolved reference" warnings inside lambda.
        tw_start_arr = inst.get("tw_start", None)

        def _pick_sort_key(n):
            z = _zloc_from_node(int(n), inst)
            tws = float(tw_start_arr[z]) if tw_start_arr is not None else 0.0
            return (float(release_min[z]), tws, int(z))

        pick_nodes_to_assign = sorted(
            list(set(active_pick_nodes) - set(served_pick_nodes) - set(committed_pick_nodes)),
            key=_pick_sort_key
        )

        unassigned_pick_nodes = []

        for z_node in pick_nodes_to_assign:
            z_node = int(z_node)
            # Choose the in-route drone using a *soft* multi-criteria score:
            #   base: incremental energy (deltaE)
            #   + risk terms that discourage tiny TW slack / tiny energy margin / tiny load margin
            best_choice = None  # (score, deltaE, idx, best_seq, best_sim)
            for i, st in enumerate(per_drone_state):
                nodes_set = set(st["rem_deliv_nodes"]) | set(st["rem_pick_nodes"]) | {z_node}
                ok, seq_best, sim_best = _build_best_suffix_order(
                    start_node=st["cur_node"],
                    start_time_h=st["cur_time_h"],
                    start_load=st["cur_load"],
                    start_E_used=st["cur_E"],
                    nodes_set=nodes_set,
                    inst=inst,
                    energy_cap=DR['E'],
                    max_enum=max_suffix_enum,
                    ground_wait_free_first=(not _is_customer(int(st["cur_node"]), inst))
                )
                if not ok:
                    continue

                deltaE = float(sim_best.get("end_energy_used", 0.0)) - float(st["cur_E"])

                # soft risk indicators
                tl_best = sim_best.get("timeline", [])
                slack_min = float(_min_slack_minutes_from_tl(tl_best))
                E_margin = float(max(0.0, float(DR['E']) - float(sim_best.get("end_energy_used", 0.0))))
                load_margin = float(_load_margin_from_tl(tl_best, start_load=float(st.get("cur_load", 0.0))))

                score = float(deltaE)
                # If slack_min is inf (no TW), the term becomes ~0.
                score += float(score_slack_weight) / (float(slack_min) + 1.0)
                score += float(score_emargin_weight) / (float(E_margin) + 1.0)
                score += float(score_lmargin_weight) / (float(load_margin) + 1.0)

                if (best_choice is None) or (score < best_choice[0] - 1e-9) or (
                        abs(score - best_choice[0]) <= 1e-9 and deltaE < best_choice[1] - 1e-9):
                    best_choice = (score, deltaE, i, seq_best, sim_best)

            if best_choice is None:
                unassigned_pick_nodes.append(z_node)
            else:
                _, _, i, _, _ = best_choice
                per_drone_state[i]["rem_pick_nodes"].add(z_node)
                per_drone_state[i]["got_new_pickups"] = True
                # mark as committed to avoid any accidental duplication downstream
                committed_pick_nodes.add(int(z_node))
                try:
                    pickup_owner[int(z_node)] = int(per_drone_state[i]["id"])
                except Exception:
                    pass

        # For each existing drone, solve its suffix order with updated remaining nodes
        updated_drone_plans = []
        total_energy_end = 0.0
        for st in per_drone_state:
            nodes_set = set(st["rem_deliv_nodes"]) | set(st["rem_pick_nodes"])
            # Optional stability lock: if a drone receives no newly-assigned pickups at this stage,
            # keep its remaining order (no intra-route reshuffle) and only re-simulate timing/energy
            # from the current state.
            ok = False
            best_seq = []
            best_sim = {"ok": False}
            if lock_unaffected and (not st.get("got_new_pickups", False)):
                old_seq = [int(z) for z in st.get("old_plan_seq", [])]
                served_now = set(int(z) for z in st.get("prefix_seq", []))
                keep_suffix = [int(z) for z in old_seq if (int(z) not in served_now) and (int(z) in nodes_set)]
                sim_keep = _simulate_route_from_state(
                    start_node=st["cur_node"],
                    start_time_h=st["cur_time_h"],
                    start_load=st["cur_load"],
                    start_energy_used=st["cur_E"],
                    seq_customers=keep_suffix,
                    inst=inst,
                    ground_wait_free=(not _is_customer(int(st["cur_node"]), inst)),
                    remaining_energy_cap=DR['E']
                )
                if sim_keep.get("ok", False):
                    ok = True
                    best_seq = keep_suffix
                    best_sim = sim_keep

            if not ok:
                ok, best_seq, best_sim = _build_best_suffix_order(
                    start_node=st["cur_node"],
                    start_time_h=st["cur_time_h"],
                    start_load=st["cur_load"],
                    start_E_used=st["cur_E"],
                    nodes_set=nodes_set,
                    inst=inst,
                    energy_cap=DR['E'],
                    max_enum=max_suffix_enum,
                    ground_wait_free_first=(not _is_customer(int(st["cur_node"]), inst))
                )
            if not ok:
                # fallback: keep old remaining order (and ignore any uninserted pickups)
                old_seq = [int(z) for z in st.get("old_plan_seq", [])]
                served_now = set(int(z) for z in st.get("prefix_seq", []))
                best_seq = [int(z) for z in old_seq if (int(z) not in served_now) and (int(z) in nodes_set)]
                best_sim = {"ok": False}

            full_seq = st["prefix_seq"] + best_seq

            # Determine end_sat from suffix simulation if ok; else from prefix end
            if best_sim.get("ok", False):
                end_sat = int(
                    best_sim.get("end_sat", _nearest_sat_from(full_seq[-1], inst)[0] if full_seq else st["start_sat"]))
                E_end = float(best_sim["end_energy_used"])
                # stitch timeline: prefix_timeline + suffix_timeline
                suffix_tl = best_sim.get("timeline", [])
                timeline_full = st["prefix_timeline"] + suffix_tl
                end_time_h = float(best_sim.get("end_time_h", st["cur_time_h"]))
                last_cus_end_h = float(timeline_full[-1]["t_end_h"]) if timeline_full else float(st["cur_time_h"])
                # update stored route for next stages
            else:
                # simulate the entire remaining using original order (conservative)
                sim_fallback = _simulate_route_from_state(
                    start_node=st["cur_node"],
                    start_time_h=st["cur_time_h"],
                    start_load=st["cur_load"],
                    start_energy_used=st["cur_E"],
                    seq_customers=[int(z) for z in best_seq],
                    inst=inst,
                    ground_wait_free=False,
                    remaining_energy_cap=DR['E']
                )
                end_sat = int(sim_fallback.get("end_sat", st["start_sat"]))
                E_end = float(sim_fallback.get("end_energy_used", st["cur_E"]))
                timeline_full = st["prefix_timeline"] + sim_fallback.get("timeline", [])
                end_time_h = float(sim_fallback.get("end_time_h", st["cur_time_h"]))
                last_cus_end_h = float(timeline_full[-1]["t_end_h"]) if timeline_full else float(st["cur_time_h"])

            updated_drone_plans.append({
                "id": st["id"],
                "start_sat": st["start_sat"],
                "full_seq": full_seq,
                "end_sat": end_sat,
                "timeline": timeline_full,
                "created_stage": st["created_stage"],
                "end_energy_used": float(E_end),
                "end_time_h": float(end_time_h),
                "last_cus_end_h": float(last_cus_end_h),
            })
            total_energy_end += E_end

        # If unassigned pickups remain, dispatch as many new pickup-only sorties as needed.
        # Each new sortie is constructed so that its first customer does not need hover-waiting
        # (implemented by allowing ground waiting before departure).
        new_drones_created = 0
        new_routes_str = []
        stage_abort = False
        remaining_pick_nodes = sorted(set(int(z) for z in unassigned_pick_nodes))
        while remaining_pick_nodes:
            allowed_start_sats = sorted({int(ds["start_sat"]) for ds in drone_states})
            best_new = _find_best_new_pickup_only_sortie(
                remaining_pick_nodes=remaining_pick_nodes,
                inst=inst,
                t_dec_h=float(t_dec_h),
                max_suffix_enum=max_suffix_enum,
                allowed_start_sats=allowed_start_sats
            )
            if best_new is None:
                # Still has pickups left, but even one additional pickup-only sortie cannot serve any of them.
                stage_rows.append({
                    "Instances": os.path.basename(csv_path),
                    "Stage": stage_id,
                    "TriggerRelease(min)": int(round(next_rel_min)),
                    "ReplanTime(min)": t_dec_min,
                    "NewPickups(zloc)": ",".join(str(int(z) + 1) for z in new_pickups_zloc),
                    "First-echelon": first_str,
                    "Second-echelon": "",
                    "NumSorties": len(updated_drone_plans) + new_drones_created,
                    "Obj1": "",
                    "Obj2": "",
                    "Sum_obj": "",
                    "PenaltyNew": round(PenaltyNew, 6),
                    "ReplanSolveTime(s)": round(time.perf_counter() - t_stage_start, 3),
                    "Status": "UnservedPickups"
                })
                stage_abort = True
                break

            s0 = int(best_new["start_sat"])
            seq_best = [int(z) for z in best_new["seq"]]
            sim_best = best_new["sim"]
            covered_nodes = sorted(set(int(z) for z in best_new["covered_nodes"]))
            E_end = float(best_new["end_energy_used"])
            end_sat = int(sim_best.get("end_sat", _nearest_sat_from(seq_best[-1], inst)[0] if seq_best else s0))

            new_id = max([d["id"] for d in drone_states] + [-1]) + 1
            drone_states.append({
                "id": new_id,
                "start_sat": int(s0),
                "prefix_seq": [],
                "plan_seq": list(seq_best),
                "timeline": sim_best.get("timeline", []),
                "deliv_owned": set(),  # pickup-only
                "pickup_assigned": set(covered_nodes),
                "created_stage": stage_id,
                "end_energy_used": float(E_end),
                "end_time_h": float(sim_best.get("end_time_h", t_dec_h)),
                "end_sat": int(end_sat),
                "last_cus_end_h": float(sim_best.get("timeline", [])[-1].get("t_end_h", t_dec_h)) if sim_best.get("timeline", []) else float(t_dec_h),
            })
            new_drones_created += 1
            new_routes_str.append(_format_route_12(int(s0), seq_best, end_sat, inst))
            total_energy_end += float(E_end)

            covered_set = set(covered_nodes)
            remaining_pick_nodes = [int(z) for z in remaining_pick_nodes if int(z) not in covered_set]

        if stage_abort:
            break

        # Update drone_states' plans/timelines for existing drones
        # Keep the list order stable by matching ids.
        upd_by_id = {d["id"]: d for d in updated_drone_plans}
        for idx in range(len(drone_states)):
            did = drone_states[idx]["id"]
            if did in upd_by_id:
                u = upd_by_id[did]
                drone_states[idx]["plan_seq"] = u["full_seq"]
                # Persist pickups assigned to this drone (for future stages)
                drone_states[idx]["pickup_assigned"] = set(
                    [int(_n) for _n in drone_states[idx]["plan_seq"] if _is_pick_node(_n)])
                drone_states[idx]["timeline"] = u["timeline"]
                drone_states[idx]["end_energy_used"] = float(
                    u.get("end_energy_used", drone_states[idx].get("end_energy_used", 0.0)))
                drone_states[idx]["end_time_h"] = float(u.get("end_time_h", drone_states[idx].get("end_time_h", 0.0)))
                drone_states[idx]["end_sat"] = int(
                    u.get("end_sat", drone_states[idx].get("end_sat", drone_states[idx].get("start_sat", 0))))
                drone_states[idx]["last_cus_end_h"] = float(
                    u.get("last_cus_end_h", drone_states[idx].get("last_cus_end_h", 0.0)))

        # Build stage second-echelon string (all drones including newly created)
        second_routes = []
        for ds in drone_states:
            # if drone finished (no plan_seq), skip
            if len(ds["plan_seq"]) == 0 and len(ds.get("pickup_assigned", set())) == 0:
                continue
            # Determine end_sat (prefer stored simulation result; fall back to nearest-sat heuristic)
            end_sat = int(ds.get("end_sat", ds.get("start_sat", 0)))
            if end_sat <= 0:
                if ds.get("timeline", []):
                    last_node = int(ds["timeline"][-1]["node"])
                    end_sat = int(_nearest_sat_from(last_node, inst)[0])
                else:
                    end_sat = int(ds["start_sat"])
            second_routes.append(_format_route_12(int(ds["start_sat"]), ds["plan_seq"], int(end_sat), inst))
        second_str_stage = " | ".join(second_routes)

        # Objective for stage: keep truck part from stage0 obj1_0/obj2_0? (truck distance fixed)
        # For simplicity and transparency, we report drone-only objective at each stage + keep stage0 total as reference.
        drone_fixed = float(h2) * len(drone_states) + float(PenaltyNew) * sum(
            1 for d in drone_states if d["created_stage"] > 0)
        total_energy_stage = sum(float(d.get("end_energy_used", 0.0)) for d in drone_states)
        drone_var = float(c_u) * float(total_energy_stage)
        obj1_stage = float(truck_fixed0) + float(drone_fixed)
        obj2_stage = float(truck_var0) + float(drone_var)

        # Recompute decision times AFTER any new drone creation / plan updates (for logging)
        decision_times_parts = []
        for i, ds in enumerate(drone_states):
            dkey = f"d{i}"
            if float(ds.get("last_cus_end_h", 0.0)) <= float(t_dec_h) + 1e-12:
                decision_times_parts.append(f"{dkey}:done@{int(round(float(ds.get('last_cus_end_h', 0.0)) * 60))}")
            else:
                dt = _next_decision_time_min(ds, float(t_dec_h))
                decision_times_parts.append(f"{dkey}:{int(round(dt))}")
        decision_times_str = "|".join(decision_times_parts)

        # Sanity check (dynamic pickups):
        # 1) all active (released) pickups must be either already served or present in exactly ONE drone plan
        planned_nodes_all = set()
        pick_count = {}
        for _ds in drone_states:
            seqp = [int(x) for x in _ds.get("plan_seq", [])]
            planned_nodes_all |= set(seqp)
            for _n in seqp:
                if _is_pick_node(int(_n)):
                    pick_count[int(_n)] = pick_count.get(int(_n), 0) + 1

        active_unserved_pick_chk = (set(active_pick_nodes) - set(served_pick_nodes))
        missing_pick_nodes = set(active_unserved_pick_chk) - planned_nodes_all
        dup_pick_nodes = set([n for n, c in pick_count.items() if c > 1 and n in active_unserved_pick_chk])

        if missing_pick_nodes:
            status_str = "MissingActivePickups"
        elif dup_pick_nodes:
            status_str = "DuplicateActivePickups"
        else:
            status_str = "OK"

        stage_rows.append({
            "Instances": os.path.basename(csv_path),
            "Stage": stage_id,
            "TriggerRelease(min)": int(round(next_rel_min)),
            "ReplanTime(min)": t_dec_min,
            "NewPickups(zloc)": ",".join(str(int(z) + 1) for z in new_pickups_zloc),
            "First-echelon": first_str,
            "Second-echelon": second_str_stage,
            "DecisionTimes(min)": decision_times_str,
            "NumSorties": len(drone_states),
            "Obj1": round(obj1_stage, 6),
            "Obj2": round(obj2_stage, 6),
            "Sum_obj": round(obj1_stage + obj2_stage, 6),
            "PenaltyNew": round(PenaltyNew, 6),
            "NewSortieCreated": new_drones_created,
            "DroneEnergy(Wh)": round(float(total_energy_stage), 6),
            "ReplanSolveTime(s)": round(time.perf_counter() - t_stage_start, 3),
            "Status": "OK"
        })

    if verbose:
        print(f"[done] stages={len(stage_rows)}, total wall time={time.perf_counter() - t_all0:.2f}s")

    return stage_rows


def run_dynamic_single(csv_path, out_xlsx=None, w1=5.0, w2=5.0, penalty_new_factor=2.0,
                       max_suffix_enum=200000, lock_unaffected=True):
    rows = solve_instance_dynamic_pickup(
        csv_path=csv_path, w1=w1, w2=w2,
        penalty_new_factor=penalty_new_factor,
        max_suffix_enum=max_suffix_enum,
        lock_unaffected=lock_unaffected,
        verbose=True
    )
    df = pd.DataFrame(rows)
    if out_xlsx is not None:
        safe_save_table(df, out_xlsx)
    return df


def run_batch():
    Set1 = ['Ca1-2,3,15', 'Ca1-3,5,15', 'Ca1-6,4,15', 'Ca2-2,3,15', 'Ca2-3,5,15', 'Ca2-6,4,15',
            'Ca3-2,3,15', 'Ca3-3,5,15', 'Ca3-6,4,15', 'Ca4-2,3,15', 'Ca4-3,5,15', 'Ca4-6,4,15',
            'Ca5-2,3,15', 'Ca5-3,5,15', 'Ca5-6,4,15']
    Set2_1 = ['Ca1-2,3,30', 'Ca1-3,5,30', 'Ca1-6,4,30']
    Set2_2 = ['Ca2-2,3,30', 'Ca2-3,5,30', 'Ca2-6,4,30']
    Set2_3 = ['Ca3-2,3,30', 'Ca3-3,5,30', 'Ca3-6,4,30']
    Set2_4 = ['Ca4-2,3,30', 'Ca4-3,5,30', 'Ca4-6,4,30']
    Set2_5 = ['Ca5-2,3,30', 'Ca5-3,5,30', 'Ca5-6,4,30']
    Set3 = [
        'Ca1-2,3,50', 'Ca1-3,5,50', 'Ca1-6,4,50', 'Ca2-2,3,50', 'Ca2-3,5,50', 'Ca2-6,4,50',
        'Ca3-2,3,50', 'Ca3-3,5,50', 'Ca3-6,4,50', 'Ca4-2,3,50', 'Ca4-3,5,50', 'Ca4-6,4,50',
        'Ca5-2,3,50', 'Ca5-3,5,50', 'Ca5-6,4,50']
    Set4 = [
        'Ca1-2,3,100', 'Ca1-3,5,100', 'Ca1-6,4,100', 'Ca2-2,3,100', 'Ca2-3,5,100', 'Ca2-6,4,100',
        'Ca3-2,3,100', 'Ca3-3,5,100', 'Ca3-6,4,100', 'Ca4-2,3,100', 'Ca4-3,5,100', 'Ca4-6,4,100',
        'Ca5-2,3,100', 'Ca5-3,5,100', 'Ca5-6,4,100', ]

    all_sets = [(1, Set1),(2, Set2_1),(3, Set2_2),(4, Set2_3),(5, Set2_4),(6, Set2_5)]

    for set_id, name_list in all_sets:
        rows = []
        for name in name_list:
            csv_path = os.path.join(BASE_DIR, name + ".csv")
            # Single weight setting: w1=w2=5 (directly sum the two objectives)
            weight_pairs = [(5, 5)]

            for (w1, w2) in weight_pairs:
                try:
                    res = solve_instance(csv_path, w1=w1, w2=w2)
                    if res.get("status") == "OK":
                        rows.append([
                            res["file"],
                            w1, w2,
                            res["time"],
                            res["first"],
                            res["second"],
                            res["obj1"],
                            res["obj2"],
                            res["sum_obj"],
                            res["weighted_obj"],
                            res["weighted_obj_norm"],
                            res["lp"],
                            res["gap"],
                        ])
                    else:
                        rows.append([os.path.basename(csv_path), w1, w2] + ["Unsolvable"] * 10)
                except Exception:
                    rows.append([os.path.basename(csv_path), w1, w2] + ["Unsolvable"] * 10)

        dfout = pd.DataFrame(rows, columns=[
            "Instances", "w1", "w2", "Time(s)", "First-echelon", "Second-echelon",
            "obj1", "obj2", "Sum_obj", "Weighted_obj", "Weighted_obj_norm", "LP Value", "Gap"
        ])
        out_xlsx = os.path.join(Target_DIR, f"WeightedSweep_Set{set_id}_results.xlsx")
        safe_save_table(dfout, out_xlsx)

def dynamic_df_to_summary(df: pd.DataFrame, instance_name: str, w1: float, w2: float) -> dict:
    """
    Convert the detail df returned by run_dynamic_single into ONE summary row.
    Strategy:
      - If df has column 'status', prefer rows with status=='OK'
      - Take the LAST valid row as the final plan (common for rolling replanning)
      - Fallback to last row if no status column
    """
    if df is None or df.empty:
        return {
            "Instances": instance_name, "w1": w1, "w2": w2,
            "Status": "Unsolvable"
        }

    dfx = df.copy()

    # 兼容性：尽量找到“可行/成功”的行
    if "status" in dfx.columns:
        ok_mask = dfx["status"].astype(str).str.upper().eq("OK")
        if ok_mask.any():
            dfx = dfx.loc[ok_mask]

    last = dfx.iloc[-1].to_dict()  # 取最后一个阶段作为汇总

    # 强制补充关键信息列名（不破坏你原df的字段）
    summary = {"Instances": instance_name, "w1": w1, "w2": w2}

    # 你df里有哪些列，就会带哪些；没有就不带
    summary.update(last)

    # 如果没有 Status 字段，也给一个兜底
    if "Status" not in summary and "status" in summary:
        summary["Status"] = summary["status"]

    return summary

def run_dynamic_batch(all_sets, detail_dir_name="DynamicPickup_details",
                      w_pairs=None, penalty_new_factor=2.0, max_suffix_enum=200000, lock_unaffected=True):
    if w_pairs is None:
        w_pairs = [(5.0, 5.0)]

    detail_dir = os.path.join(Target_DIR, detail_dir_name)
    os.makedirs(detail_dir, exist_ok=True)

    for set_id, name_list in all_sets:
        summary_rows = []

        for name in name_list:
            csv_path = os.path.join(BASE_DIR, name + ".csv")

            for (w1, w2) in w_pairs:
                detail_xlsx = os.path.join(detail_dir, f"Dynamic_{name}_RH-BPC.xlsx")

                try:
                    df_detail = run_dynamic_single(
                        csv_path,
                        out_xlsx=detail_xlsx,
                        w1=w1, w2=w2,
                        penalty_new_factor=penalty_new_factor,
                        max_suffix_enum=max_suffix_enum,
                        lock_unaffected=lock_unaffected
                    )

                    summary = dynamic_df_to_summary(df_detail, instance_name=name + ".csv", w1=w1, w2=w2)
                    summary["Detail_xlsx"] = os.path.basename(detail_xlsx)
                    summary_rows.append(summary)

                except Exception as e:
                    summary_rows.append({
                        "Instances": name + ".csv",
                        "w1": w1, "w2": w2,
                        "Status": "Unsolvable",
                        "Error": str(e),
                        "Detail_xlsx": os.path.basename(detail_xlsx)
                    })

        df_sum = pd.DataFrame(summary_rows)
        out_xlsx = os.path.join(Target_DIR, f"DynamicPickup_Set{set_id}_results.xlsx")
        safe_save_table(df_sum, out_xlsx)


def build_sets_for_dynamic():
    Set1 = ['Ca1-2,3,15', 'Ca1-3,5,15', 'Ca1-6,4,15', 'Ca2-2,3,15', 'Ca2-3,5,15', 'Ca2-6,4,15',
            'Ca3-2,3,15', 'Ca3-3,5,15', 'Ca3-6,4,15', 'Ca4-2,3,15', 'Ca4-3,5,15', 'Ca4-6,4,15',
            'Ca5-2,3,15', 'Ca5-3,5,15', 'Ca5-6,4,15']
    Set2 = ['Ca1-2,3,30', 'Ca1-3,5,30', 'Ca1-6,4,30', 'Ca2-2,3,30', 'Ca2-3,5,30', 'Ca2-6,4,30',
            'Ca3-2,3,30', 'Ca3-3,5,30', 'Ca3-6,4,30', 'Ca4-2,3,30', 'Ca4-3,5,30', 'Ca4-6,4,30',
            'Ca5-2,3,30', 'Ca5-3,5,30', 'Ca5-6,4,30']
    Set2_1 = ['Ca1-2,3,30', 'Ca1-3,5,30', 'Ca1-6,4,30']
    Set2_2 = ['Ca2-2,3,30', 'Ca2-3,5,30', 'Ca2-6,4,30']
    Set2_3 = ['Ca3-2,3,30', 'Ca3-3,5,30', 'Ca3-6,4,30']
    Set2_4 = ['Ca4-2,3,30', 'Ca4-3,5,30', 'Ca4-6,4,30']
    Set2_5 = ['Ca5-2,3,30', 'Ca5-3,5,30', 'Ca5-6,4,30']
    Set3 = [
        'Ca1-2,3,50', 'Ca1-3,5,50', 'Ca1-6,4,50', 'Ca2-2,3,50', 'Ca2-3,5,50', 'Ca2-6,4,50',
        'Ca3-2,3,50', 'Ca3-3,5,50', 'Ca3-6,4,50', 'Ca4-2,3,50', 'Ca4-3,5,50', 'Ca4-6,4,50',
        'Ca5-2,3,50', 'Ca5-3,5,50', 'Ca5-6,4,50']
    Set4 = [
        'Ca1-2,3,100', 'Ca1-3,5,100', 'Ca1-6,4,100', 'Ca2-2,3,100', 'Ca2-3,5,100', 'Ca2-6,4,100',
        'Ca3-2,3,100', 'Ca3-3,5,100', 'Ca3-6,4,100', 'Ca4-2,3,100', 'Ca4-3,5,100', 'Ca4-6,4,100',
        'Ca5-2,3,100', 'Ca5-3,5,100', 'Ca5-6,4,100', ]

    all_sets = [(1, Set1),(2, Set2)]
    return all_sets

# ------------------ Quick switch ------------------
DYNAMIC_MODE = False
DYNAMIC_BATCH_MODE = True
DYNAMIC_CSV_NAME = "Ca1-3,5,50.csv"
DYNAMIC_OUT_XLSX = "DynamicPickup_results_Ca1-3,5,50-BPC.xlsx"
if __name__ == "__main__":
    if DYNAMIC_BATCH_MODE:
        all_sets = build_sets_for_dynamic()
        run_dynamic_batch(all_sets)
    elif DYNAMIC_MODE:
        csv_path = os.path.join(BASE_DIR, DYNAMIC_CSV_NAME)
        out_xlsx = os.path.join(Target_DIR, DYNAMIC_OUT_XLSX)
        run_dynamic_single(csv_path, out_xlsx=out_xlsx, w1=5.0, w2=5.0,
                           penalty_new_factor=2.0, max_suffix_enum=200000, lock_unaffected=True)
    else:
        run_batch()
