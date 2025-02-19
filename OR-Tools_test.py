import itertools
from ortools.linear_solver import pywraplp
import pandas as pd
import os
from collections import defaultdict

import scratch

def generate_patterns(stock_length, demand_lengths, kerf_width):
    """生成考虑锯缝的有效切割模式"""
    patterns = []
    max_counts = [stock_length // length for length in demand_lengths]

    # 生成所有可能的切割组合
    for combo in itertools.product(*[range(0, c + 1) for c in max_counts]):
        total_pieces = sum(combo)
        if total_pieces == 0:
            continue

        # 计算总消耗长度（含锯缝）
        total_used = sum(c * l for c, l in zip(combo, demand_lengths))
        total_kerf = kerf_width * (total_pieces - 1) if total_pieces > 1 else 0
        total_consumption = total_used + total_kerf

        if total_consumption <= stock_length:
            utilization = round((total_consumption / stock_length) * 100, 2)
            patterns.append({
                "combo": combo,
                "waste": stock_length - total_consumption,
                "kerf": total_kerf,
                "utilization": utilization  # 新增单个方案利用率
            })
    return patterns


def create_data_model():
    """包含锯缝参数的数据模型"""
    return {
        "kerf_width": 5,  # 默认锯缝宽度改为5mm
        "stock": [
            {"length": 6000, "quantity": 148},
            {"length": 5400, "quantity": 30},
            {"length": 5000, "quantity": 30},
        ],
        "demands": [
            {"length": 2465, "quantity": 45},
            {"length": 2215, "quantity": 75},
            {"length": 2135, "quantity": 55},
            {"length": 2015, "quantity": 165},
            {"length": 1885, "quantity": 70},
            {"length": 1745, "quantity": 122},
            {"length": 986, "quantity": 33},
            {"length": 785, "quantity": 45},
            {"length": 656, "quantity": 25},
        ]
    }


def main():
    data = create_data_model()
    kerf_width = data["kerf_width"]
    stock = data["stock"]
    demands = data["demands"]
    demand_lengths = [d["length"] for d in demands]

    # 生成所有原材料的切割模式
    stock_patterns = {}
    for s in stock:
        patterns = generate_patterns(s["length"], demand_lengths, kerf_width)
        stock_patterns[s["length"]] = {
            "patterns": patterns,
            "stock_qty": s["quantity"]
        }

    # 创建求解器
    solver = pywraplp.Solver.CreateSolver("SCIP")

    # 创建变量字典
    variables = defaultdict(list)
    for stock_len in stock_patterns:
        stock_qty = stock_patterns[stock_len]["stock_qty"]
        patterns = stock_patterns[stock_len]["patterns"]
        vars = [
            solver.IntVar(0, stock_qty, f"x_{stock_len}_{i}")
            for i in range(len(patterns))
        ]
        variables[stock_len] = vars

    # 库存约束
    for stock_len in stock_patterns:
        solver.Add(sum(variables[stock_len]) <= stock_patterns[stock_len]["stock_qty"])

    # 需求约束
    for i, demand in enumerate(demands):
        constraint = solver.Constraint(demand["quantity"], solver.infinity())
        for stock_len in variables:
            for var_idx, var in enumerate(variables[stock_len]):
                pattern = stock_patterns[stock_len]["patterns"][var_idx]["combo"]
                coefficient = pattern[i]
                constraint.SetCoefficient(var, coefficient)

    # 目标函数：最小化总使用次数
    objective = solver.Objective()
    for stock_len in variables:
        for var in variables[stock_len]:
            objective.SetCoefficient(var, 1)
    objective.SetMinimization()

    # 求解
    status = solver.Solve()

    if status == solver.OPTIMAL:
        print("优化成功，正在生成报告...")
        detailed_records = []
        plan_summary = []
        serial_no = 1

        # 解析结果
        for stock_len in variables:
            for var_idx, var in enumerate(variables[stock_len]):
                used = int(var.solution_value())
                if used == 0:
                    continue

                pattern = stock_patterns[stock_len]["patterns"][var_idx]
                # 汇总统计
                plan_summary.append({
                    "原材料长度(mm)": stock_len,
                    "使用次数": used,
                    "总余料(mm)": pattern["waste"] * used,
                    "总锯缝损耗(mm)": pattern["kerf"] * used,
                    "平均利用率(%)": pattern["utilization"],
                    "切割模式": " + ".join([f"{c}×{l}mm" for c, l in zip(pattern["combo"], demand_lengths)])
                })

                # 详细记录
                for _ in range(used):
                    detail = {
                        "序号": serial_no,
                        "原材料长度(mm)": stock_len,
                        "成品组合": " , ".join([f"{l}mm×{c}"
                                                for l, c in zip(demand_lengths, pattern["combo"])
                                                if c > 0]),
                        "余料(mm)": pattern["waste"],
                        "锯缝损耗(mm)": pattern["kerf"],
                        "总消耗(mm)": stock_len - pattern["waste"],
                        "材料利用率(%)": pattern["utilization"]  # 添加单个方案利用率
                    }
                    detailed_records.append(detail)
                    serial_no += 1

        # 生成Excel报告
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        output_path = os.path.join(desktop, "优化切割方案.xlsx")

        with pd.ExcelWriter(output_path) as writer:
            # 详细记录表
            df_detail = pd.DataFrame(detailed_records)
            df_detail = df_detail[[
                "序号", "原材料长度(mm)", "成品组合",
                "总消耗(mm)", "锯缝损耗(mm)", "余料(mm)", "材料利用率(%)"
            ]]
            df_detail.to_excel(writer, sheet_name="详细记录", index=False)

            # 方案汇总表
            df_summary = pd.DataFrame(plan_summary)
            # 计算整体利用率
            total_consumption = df_summary["原材料长度(mm)"] * df_summary["使用次数"] - df_summary["总余料(mm)"]
            total_stock = df_summary["原材料长度(mm)"] * df_summary["使用次数"]
            df_summary["整体利用率(%)"] = round((total_consumption / total_stock) * 100, 2)

            df_summary = df_summary[[
                "原材料长度(mm)", "使用次数", "切割模式",
                "总锯缝损耗(mm)", "总余料(mm)", "平均利用率(%)", "整体利用率(%)"
            ]]
            df_summary.to_excel(writer, sheet_name="方案汇总", index=False)

            # 需求完成情况
            completed = []
            for i, d in enumerate(demands):
                total = sum(
                    p["combo"][i] * var.solution_value()
                    for stock_len in variables
                    for p, var in zip(
                        stock_patterns[stock_len]["patterns"],
                        variables[stock_len]
                    )
                )
                completed.append({
                    "成品规格(mm)": d["length"],
                    "需求数量": d["quantity"],
                    "完成数量": int(total),
                    "完成率(%)": round(min(100, 100 * total / d["quantity"]), 2)
                })
            pd.DataFrame(completed).to_excel(writer, sheet_name="需求完成", index=False)

        print(f"报告已生成至：{output_path}")
    else:
        print("未找到可行解")


if __name__ == "__main__":
    main()
