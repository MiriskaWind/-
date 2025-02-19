import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QGroupBox, QLineEdit, QTableWidget,
                             QHeaderView, QMessageBox, QFileDialog, QLabel, QItemDelegate, QTableWidgetItem)
from PyQt5.QtGui import QFont, QIntValidator, QIcon
from PyQt5.QtCore import Qt
import pandas as pd
import itertools
from ortools.linear_solver import pywraplp
import os
from collections import defaultdict


# 定义全局变量
stock_data = []
demands_data = []


class IntegerDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)

    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        validator = QIntValidator()
        editor.setValidator(validator)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        editor.setText(str(value))

    def setModelData(self, editor, model, index):
        value = editor.text()
        model.setData(index, value, Qt.EditRole)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LinerCut")  # 设置窗口标题
        self.setGeometry(100, 100, 800, 600)  # 设置窗口大小
        try:
            self.setWindowIcon(QIcon('./icon/LinerCut.png'))  # 设置窗口图标
        except Exception as e:
            print(f"Error setting icon: {e}")
        self.initUI()

    def initUI(self):
        # 设置全局字体
        font = QFont("Microsoft YaHei")
        font.setBold(False)  # 取消粗体
        self.setFont(font)

        # 布局
        main_layout = QVBoxLayout()

        # 第一排按钮
        button_layout = QHBoxLayout()
        self.new_button = QPushButton("新建")
        self.open_button = QPushButton("打开")
        self.calculate_button = QPushButton("计算")
        self.save_button = QPushButton("保存")
        self.template_button = QPushButton("模板生成")  # 新增模板生成按钮

        button_layout.addWidget(self.new_button)
        button_layout.addWidget(self.open_button)
        button_layout.addWidget(self.calculate_button)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.template_button)  # 添加模板生成按钮到布局

        main_layout.addLayout(button_layout)

        # 第二排：参数设置
        parameter_group = QGroupBox("参数设置")
        parameter_layout = QHBoxLayout()
        self.saw_kerf_label = QLabel("刀口锯缝 (mm):")
        self.saw_kerf_input = QLineEdit("5")
        self.saw_kerf_input.setValidator(QIntValidator())  # 只允许整数
        parameter_layout.addWidget(self.saw_kerf_label)
        parameter_layout.addWidget(self.saw_kerf_input)
        parameter_group.setLayout(parameter_layout)
        main_layout.addWidget(parameter_group)

        # 第三排：表格
        table_layout = QHBoxLayout()

        # Stock表格
        self.stock_group = QGroupBox("Stock")
        self.stock_table = QTableWidget(30, 2)  # 30行2列
        self.stock_table.setHorizontalHeaderLabels(["Length", "Quantity"])
        self.stock_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.stock_table.setItemDelegate(IntegerDelegate(self))  # 设置整数代理
        stock_layout = QVBoxLayout()
        stock_layout.addWidget(self.stock_table)
        self.stock_group.setLayout(stock_layout)
        table_layout.addWidget(self.stock_group)

        # Demands表格
        self.demands_group = QGroupBox("Demands")
        self.demands_table = QTableWidget(30, 2)  # 30行2列
        self.demands_table.setHorizontalHeaderLabels(["Length", "Quantity"])
        self.demands_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.demands_table.setItemDelegate(IntegerDelegate(self))  # 设置整数代理
        demands_layout = QVBoxLayout()
        demands_layout.addWidget(self.demands_table)
        self.demands_group.setLayout(demands_layout)
        table_layout.addWidget(self.demands_group)

        main_layout.addLayout(table_layout)

        # 设置主布局
        self.setLayout(main_layout)

        # 连接信号和槽
        self.new_button.clicked.connect(self.new_data)
        self.open_button.clicked.connect(self.open_excel)
        self.calculate_button.clicked.connect(self.run_optimization)  # 连接计算按钮
        self.save_button.clicked.connect(self.save_data)
        self.template_button.clicked.connect(self.generate_template)  # 连接模板生成按钮

    def new_data(self):
        # 清空stock和demands表格的数据
        self.stock_table.clearContents()
        self.demands_table.clearContents()

        # 清空全局变量
        global stock_data, demands_data
        stock_data = []
        demands_data = []

    def open_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "打开Excel文件", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                # 读取Excel文件
                excel_data = pd.read_excel(file_path, sheet_name=None)  # 读取所有sheet

                # 检查是否存在名为 "Stock" 和 "Demands" 的 sheet
                if "Stock" in excel_data and "Demands" in excel_data:
                    stock_data = excel_data["Stock"]
                    demands_data = excel_data["Demands"]

                    # 将数据填充到对应的表格中
                    self.fill_table_with_data(self.stock_table, stock_data)
                    self.fill_table_with_data(self.demands_table, demands_data)

                    # 更新全局变量
                    self.update_global_data()

                else:
                    QMessageBox.warning(self, "警告", "Excel文件中缺少名为 'Stock' 或 'Demands' 的sheet。")

            except FileNotFoundError:
                QMessageBox.critical(self, "错误", f"文件未找到：{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"打开Excel文件时出错：{str(e)}")

    def fill_table_with_data(self, table, data):
        # 清空表格
        table.clearContents()
        table.setRowCount(0)

        # 设置表格行数
        table.setRowCount(len(data.index))

        # 填充数据
        for row in range(len(data.index)):  # Use data.index for row iteration
            for col in range(len(data.columns)):
                item = str(data.iloc[row, col])  # 将数据转换为字符串
                table.setItem(row, col, QTableWidgetItem(item))

    def save_data(self):
        # 使用全局变量
        global stock_data, demands_data
        self.update_global_data()

        # 输出到控制台
        print({"stock": stock_data, "demands": demands_data})

    def update_global_data(self):
        # 更新全局变量
        global stock_data, demands_data

        stock_data = []
        for row in range(self.stock_table.rowCount()):
            try:
                length_item = self.stock_table.item(row, 0)
                quantity_item = self.stock_table.item(row, 1)

                # 确保 length 和 quantity 都不是 None 并且有文本内容
                if length_item is not None and length_item.text() != "" and quantity_item is not None and quantity_item.text() != "":
                    length_val = int(length_item.text())
                    quantity_val = int(quantity_item.text())
                    stock_data.append({"length": length_val, "quantity": quantity_val})
            except ValueError:
                print(f"Invalid data in stock table row {row}. Skipping.")
            except Exception as e:
                print(f"Error processing stock table row {row}: {e}")

        demands_data = []
        for row in range(self.demands_table.rowCount()):
            try:
                length_item = self.demands_table.item(row, 0)
                quantity_item = self.demands_table.item(row, 1)

                # 确保 length 和 quantity 都不 None 并且有文本内容
                if length_item is not None and length_item.text() != "" and quantity_item is not None and quantity_item.text() != "":
                    length_val = int(length_item.text())
                    quantity_val = int(quantity_item.text())
                    demands_data.append({"length": length_val, "quantity": quantity_val})
            except ValueError:
                print(f"Invalid data in demand table row {row}. Skipping.")
            except Exception as e:
                print(f"Error processing demand table row {row}: {e}")

    def run_optimization(self):
        # 更新全局变量
        self.update_global_data()

        # 获取刀口锯缝的值
        try:
            kerf_width = int(self.saw_kerf_input.text())
        except ValueError:
            QMessageBox.warning(self, "警告", "无效的刀口锯缝值，请使用整数。")
            return

        # 运行优化算法
        try:
            main(kerf_width)
        except Exception as e:
            print(f"Error during optimization: {e}")
            QMessageBox.critical(self, "错误", f"优化失败: {e}")

    def generate_template(self):
        """在桌面生成下料模板Excel文件"""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        output_path = os.path.join(desktop, "下料模板.xlsx")

        # 创建空的DataFrame，只包含列名
        stock_df = pd.DataFrame(columns=["Length", "Quantity"])
        demands_df = pd.DataFrame(columns=["Length", "Quantity"])

        try:
            with pd.ExcelWriter(output_path) as writer:
                stock_df.to_excel(writer, sheet_name="Stock", index=False)
                demands_df.to_excel(writer, sheet_name="Demands", index=False)

            QMessageBox.information(self, "模板生成", f"下料模板已生成至：{output_path}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成模板时出错：{str(e)}")


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
        total_used = sum(int(c) * int(l) for c, l in zip(combo, demand_lengths))
        total_kerf = int(kerf_width) * (total_pieces - 1) if total_pieces > 1 else 0
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


def create_data_model(kerf_width):
    """包含锯缝参数的数据模型"""
    global stock_data, demands_data
    return {
        "kerf_width": kerf_width,
        "stock": stock_data,
        "demands": demands_data
    }


def main(kerf_width):
    try:
        data = create_data_model(kerf_width)
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
                for i, demand in enumerate(demands):
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
    except Exception as e:
        print(f"Error in main function: {e}")  # 打印错误信息


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())