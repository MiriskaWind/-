import sys
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QGroupBox, QLineEdit, QTableWidget,
                             QHeaderView, QMessageBox, QFileDialog, QLabel, QItemDelegate, QTableWidgetItem,
                             QProgressDialog, QDesktopWidget, QSystemTrayIcon, QMenu, QAction)
from PyQt5.QtGui import QFont, QIntValidator, QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import itertools
from ortools.linear_solver import pywraplp
from collections import defaultdict
import datetime


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


class OptimizationThread(QThread):
    """
    A QThread class to run the optimization in a separate thread,
    allowing the GUI to remain responsive and display progress.
    """
    progress_update = pyqtSignal(int)  # Signal to update progress bar
    result_ready = pyqtSignal(str)  # Signal to send the result (path to the Excel file) or error message

    def __init__(self, kerf_width):
        super().__init__()
        self.kerf_width = kerf_width
        self.error_message = None  # Store error message if optimization fails

    def run(self):
        try:
            output_path = main(self.kerf_width, self.progress_update)
            self.result_ready.emit(output_path)  # Emit the path to the Excel file
        except Exception as e:
            self.error_message = str(e)  # Store the error message
            self.result_ready.emit(None)  # Emit None to signal an error


class CustomProgressDialog(QProgressDialog):  # 继承自QProgressDialog
    def __init__(self, parent=None):
        super().__init__("优化计算中...", "取消", 0, 100, parent)
        self.setWindowModality(Qt.WindowModal)
        self.setWindowTitle("优化进度")
        self.setCancelButtonText("取消")
        self.setAutoClose(True)  # Close automatically when finished
        self.setAutoReset(True)  # Reset when finished
        self.setMinimumDuration(0)  # Show immediately
        self.setValue(0)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint & ~Qt.WindowContextHelpButtonHint)  # 禁止最大化按钮
        self.setFixedSize(self.width(), self.height()) # 禁止拉伸窗口大小

        # 移除进度条的文本显示
        self.setLabelText("")  # 移除标签文本

        # 设置进度条样式 - Modern Flat Green
        self.setStyleSheet("""
            QProgressDialog {
                background-color: #f5f5f5; /* Very light grey background */
                color: #444444; /* Dark grey text */
                border: none; /* No border */
            }
            QProgressBar {
                border: none;
                border-radius: 8px; /* Rounded corners */
                text-align: center;
                background-color: #e0e0e0; /* Light grey background for the bar */
                color: #444444;
                height: 20px; /* Adjust height as needed */
            }
            QProgressBar::chunk {
                background-color: #2ecc71; /* Emerald green */
                border-radius: 8px;
                /* No width specified, it will fill */
            }
            QPushButton {
                background-color: #ffffff;
                border: 1px solid #bdc3c7; /* Light grey border */
                border-radius: 5px;
                padding: 5px 15px; /* More padding for better button appearance */
                color: #444444;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #ecf0f1; /* Light hover effect */
            }
            QPushButton:pressed {
                background-color: #d4e6f1; /* Even lighter pressed effect */
            }
        """)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LinerCut")  # 设置窗口标题
        self.setGeometry(100, 100, 400, 600)  # 设置窗口大小
        self.center()  # Center the window on the screen
        try:
            # Determine if running as a bundled application
            if getattr(sys, 'frozen', False):
                # If bundled, use the path of the executable
                base_path = os.path.dirname(sys.executable)
                icon_path = os.path.join(base_path, 'icon', 'LinerCut.png')
            else:
                # If not bundled, use the script's directory
                base_path = os.path.dirname(os.path.abspath(__file__))
                icon_path = os.path.join(base_path, 'icon', 'LinerCut.png')

            self.app_icon = QIcon(icon_path)
            self.setWindowIcon(self.app_icon)  # 设置窗口图标
            self.tray_icon = None  # Initialize tray_icon
            self.setup_tray_icon()  # Set up the system tray icon
        except Exception as e:
            print(f"Error setting icon: {e}")
            self.app_icon = None
        self.initUI()

    def initUI(self):
        # 设置全局字体
        font = QFont("Microsoft YaHei", 9) # 9 is a reasonable default size
        font.setBold(False)  # 取消粗体
        self.setFont(font)

        # 布局
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(5, 5, 5, 5)  # Add a little padding around the main layout
        main_layout.setSpacing(8)  # Adjust spacing between widgets

        # 第一排按钮
        button_layout = QHBoxLayout()
        button_layout.setSpacing(6) # Adjust spacing between buttons
        self.new_button = QPushButton("新建")
        self.open_button = QPushButton("打开")
        self.calculate_button = QPushButton("计算")
        self.save_button = QPushButton("保存")
        self.template_button = QPushButton("模板生成")  # 新增模板生成按钮

        # Apply consistent styling to buttons
        for button in [self.new_button, self.open_button, self.calculate_button, self.save_button, self.template_button]:
            button_layout.addWidget(button)

        main_layout.addLayout(button_layout)

        # 第二排：参数设置
        parameter_group = QGroupBox("参数设置")
        parameter_layout = QHBoxLayout()

        # 创建一个 QHBoxLayout 用于锯缝标签和输入框
        saw_kerf_hbox = QHBoxLayout()
        self.saw_kerf_label = QLabel("锯缝 (mm):")
        self.saw_kerf_input = QLineEdit("5")
        self.saw_kerf_input.setValidator(QIntValidator())  # 只允许整数
        self.saw_kerf_input.setMaximumWidth(50)  # 进一步减少锯缝输入栏宽度

        # 将标签和输入框添加到水平布局中
        saw_kerf_hbox.addWidget(self.saw_kerf_label)
        saw_kerf_hbox.addWidget(self.saw_kerf_input)

        # 将水平布局添加到参数布局中
        parameter_layout.addLayout(saw_kerf_hbox)

        # 添加伸缩器，使标签和输入框靠左对齐
        parameter_layout.addStretch(1)

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
        # 设置行高
        self.stock_table.verticalHeader().setDefaultSectionSize(18)
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
        # 设置行高
        self.demands_table.verticalHeader().setDefaultSectionSize(18)
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

    def fill_table_with_data(self, self_table, data):
        # 清空表格
        self_table.clearContents()
        self_table.setRowCount(0)

        # 设置表格行数
        self_table.setRowCount(len(data.index))

        # 填充数据
        for row in range(len(data.index)):  # Use data.index for row iteration
            for col in range(len(data.columns)):
                item = str(data.iloc[row, col])  # 将数据转换为字符串
                self_table.setItem(row, col, QTableWidgetItem(item))

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

                # 确保 length 和 quantity 都不是 None 并且有文本内容
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
            QMessageBox.warning(self, "警告", "无效的锯缝值，请使用整数。")
            return

        # Create and show the progress dialog
        self.progress_dialog = CustomProgressDialog(self)

        self.progress_dialog.show()

        # 创建并启动优化线程
        self.optimization_thread = OptimizationThread(kerf_width)
        try:
            self.optimization_thread.progress_update.connect(self.update_progress)
            self.optimization_thread.result_ready.connect(self.optimization_finished)
        except TypeError as e:
            print(f"Error connecting signals: {e}")  # Debugging
            QMessageBox.critical(self, "错误", f"信号连接失败: {e}")
            return # Exit if signal connection fails
        self.optimization_thread.start()

    def update_progress(self, value):
        """
        Updates the progress bar in the progress dialog.
        """
        self.progress_dialog.setValue(value)
        if self.progress_dialog.wasCanceled():
            self.optimization_thread.terminate()  # Stop the optimization thread if canceled
            self.optimization_thread.wait()

    def optimization_finished(self, output_path):
        """
        Handles the result of the optimization thread.
        Displays a message box with the result or any error message.
        """
        if output_path:
            QMessageBox.information(self, "优化完成", f"优化切割方案已生成至：{output_path}")
        else:
            QMessageBox.critical(self, "优化失败", f"优化失败: {self.optimization_thread.error_message}")

    def center(self):
        """Centers the window on the screen."""
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def setup_tray_icon(self):
        """Sets up the system tray icon."""
        if self.app_icon is None:
            return

        self.tray_icon = QSystemTrayIcon(self.app_icon, self)
        self.tray_icon.setToolTip("LinerCut")

        # Create a context menu
        tray_menu = QMenu()
        show_action = QAction("显示", self)
        hide_action = QAction("隐藏", self)
        exit_action = QAction("退出", self)

        show_action.triggered.connect(self.show)
        hide_action.triggered.connect(self.hide)
        exit_action.triggered.connect(QApplication.instance().quit)

        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addSeparator()
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        # Connect the double-click event to show the window
        self.tray_icon.activated.connect(self.tray_icon_activated)

    def tray_icon_activated(self, reason):
        """Handles click events on the system tray icon."""
        if reason == QSystemTrayIcon.DoubleClick:
            self.show()
            self.activateWindow()  # Bring to front

    def closeEvent(self, event):
        """Overrides the close event to minimize to tray instead of closing."""
        # 修改为完全退出程序
        self.tray_icon.hide()  # 确保托盘图标被移除
        QApplication.instance().quit()
        event.accept()

    def generate_template(self):
        """在桌面生成下料模板Excel文件"""
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        # 获取当前时间并格式化
        current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_path = os.path.join(desktop, f"下料模板_{current_time}.xlsx")

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


def main(kerf_width, progress_callback):
    """
    Main function to run the optimization.
    Includes a callback to update the progress bar.
    """
    try:
        data = create_data_model(kerf_width)
        kerf_width = data["kerf_width"]
        stock = data["stock"]
        demands = data["demands"]
        demand_lengths = [d["length"] for d in demands]

        # 生成所有原材料的切割模式
        stock_patterns = {}
        for i, s in enumerate(stock):
            patterns = generate_patterns(s["length"], demand_lengths, kerf_width)
            stock_patterns[s["length"]] = {
                "patterns": patterns,
                "stock_qty": s["quantity"]
            }
            progress = int((i + 1) / len(stock) * 20)  # 20% for pattern generation
            progress_callback.emit(progress)

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
            # Change the constraint to an equality constraint
            constraint = solver.Constraint(demand["quantity"], demand["quantity"])
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
        progress_callback.emit(50)  # Indicate solver is running

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
                    cutting_pattern = " + ".join([f"{l}mm×{c}" for c, l in zip(pattern["combo"], demand_lengths) if c > 0])
                    plan_summary.append({
                        "原材料长度(mm)": stock_len,
                        "使用次数": used,
                        "总余料(mm)": pattern["waste"] * used,
                        "总锯缝损耗(mm)": pattern["kerf"] * used,
                        "平均利用率(%)": pattern["utilization"],
                        "切割模式": cutting_pattern  # 剔除数量为0的成品尺寸
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
            progress_callback.emit(70)

            # 生成Excel报告
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            # 获取当前时间并格式化
            current_time = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            output_path = os.path.join(desktop, f"优化切割方案_{current_time}.xlsx")

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
                    # Ensure that the completed quantity does not exceed the demand quantity
                    completed_quantity = min(int(total), demands[i]["quantity"])

                    completed.append({
                        "成品规格(mm)": demands[i]["length"],
                        "需求数量": demands[i]["quantity"],
                        "完成数量": completed_quantity,
                        "完成率(%)": round(min(100, 100 * completed_quantity / demands[i]["quantity"]), 2)
                    })
                pd.DataFrame(completed).to_excel(writer, sheet_name="需求完成", index=False)

            print(f"报告已生成至：{output_path}")
            progress_callback.emit(100) # Indicate completion of Excel writing

            return output_path
        else:
            print("未找到可行解")
            return None
    except Exception as e:
        print(f"Error in main function: {e}")  # 打印错误信息
        raise e


if __name__ == "__main__":
    # Set the appropriate flag for disabling the console window on Windows
    if os.name == 'nt':
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    app = QApplication(sys.argv)
    # 修改为False，确保所有窗口关闭后程序退出
    app.setQuitOnLastWindowClosed(False)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
