import sys
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QGroupBox, QLineEdit, QTableWidget,
                             QHeaderView, QMessageBox, QFileDialog, QLabel, QItemDelegate, QTableWidgetItem,
                             QProgressDialog, QDesktopWidget, QSystemTrayIcon, QMenu, QAction)
from PyQt5.QtGui import QFont, QIntValidator, QIcon, QPalette, QColor, QPixmap, QClipboard
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMutex, QWaitCondition
import pandas as pd
import itertools
from ortools.linear_solver import pywraplp
from collections import defaultdict
import datetime
import math
import io


# 定义全局变量
stock_data = []
demands_data = []


class IntegerDelegate(QItemDelegate):
    """
    A delegate class to ensure only integers can be entered into the table cells.
    """
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
    error_signal = pyqtSignal(str)

    def __init__(self, kerf_width, solver_time_limit):
        super().__init__()
        self.kerf_width = kerf_width
        self.solver_time_limit = solver_time_limit
        self.error_message = None  # Store error message if optimization fails
        self.mutex = QMutex()
        self.wait_condition = QWaitCondition()
        self.cancelled = False

    def run(self):
        try:
            output_path = main(self.kerf_width, self.solver_time_limit, self.progress_update, self.mutex, self.wait_condition, self)
            if not self.cancelled:
                self.result_ready.emit(output_path)  # Emit the path to the Excel file
        except Exception as e:
            self.error_message = str(e)  # Store the error message
            self.error_signal.emit(self.error_message)  # Emit the error message
            print(f"OptimizationThread.run error: {e}")

    def cancel(self):
        self.mutex.lock()
        self.cancelled = True
        self.wait_condition.wakeAll()  # Wake up the thread if it's waiting
        self.mutex.unlock()

class CustomProgressDialog(QProgressDialog):  # 继承自QProgressDialog
    def __init__(self, parent=None): # 继承自QProgressDialog
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
                background-color: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0,
                                    stop: 0 #32CD32, stop: 1 #008000); /* LimeGreen to Green */
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

    def cancel(self):
        """
        Handles the cancellation of the optimization process.
        """
        print("取消优化")
        self.setValue(0)
        self.close()


class MainWindow(QWidget):
    def __init__(self): # 继承自QWidget
        super().__init__()
        self.setWindowTitle("LinerCut")  # 设置窗口标题
        self.setGeometry(100, 100, 400, 600)  # 设置窗口大小
        self.center()  # Center the window on the screen

        # 设置窗口属性
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_NoSystemBackground, False)

        try:
            # 使用绝对路径
            icon_path = "C:/Users/tobei/AppData/Roaming/JetBrains/PyCharm2024.3/scratches/LinerCut/icon/LinerCut.png"

            pixmap = QPixmap(icon_path)
            pixmap = pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)  # 调整图标大小
            self.app_icon = QIcon(pixmap)
            self.setWindowIcon(self.app_icon)  # 设置窗口图标
            self.tray_icon = None  # Initialize tray_icon
            self.setup_tray_icon()  # Set up the system tray icon
        except Exception as e:
            print(f"Error setting icon: {e}")
            self.app_icon = None
        self.initUI()
        self.set_table_style()  # 应用表格样式

    def set_table_style(self):
        """设置表格样式"""
        palette = self.stock_table.palette()
        palette.setColor(QPalette.Highlight, QColor("#b7ffb7"))  # 设置选中颜色
        palette.setColor(QPalette.HighlightedText, QColor("black"))  # 设置选中文字颜色
        self.stock_table.setPalette(palette)
        self.demands_table.setPalette(palette)

    def initUI(self):
        # 设置全局字体，整个UI使用统一字体
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

        # 统一设置按钮样式
        for button in [self.new_button, self.open_button, self.calculate_button, self.save_button, self.template_button]:
            button_layout.addWidget(button)

        main_layout.addLayout(button_layout)

        # 第二排：参数设置
        parameter_group = QGroupBox("参数设置")
        parameter_layout = QHBoxLayout()

        # 创建一个 QHBoxLayout 用于锯缝标签和输入框，实现水平布局
        saw_kerf_hbox = QHBoxLayout()
        self.saw_kerf_label = QLabel("锯缝 (mm):")
        self.saw_kerf_input = QLineEdit("5")
        self.saw_kerf_input.setValidator(QIntValidator())  # 只允许整数
        self.saw_kerf_input.setMaximumWidth(50)  # 进一步减少锯缝输入栏宽度

        # 将标签和输入框添加到水平布局中
        saw_kerf_hbox.addWidget(self.saw_kerf_label)
        saw_kerf_hbox.addWidget(self.saw_kerf_input)

        # 创建一个 QHBoxLayout 用于求解时间标签和输入框，实现水平布局
        solver_time_hbox = QHBoxLayout()
        self.solver_time_label = QLabel("求解时间 (秒):")
        self.solver_time_input = QLineEdit("60")
        self.solver_time_input.setValidator(QIntValidator())  # 只允许整数
        self.solver_time_input.setMaximumWidth(50)  # 进一步减少求解时间输入栏宽度

        # 将标签和输入框添加到水平布局中
        solver_time_hbox.addWidget(self.solver_time_label)
        solver_time_hbox.addWidget(self.solver_time_input)

        # 将水平布局添加到参数布局中
        parameter_layout.addLayout(saw_kerf_hbox)
        parameter_layout.addLayout(solver_time_hbox)

        # 添加伸缩器，使标签和输入框靠左对齐
        parameter_layout.addStretch(1)

        parameter_group.setLayout(parameter_layout)
        main_layout.addWidget(parameter_group)

        # 第三排：表格
        table_layout = QHBoxLayout()

        # Stock表格
        self.stock_group = QGroupBox("Stock")
        stock_table_layout = QVBoxLayout()

        # Stock表格的按钮，增加和删除行
        stock_button_layout = QHBoxLayout()
        self.add_stock_row_button = QPushButton("增加行")
        self.delete_stock_row_button = QPushButton("删除行")
        stock_button_layout.addWidget(self.add_stock_row_button)
        stock_button_layout.addWidget(self.delete_stock_row_button)
        stock_table_layout.addLayout(stock_button_layout)

        self.stock_table = QTableWidget(15, 2)  # 15行2列
        self.stock_table.setHorizontalHeaderLabels(["Length", "Quantity"]) # 设置表头
        self.stock_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.stock_table.setItemDelegate(IntegerDelegate(self))  # 设置整数代理
        # 设置行高，使得表格更美观
        self.stock_table.verticalHeader().setDefaultSectionSize(18)
        stock_table_layout.addWidget(self.stock_table)
        self.stock_group.setLayout(stock_table_layout)
        table_layout.addWidget(self.stock_group)

        # Demands表格
        self.demands_group = QGroupBox("Demands") # 需求表
        demands_table_layout = QVBoxLayout()

        # Demands表格的按钮
        demands_button_layout = QHBoxLayout()
        self.add_demands_row_button = QPushButton("增加行")
        self.delete_demands_row_button = QPushButton("删除行")
        demands_button_layout.addWidget(self.add_demands_row_button)
        demands_button_layout.addWidget(self.delete_demands_row_button)
        demands_table_layout.addLayout(demands_button_layout)

        self.demands_table = QTableWidget(15, 2)  # 15行2列
        self.demands_table.setHorizontalHeaderLabels(["Length", "Quantity"]) # 设置表头
        self.demands_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.demands_table.setItemDelegate(IntegerDelegate(self))  # 设置整数代理
        # 设置行高
        self.demands_table.verticalHeader().setDefaultSectionSize(18)
        demands_table_layout.addWidget(self.demands_table)
        self.demands_group.setLayout(demands_table_layout)
        table_layout.addWidget(self.demands_group)

        main_layout.addLayout(table_layout)

        # 设置主布局
        self.setLayout(main_layout)

        # 连接信号和槽，实现按钮点击功能
        self.new_button.clicked.connect(self.new_data)
        self.open_button.clicked.connect(self.open_excel)
        self.calculate_button.clicked.connect(self.run_optimization)  # 连接计算按钮
        self.save_button.clicked.connect(self.save_data)
        self.template_button.clicked.connect(self.generate_template)  # 连接模板生成按钮

        # 连接表格按钮信号和槽
        self.add_stock_row_button.clicked.connect(self.add_stock_row)
        self.delete_stock_row_button.clicked.connect(self.delete_stock_row)
        self.add_demands_row_button.clicked.connect(self.add_demands_row)
        self.delete_demands_row_button.clicked.connect(self.delete_demands_row)

        # 添加粘贴快捷键，方便用户输入数据
        self.stock_table.keyPressEvent = self.stock_table_keyPressEvent
        self.demands_table.keyPressEvent = self.demands_table_keyPressEvent

    def stock_table_keyPressEvent(self, event):
        if event.key() == Qt.Key_V and (event.modifiers() & Qt.ControlModifier):
            self.paste_data(self.stock_table)
        else:
            QTableWidget.keyPressEvent(self.stock_table, event)

    def demands_table_keyPressEvent(self, event):
        if event.key() == Qt.Key_V and (event.modifiers() & Qt.ControlModifier):
            self.paste_data(self.demands_table)
        else:
            QTableWidget.keyPressEvent(self.demands_table, event)

    def paste_data(self, table):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        if not text:
            return

        # 将剪贴板数据转换为 DataFrame，方便处理
        try:
            df = pd.read_csv(io.StringIO(text), sep='\t', header=None)
        except Exception as e:
            QMessageBox.warning(self, "警告", f"粘贴数据解析失败：{str(e)}")
            return

        # 获取当前选中的单元格
        selected_range = table.selectedRanges()
        if not selected_range:
            start_row = 0
            start_col = 0
        else:
            start_row = selected_range[0].topRow()
            start_col = selected_range[0].leftColumn()

        # 粘贴数据
        for i in range(df.shape[0]):
            row = start_row + i
            if row >= table.rowCount():
                table.insertRow(row)
            for j in range(min(df.shape[1], 2)):  # 只取前两列
                col = start_col + j
                if col >= table.columnCount():
                    break
                item = str(df.iloc[i, j])
                table.setItem(row, col, QTableWidgetItem(item))

    def new_data(self): # 新建数据
        # 清空stock和demands表格的数据
        self.stock_table.clearContents()
        self.demands_table.clearContents()

        # 重置表格行列数
        self.stock_table.setRowCount(15)
        self.demands_table.setRowCount(15)

        # 清空全局变量
        global stock_data, demands_data
        stock_data = []
        demands_data = []

    def open_excel(self): # 打开excel文件
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

    def fill_table_with_data(self, self_table, data): # 填充表格数据
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

    def save_data(self): # 保存数据
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

    def run_optimization(self): # 运行优化
        # 更新全局变量
        self.update_global_data()

        # 获取刀口锯缝的值
        try:
            kerf_width = int(self.saw_kerf_input.text())
        except ValueError:
            QMessageBox.warning(self, "警告", "无效的锯缝值，请使用整数。")
            return

        # 获取求解时间的值
        try:
            solver_time_limit = int(self.solver_time_input.text()) * 1000  # 转换为毫秒
        except ValueError:
            QMessageBox.warning(self, "警告", "无效的求解时间值，请使用整数。")
            return

        # 创建并显示进度对话框
        self.progress_dialog = CustomProgressDialog(self)
        self.progress_dialog.canceled.connect(self.cancel_optimization)  # Connect cancel signal

        self.progress_dialog.show()

        # 创建并启动优化线程
        self.optimization_thread = OptimizationThread(kerf_width, solver_time_limit)
        try:
            self.optimization_thread.progress_update.connect(self.update_progress)
            self.optimization_thread.result_ready.connect(self.optimization_finished)
            self.optimization_thread.error_signal.connect(self.optimization_failed)
        except TypeError as e:
            print(f"Error connecting signals: {e}")  # Debugging
            QMessageBox.critical(self, "错误", f"信号连接失败: {e}")
            return # Exit if signal connection fails
        try:
            self.optimization_thread.start()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"启动优化线程失败: {e}")
            print(f"启动优化线程失败: {e}")

    def update_progress(self, value): # 更新进度条
        """
        Updates the progress bar in the progress dialog.
        """
        self.progress_dialog.setValue(value)

    def optimization_finished(self, output_path):
        """
        Handles the result of the optimization thread.
        Displays a message box with the result or any error message.
        """
        if output_path:
            QMessageBox.information(self, "优化完成", f"优化切割方案已生成至：{output_path}")
        else:
            QMessageBox.critical(self, "优化失败", f"优化失败，请查看控制台输出。")

    def optimization_failed(self, error_message):
        QMessageBox.critical(self, "优化失败", f"优化失败: {error_message}")

    def cancel_optimization(self): # 取消优化
        """
        Cancels the optimization process.
        """
        print("取消优化")
        self.optimization_thread.cancel()
        self.progress_dialog.cancel()

    def center(self): # 窗口居中显示
        """Centers the window on the screen."""
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def setup_tray_icon(self):
        """设置系统托盘图标"""
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
        """处理系统托盘图标的点击事件"""
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
        """在桌面生成下料模板Excel文件，方便用户填写数据"""
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

    def add_stock_row(self):
        """增加stock表格的行"""
        self.stock_table.insertRow(self.stock_table.rowCount())

    def delete_stock_row(self): # 删除stock表格的行
        """Deletes the selected rows from the stock table."""
        selected_ranges = self.stock_table.selectedRanges()
        rows_to_delete = set()
        for selected_range in selected_ranges:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                rows_to_delete.add(row)

        rows_to_delete = sorted(list(rows_to_delete), reverse=True)
        for row in rows_to_delete:
            self.stock_table.removeRow(row)

    def delete_demands_row(self): # 删除demands表格的行
        """Deletes the selected rows from the demands table."""
        selected_ranges = self.demands_table.selectedRanges()
        rows_to_delete = set()
        for selected_range in selected_ranges:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                rows_to_delete.add(row)

        rows_to_delete = sorted(list(rows_to_delete), reverse=True)
        for row in rows_to_delete:
            self.demands_table.removeRow(row)

    def add_demands_row(self): # 增加demands表格的行
        """Adds a row to the demands table."""
        self.demands_table.insertRow(self.demands_table.rowCount())

def generate_patterns(stock_length, demand_lengths, kerf_width, progress_callback, total_demands):
    """生成考虑锯缝的有效切割模式"""
    patterns = []
    max_counts = [stock_length // length for length in demand_lengths]

    # 生成所有可能的切割组合
    for i, combo in enumerate(itertools.product(*[range(0, c + 1) for c in max_counts])):
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
        # 计算进度
        progress = int((i + 1) / total_demands * 100)
        progress_callback.emit(progress)
    return patterns

def create_data_model(kerf_width):
    """包含锯缝参数的数据模型"""
    global stock_data, demands_data
    return {
        "kerf_width": kerf_width,
        "stock": stock_data,
        "demands": demands_data
    }


def main(kerf_width, solver_time_limit, progress_callback, mutex, wait_condition, thread):
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
        total_stock_count = len(stock)

        # 计算总的需求数量，用于计算进度
        total_demands = 1
        for length in demand_lengths:
            total_demands *= (stock[0]["length"] // length + 1)

        # 生成所有原材料的切割模式
        stock_patterns = {}
        for i, s in enumerate(stock):
            # Check for cancellation
            mutex.lock()
            if thread.cancelled:
                mutex.unlock()
                return None
            mutex.unlock()

            # 将 progress_callback 传递给 generate_patterns
            patterns = generate_patterns(s["length"], demand_lengths, kerf_width, progress_callback, total_demands)
            stock_patterns[s["length"]] = {
                "patterns": patterns,
                "stock_qty": s["quantity"]
            }
            #progress = int(math.pow((i + 1) / total_stock_count, 0.5) * 20)  # Apply a non-linear mapping (sqrt)
            #progress_callback.emit(progress)

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
        #progress_callback.emit(20)  # Indicate solver is running

        # 在求解器运行过程中更新进度条
        #solver_time_limit = 60000  # 设置求解器时间限制为60秒 (milliseconds)
        solver.SetTimeLimit(solver_time_limit)
        status = solver.Solve()

        if status == solver.OPTIMAL:
            print("优化成功，正在生成报告...")
            detailed_records = []
            plan_summary = []
            serial_no = 1

            # 初始化统计数据
            total_stock_length_used = 0
            total_finished_count = 0
            total_finished_length = 0
            max_waste = 0

            # 解析结果
            for stock_len in variables:
                for var_idx, var in enumerate(variables[stock_len]):
                    # Check for cancellation
                    mutex.lock()
                    if thread.cancelled:
                        mutex.unlock()
                        return None
                    mutex.unlock()

                    used = int(var.solution_value())
                    if used == 0:
                        continue

                    pattern = stock_patterns[stock_len]["patterns"][var_idx]
                    # 汇总统计
                    cutting_pattern = " + ".join([f"{l}mm{c}" for c, l in zip(pattern["combo"], demand_lengths) if c > 0])
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

                        # 更新统计数据
                        total_stock_count += 1
                        total_stock_length_used += stock_len
                        total_finished_count += sum(pattern["combo"])
                        total_finished_length += sum(c * l for c, l in zip(pattern["combo"], demand_lengths))
                        max_waste = max(max_waste, pattern["waste"])

            progress_callback.emit(70)

            # 计算总材料利用率
            total_initial_stock_length = sum(s["length"] * s["quantity"] for s in stock)
            total_utilization = round((total_finished_length / total_stock_length_used) * 100, 2) if total_stock_length_used else 0

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

                # 新增统计信息表
                summary_data = {
                    "项目": [
                        "原材料总数",
                        "原材料总使用长度(mm)",
                        "切割出来的成品总数量",
                        "切割出来的成品总长度(mm)",
                        "总材料利用率(%)",
                        "最长余料长度(mm)"
                    ],
                    "数值": [
                        total_stock_count,
                        total_stock_length_used,
                        total_finished_count,
                        total_finished_length,
                        total_utilization,
                        max_waste
                    ]
                }
                df_summary_info = pd.DataFrame(summary_data)
                df_summary_info.to_excel(writer, sheet_name="统计信息", index=False)

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
    # 设置适当的标志以禁用Windows上的控制台窗口
    if os.name == 'nt':
        import ctypes
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)
    app = QApplication(sys.argv)
    # 修改为False，确保所有窗口关闭后程序退出
    app.setQuitOnLastWindowClosed(False)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
