import json
import sys
import threading
import socket
import time
import pingparsing
from PyQt6 import sip
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QThread, QTimer
from PyQt6.QtGui import QIcon, QFont, QColor, QValidator
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QProgressBar,
    QGroupBox, QMessageBox, QFileDialog, QTableWidget,
    QHeaderView, QTableWidgetItem, QTabWidget, QSpacerItem, QSizePolicy
)


class NetworkTestWorker(QObject):
    """整合Ping和端口测试的工作线程"""
    test_result = pyqtSignal(str, str)  # 目标地址, JSON结果
    finished = pyqtSignal(str)  # 目标地址
    error = pyqtSignal(str, str)  # 目标地址, 错误信息

    def __init__(self, target, ping_count=4, port=None, port_timeout=3):
        super().__init__()
        self.target = target
        self.ping_count = ping_count
        self.port = port
        self.port_timeout = port_timeout
        self._is_running = True
        self._socket = None  # 用于存储socket对象

        # Ping相关初始化
        if ping_count > 0:
            self.ping_parser = pingparsing.PingParsing()
            self.ping_transmitter = pingparsing.PingTransmitter()
            self.ping_transmitter.destination = self.target
            self.ping_transmitter.timeout = 1
            self.ping_transmitter.count = self.ping_count

    def run(self):
        """执行测试"""
        if not self._is_running:
            return

        try:
            result = {
                "target": self.target,
                "ping": None,
                "port": None
            }

            # 执行Ping测试
            if self.ping_count > 0:
                ping_result = self._run_ping_test()
                if ping_result:
                    result["ping"] = ping_result

            # 执行端口测试
            if self.port:
                port_result = self._run_port_test()
                if port_result:
                    result["port"] = port_result

            if self._is_running:
                self.test_result.emit(self.target, json.dumps(result))

        except Exception as e:
            if self._is_running:
                self.error.emit(self.target, f"测试错误: {str(e)}")
        finally:
            if self._is_running:
                self.finished.emit(self.target)
            self._is_running = False

    def _run_ping_test(self):
        """执行Ping测试"""
        try:
            result = self.ping_transmitter.ping()
            if not self._is_running:
                return None

            parsed_result = self.ping_parser.parse(result)
            stats = parsed_result.as_dict()

            print(f"原始Ping结果: {stats}")  # 调试输出

            return {
                "latency": stats["rtt_avg"],
                "loss": stats["packet_loss_rate"],
                "raw": stats
            }
        except Exception as e:
            print(f"Ping测试错误: {str(e)}")  # 调试输出
            return {
                "latency": "--",
                "loss": "--",
                "error": str(e)
            }

    def _run_port_test(self):
        """执行端口测试"""
        if not self.port:
            return None

        try:
            # 创建socket连接
            self._socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self._socket.settimeout(self.port_timeout)

            # 尝试连接
            start_time = time.time()
            self._socket.connect((self.target, self.port))
            end_time = time.time()

            return {
                "status": True,
                "port": self.port,
                "response_time": (end_time - start_time) * 1000,
                "msg": f"连接成功 (响应时间: {(end_time - start_time) * 1000:.2f}ms)"
            }

        except socket.timeout:
            return {
                "status": False,
                "port": self.port,
                "msg": "连接超时"
            }
        except ConnectionRefusedError:
            return {
                "status": False,
                "port": self.port,
                "msg": "连接被拒绝"
            }
        except Exception as e:
            return {
                "status": False,
                "port": self.port,
                "msg": f"连接失败: {str(e)}"
            }
        finally:
            if self._socket:
                try:
                    self._socket.close()
                except:
                    pass
                self._socket = None

    def stop(self):
        """停止测试操作"""
        self._is_running = False

        # 停止ping操作
        if hasattr(self, 'ping_transmitter') and self.ping_transmitter:
            try:
                if hasattr(self.ping_transmitter, 'stop'):
                    self.ping_transmitter.stop()
            except:
                pass

        # 关闭socket连接
        if self._socket:
            try:
                self._socket.close()
            except:
                pass
            self._socket = None


class NetworkTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("高级网络诊断工具")
        self.setGeometry(200, 200, 1000, 800)
        self.setMinimumSize(900, 700)

        # 初始化UI
        self.init_ui()
        self.init_styles()

        # 工作线程管理
        self.test_workers = {}
        self.test_threads = {}

        # 状态跟踪
        self.target_count = 0
        self.test_finished_count = 0
        self._lock = threading.Lock()
        self._is_stopping = False

        self.test_workers = {}
        self.test_threads = {}
        self._stop_event = threading.Event()  # 用于协调停止操作
        self._stop_lock = threading.Lock()  # 停止操作专用锁

    def init_ui(self):
        """初始化用户界面"""
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # 主布局
        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # 创建标签页
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Ping测试标签页
        self.setup_ping_tab()

        # 状态栏
        self.setup_status_bar()

    def setup_ping_tab(self):
        """设置Ping测试标签页"""
        ping_tab = QWidget()
        self.tab_widget.addTab(ping_tab, "Ping & 端口测试")

        # Ping测试布局
        ping_layout = QVBoxLayout(ping_tab)
        ping_layout.setContentsMargins(10, 10, 10, 10)
        ping_layout.setSpacing(15)

        # 控制面板
        self.setup_control_panel(ping_layout)

        # 结果表格
        self.setup_results_table(ping_layout)

    def setup_control_panel(self, parent_layout):
        """设置控制面板"""
        control_group = QGroupBox("测试设置")
        control_layout = QVBoxLayout(control_group)
        control_layout.setContentsMargins(15, 15, 15, 15)
        control_layout.setSpacing(15)

        # 第一行：输入设置
        input_layout = QHBoxLayout()

        input_layout.addWidget(QLabel("目标地址:"))
        self.ping_host = QLineEdit()
        self.ping_host.setPlaceholderText("IP地址或域名，多个用逗号或空格分隔")
        self.ping_host.setMinimumWidth(300)
        input_layout.addWidget(self.ping_host)

        input_layout.addWidget(QLabel("Ping次数:"))
        self.ping_count = QLineEdit("4")
        self.ping_count.setFixedWidth(50)
        self.ping_count.setValidator(QIntValidator(1, 100))
        input_layout.addWidget(self.ping_count)

        input_layout.addWidget(QLabel("端口号:"))
        self.port_test_port = QLineEdit("23")
        self.port_test_port.setFixedWidth(50)
        self.port_test_port.setValidator(QIntValidator(1, 65535))
        input_layout.addWidget(self.port_test_port)

        control_layout.addLayout(input_layout)

        # 第二行：按钮布局
        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)

        # 添加弹性空间使按钮居中
        button_layout.addItem(QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))

        self.run_btn = QPushButton("开始测试")
        self.run_btn.setIcon(QIcon.fromTheme("media-playback-start"))
        self.run_btn.clicked.connect(self.start_testing)
        button_layout.addWidget(self.run_btn)

        self.export_btn = QPushButton("导出结果")
        self.export_btn.setIcon(QIcon.fromTheme("document-save"))
        self.export_btn.clicked.connect(self.export_results)
        button_layout.addWidget(self.export_btn)

        self.import_btn = QPushButton("导入IP")
        self.import_btn.setIcon(QIcon.fromTheme("document-open"))
        self.import_btn.clicked.connect(self.import_targets)
        button_layout.addWidget(self.import_btn)

        self.clear_btn = QPushButton("清空结果")
        self.clear_btn.setIcon(QIcon.fromTheme("edit-clear"))
        self.clear_btn.clicked.connect(self.clear_results)
        button_layout.addWidget(self.clear_btn)

        # 添加弹性空间使按钮居中
        button_layout.addItem(QSpacerItem(20, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum))

        control_layout.addLayout(button_layout)
        parent_layout.addWidget(control_group)

    def setup_results_table(self, parent_layout):
        """设置结果表格"""
        results_group = QGroupBox("测试结果")
        results_layout = QVBoxLayout(results_group)
        results_layout.setContentsMargins(10, 15, 10, 10)

        # 创建表格
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(5)
        self.results_table.setHorizontalHeaderLabels([
            "IP地址", "Ping延迟(ms)", "Ping丢包率(%)", "端口号", "端口状态"
        ])

        # 表格样式设置
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.results_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.results_table.setAlternatingRowColors(True)

        results_layout.addWidget(self.results_table)
        parent_layout.addWidget(results_group, stretch=1)

    def setup_status_bar(self):
        """设置状态栏"""
        self.status_bar = self.statusBar()

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)

        # 状态标签
        self.status_label = QLabel("就绪")
        self.status_bar.addWidget(self.status_label)

    def init_styles(self):
        """初始化样式表"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f7fa;
            }
            QGroupBox {
                background-color: white;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                margin-top: 20px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
                color: #374151;
                font-weight: bold;
            }
            QLabel {
                color: #374151;
            }
            QLineEdit, QComboBox {
                border: 1px solid #d1d5db;
                border-radius: 4px;
                padding: 5px 10px;
                background-color: white;
            }
            QLineEdit:disabled, QComboBox:disabled {
                background-color: #f3f4f6;
            }
            QPushButton {
                background-color: #3b82f6;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QPushButton:disabled {
                background-color: #9ca3af;
            }
            QPushButton:pressed {
                background-color: #1d4ed8;
            }
            QTableWidget {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: white;
                gridline-color: #e5e7eb;
            }
            QHeaderView::section {
                background-color: #3b82f6;
                color: white;
                padding: 6px;
                border: none;
            }
            QProgressBar {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                text-align: center;
                background-color: #f3f4f6;
            }
            QProgressBar::chunk {
                background-color: #3b82f6;
                border-radius: 6px;
            }
            QTabWidget::pane {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background: white;
                margin-top: 5px;
            }
            QTabBar::tab {
                background: transparent;
                padding: 8px 16px;
                border: none;
                color: #6b7280;
            }
            QTabBar::tab:selected {
                color: #3b82f6;
                font-weight: bold;
                border-bottom: 2px solid #3b82f6;
            }
        """)

        # 设置字体
        font = QFont()
        font.setFamily("Segoe UI" if sys.platform == "win32" else "Arial")
        font.setPointSize(10)
        self.setFont(font)

    def start_testing(self):
        """开始测试"""
        print("开始测试")  # 调试输出
        # 获取并验证输入
        targets = self._parse_targets()
        if not targets:
            self.show_warning("请输入至少一个目标地址")
            return

        try:
            ping_count = int(self.ping_count.text())
            if ping_count <= 0:
                raise ValueError
        except ValueError:
            self.show_warning("请输入有效的Ping测试次数")
            return

        try:
            port = int(self.port_test_port.text())
            if not (0 < port < 65536):
                raise ValueError
        except ValueError:
            self.show_warning("请输入有效的端口号 (1-65535)")
            return

        # 重置计数器
        self.target_count = len(targets)
        self.test_finished_count = 0
        print(f"目标数量: {self.target_count}")  # 调试输出

        # 初始化UI状态
        self._init_results_table(targets)
        self._set_ui_testing_state(True)

        # 开始测试
        for target in targets:
            self._start_test(target, ping_count, port)

    def _parse_targets(self):
        """解析目标地址输入"""
        input_text = self.ping_host.text().strip()
        if not input_text:
            return []

        # 支持逗号、空格、换行分隔
        separators = [',', ' ', '\n', ';']
        for sep in separators:
            input_text = input_text.replace(sep, ' ')

        # 分割并去重
        targets = list(set(t.strip() for t in input_text.split() if t.strip()))
        return targets

    def _init_results_table(self, targets):
        """初始化结果表格"""
        self.results_table.setRowCount(len(targets))

        for row, target in enumerate(targets):
            # IP地址列
            ip_item = QTableWidgetItem(target)
            self.results_table.setItem(row, 0, ip_item)

            # Ping延迟列
            ping_latency_item = QTableWidgetItem("测试中...")
            ping_latency_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.results_table.setItem(row, 1, ping_latency_item)

            # Ping丢包率列
            ping_loss_item = QTableWidgetItem("测试中...")
            ping_loss_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.results_table.setItem(row, 2, ping_loss_item)

            # 端口号列
            port_item = QTableWidgetItem(self.port_test_port.text())
            port_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.results_table.setItem(row, 3, port_item)

            # 端口状态列
            port_status_item = QTableWidgetItem("测试中...")
            port_status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.results_table.setItem(row, 4, port_status_item)

    def _set_ui_testing_state(self, testing):
        """设置UI测试状态"""
        self.run_btn.setEnabled(not testing)
        self.ping_host.setEnabled(not testing)
        self.ping_count.setEnabled(not testing)
        self.port_test_port.setEnabled(not testing)
        self.import_btn.setEnabled(not testing)
        self.clear_btn.setEnabled(not testing)
        self.export_btn.setEnabled(not testing)
        self.progress_bar.setVisible(testing)

        if testing:
            self.progress_bar.setRange(0, self.target_count)
            self.progress_bar.setValue(0)
            self.status_label.setText("测试进行中...")
        else:
            self.status_label.setText("就绪")

    def _start_test(self, target, ping_count, port):
        """启动单个测试"""
        print(f"启动测试: {target}")  # 调试输出
        if target in self.test_workers:
            self._stop_test(target)

        # 创建工作线程
        worker = NetworkTestWorker(target, ping_count, port)
        thread = QThread()
        worker.moveToThread(thread)

        # 确保信号连接正确
        def handle_result(t, r):
            print(f"收到结果: {t}")  # 调试输出
            print(r)
            self._handle_test_result(t, r)

        def handle_error(t, e):
            print(f"收到错误: {t}, {e}")  # 调试输出
            self._handle_test_error(t, e)

        def handle_finished(t):
            print(f"测试完成信号: {t}")  # 调试输出
            self._on_test_finished(t)

        # 连接信号
        thread.started.connect(worker.run)
        worker.test_result.connect(handle_result)
        worker.error.connect(handle_error)
        worker.finished.connect(handle_finished)

        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(worker.deleteLater)

        # 存储引用
        self.test_workers[target] = worker
        self.test_threads[target] = thread

        # 启动线程
        thread.start()

    def _stop_test(self, target):
        """停止单个测试"""
        worker = self.test_workers.pop(target, None)
        thread = self.test_threads.pop(target, None)

        if worker:
            worker.stop()
            if not sip.isdeleted(worker):
                worker.deleteLater()

        if thread and thread.isRunning():
            thread.quit()
            if not thread.wait(500):
                thread.terminate()
                thread.wait()
            if not sip.isdeleted(thread):
                thread.deleteLater()

    def _find_target_row(self, target):
        """查找目标地址在表格中的行号"""
        try:
            for row in range(self.results_table.rowCount()):
                item = self.results_table.item(row, 0)
                if item and item.text() == target:
                    return row
            return None
        except Exception as e:
            print(f"查找行错误: {str(e)}")  # 调试输出
            return None

    def _handle_test_error(self, target, error_msg):
        """处理测试错误"""
        row = self._find_target_row(target)
        if row is None:
            return

        # 标记Ping结果为错误
        for col in [1, 2]:
            item = QTableWidgetItem("错误")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setForeground(QColor("#ef4444"))
            self.results_table.setItem(row, col, item)

        # 标记端口测试结果为错误
        item = QTableWidgetItem("错误")
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        item.setForeground(QColor("#ef4444"))
        item.setToolTip(error_msg)
        self.results_table.setItem(row, 4, item)

        self.status_label.setText(f"{target}: {error_msg}")

    def _handle_test_result(self, target, result_json):
        """处理测试结果 - 确保在主线程执行"""
        try:
            if sip.isdeleted(self.results_table):  # 检查表格是否已被删除
                return

            result = json.loads(result_json)
            row = self._find_target_row(target)
            if row is None:
                return

            # 处理Ping结果
            if "ping" in result:
                ping_result = result["ping"]
                latency = ping_result.get("latency", "--")
                loss = ping_result.get("loss", "--")

                # 确保在主线程更新UI
                def update_ping():
                    if sip.isdeleted(self.results_table):
                        return
                    # 更新延迟
                    latency_text = str(latency) if latency != "--" else "--"
                    latency_item = QTableWidgetItem(latency_text)
                    latency_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.results_table.setItem(row, 1, latency_item)

                    # 更新丢包率
                    loss_percent = f"{float(loss):.1f}%" if loss != "--" else "--"
                    loss_item = QTableWidgetItem(loss_percent)
                    loss_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.results_table.setItem(row, 2, loss_item)

                    # 设置颜色
                    color = QColor()
                    if latency == "--" or loss == "--":
                        color.setNamedColor("#ef4444")
                    elif float(loss) > 0:
                        color.setNamedColor("#f59e0b")
                    else:
                        color.setNamedColor("#10b981")

                    for col in [1, 2]:
                        item = self.results_table.item(row, col)
                        if item:
                            item.setForeground(color)

                QTimer.singleShot(0, update_ping)

            # 处理端口测试结果
            if "port" in result:
                port_result = result["port"]
                status = port_result.get("status", False)
                msg = port_result.get("msg", "未知状态")

                def update_port():
                    if sip.isdeleted(self.results_table):
                        return
                    status_item = QTableWidgetItem("成功" if status else "失败")
                    status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    status_item.setToolTip(msg)
                    status_item.setForeground(QColor("#10b981" if status else "#ef4444"))
                    self.results_table.setItem(row, 4, status_item)

                QTimer.singleShot(0, update_port)

        except Exception as e:
            print(f"处理结果错误: {str(e)}")

    def _on_test_finished(self, target):
        """单个测试完成处理"""
        print(f"测试完成: {target}")
        try:
            with self._lock:
                self.test_finished_count += 1
                print(f"进度: {self.test_finished_count}/{self.target_count}")

                # 使用invokeMethod确保线程安全
                self.progress_bar.setValue(self.test_finished_count)

                # 清理资源
                if target in self.test_workers:
                    worker = self.test_workers.pop(target)
                    if worker and not sip.isdeleted(worker):
                        worker.deleteLater()

                if target in self.test_threads:
                    thread = self.test_threads.pop(target)
                    if thread and thread.isRunning():
                        thread.quit()
                        if not thread.wait(500):
                            thread.terminate()
                        if not sip.isdeleted(thread):
                            thread.deleteLater()

                # 检查是否全部完成
                if self.test_finished_count >= self.target_count:
                    print("所有测试已完成，准备更新UI")
                    # 确保在主线程执行UI更新
                    QTimer.singleShot(0, self._finalize_testing)
        except Exception as e:
            print(f"完成处理错误: {str(e)}")

    def _finalize_testing(self):
        """最终完成测试处理"""
        try:
            print(f"最终完成测试处理 called, 当前按钮文本: {self.run_btn.text()}")

            # 确保在主线程执行UI更新
            if not sip.isdeleted(self.run_btn):
                self.run_btn.setEnabled(True)

            if not sip.isdeleted(self.ping_host):
                self.ping_host.setEnabled(True)
                self.ping_count.setEnabled(True)
                self.port_test_port.setEnabled(True)
                self.import_btn.setEnabled(True)
                self.clear_btn.setEnabled(True)
                self.export_btn.setEnabled(True)

            if not sip.isdeleted(self.progress_bar):
                self.progress_bar.setVisible(False)

            if not sip.isdeleted(self.status_label):
                self.status_label.setText("测试完成")

            print("UI状态已更新")
            # 显示完成通知
            QMessageBox.information(self, "完成", "所有测试已完成!")
        except Exception as e:
            print(f"最终处理错误: {str(e)}")

    def _on_all_tests_finished(self):
        """所有测试完成处理"""
        # 确保在主线程执行UI更新
        QTimer.singleShot(0, self._update_ui_after_tests_complete)

    def _update_ui_after_tests_complete(self):
        """测试完成后更新UI"""
        if not sip.isdeleted(self.run_btn):  # 确保按钮未被删除
            self.run_btn.setEnabled(True)

        self.ping_host.setEnabled(True)
        self.ping_count.setEnabled(True)
        self.port_test_port.setEnabled(True)
        self.import_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.export_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText("测试完成")

        # 显示完成通知
        QMessageBox.information(self, "完成", "所有测试已完成!")

    def export_results(self):
        """导出测试结果"""
        if self.results_table.rowCount() == 0:
            self.show_warning("没有可导出的结果")
            return

        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("导出测试结果")
        file_dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
        file_dialog.setNameFilter("JSON文件 (*.json);;CSV文件 (*.csv);;所有文件 (*)")

        if file_dialog.exec():
            file_paths = file_dialog.selectedFiles()
            if file_paths:
                file_path = file_paths[0]
                try:
                    if file_path.endswith('.json'):
                        self._export_to_json(file_path)
                    elif file_path.endswith('.csv'):
                        self._export_to_csv(file_path)
                    else:
                        # 默认导出为JSON
                        self._export_to_json(file_path + '.json')

                    self.status_label.setText(f"结果已导出到: {file_path}")
                except Exception as e:
                    self.show_warning(f"导出失败: {str(e)}")

    def _export_to_json(self, file_path):
        """将结果导出为JSON格式"""
        results = []
        for row in range(self.results_table.rowCount()):
            result = {
                "target": self.results_table.item(row, 0).text(),
                "ping_latency": self.results_table.item(row, 1).text(),
                "ping_loss": self.results_table.item(row, 2).text(),
                "port": self.results_table.item(row, 3).text(),
                "port_status": self.results_table.item(row, 4).text()
            }
            results.append(result)

        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=4, ensure_ascii=False)

    def _export_to_csv(self, file_path):
        """将结果导出为CSV格式"""
        import csv
        with open(file_path, 'w', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            # 写入表头
            writer.writerow(["IP地址", "Ping延迟(ms)", "Ping丢包率(%)", "端口号", "端口状态"])

            # 写入数据
            for row in range(self.results_table.rowCount()):
                row_data = [
                    self.results_table.item(row, 0).text(),
                    self.results_table.item(row, 1).text(),
                    self.results_table.item(row, 2).text(),
                    self.results_table.item(row, 3).text(),
                    self.results_table.item(row, 4).text()
                ]
                writer.writerow(row_data)

    def import_targets(self):
        """从文件导入目标地址"""
        file_dialog = QFileDialog(self)
        file_dialog.setWindowTitle("选择目标地址文件")
        file_dialog.setNameFilter("文本文件 (*.txt);;所有文件 (*)")

        if file_dialog.exec():
            file_paths = file_dialog.selectedFiles()
            if file_paths:
                try:
                    with open(file_paths[0], 'r', encoding='utf-8') as f:
                        targets = [line.strip() for line in f if line.strip()]
                        self.ping_host.setText(",".join(targets))
                except Exception as e:
                    self.show_warning(f"无法读取文件: {str(e)}")

    def clear_results(self):
        """清空结果表格"""
        self.results_table.setRowCount(0)
        self.status_label.setText("结果已清空")

    def show_warning(self, message):
        """显示警告消息"""
        QMessageBox.warning(self, "警告", message)

    def closeEvent(self, event):
        """窗口关闭事件处理"""
        if len(self.test_workers) > 0:
            reply = QMessageBox.question(
                self, "确认退出",
                "测试正在进行中，确定要退出吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.No:
                event.ignore()
                return

        # 停止所有测试
        for target in list(self.test_workers.keys()):
            self._stop_test(target)
        event.accept()


class QIntValidator(QValidator):
    """简单的整数验证器"""

    def __init__(self, min_val, max_val):
        super().__init__()
        self.min_val = min_val
        self.max_val = max_val

    def validate(self, input_str, pos):
        if not input_str:
            return (QValidator.State.Intermediate, input_str, pos)

        try:
            value = int(input_str)
            if self.min_val <= value <= self.max_val:
                return (QValidator.State.Acceptable, input_str, pos)
            return (QValidator.State.Invalid, input_str, pos)
        except ValueError:
            return (QValidator.State.Invalid, input_str, pos)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # 设置应用程序信息
    app.setApplicationName("高级网络诊断工具")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("NetworkTools")

    window = NetworkTool()
    window.show()
    sys.exit(app.exec())
