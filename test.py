import collections
import json
import re
import sys
import platform
import subprocess
import telnetlib
import socket
import threading
import time
from struct import error

import dns.resolver
import pingparsing
from PyQt6 import sip
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QThread, QTimer
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QComboBox, QProgressBar,
    QGroupBox, QMessageBox, QFileDialog, QTableWidget, QHeaderView, QTableWidgetItem
)
from PyQt6.QtGui import QFont, QIcon, QPalette, QColor

# ==================== 工作线程类 ====================
class PingWorker(QObject):
    output = pyqtSignal(str, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str, str)

    def __init__(self, target, count):
        super().__init__()
        self.target = target
        self.count = count
        self._is_running = True
        self.process = None
        self.ping_parser = pingparsing.PingParsing()  # 创建 PingParsing 实例
        self.transmitter = pingparsing.PingTransmitter()  # 创建 PingTransmitter 实例
        self.transmitter.destination = self.target
        self.transmitter.count = self.count

    def run(self):
        try:
            # 执行ping操作并获取结果
            result = self.transmitter.ping()

            # 解析ping结果
            parsed_result = self.ping_parser.parse(result)

            # 获取解析后的数据
            stats = parsed_result.as_dict()

            # 向前端发送结果
            self.output.emit(self.target, json.dumps(stats))

            self.finished.emit(self.target)

        except Exception as e:
            self.error.emit(self.target, f"Ping Error: {str(e)}")
        finally:
            self.cleanup()

    def cleanup(self):
        """确保进程和资源被正确清理"""
        if self.process:
            try:
                self.process.terminate()
                self.process.wait(timeout=1)
            except:
                try:
                    self.process.kill()
                except:
                    pass
            finally:
                self.process = None

    def stop(self):
        """停止ping操作"""
        self._is_running = False
        self.cleanup()


class TelnetWorker(QObject):
    output = pyqtSignal(str)
    connected = pyqtSignal()
    disconnected = pyqtSignal()

    def __init__(self, host, port):
        super().__init__()
        self.host = host
        self.port = port
        self.tn = None
        self._is_connected = False
        self._is_running = False  # 新增运行状态标志

    def run(self):
        self._is_running = True
        try:
            self.tn = telnetlib.Telnet(self.host, self.port, timeout=3)
            self._is_connected = True
            success = {
                "status": True,
                "target": self.host,
                "port": self.port,
                "msg": "连接成功"
            }
            self.output.emit(json.dumps(success))
        except Exception as e:
            failed = {
                "status": False,
                "target": self.host,
                "port": self.port,
                "msg": "连接失败"
            }
            self.output.emit(json.dumps(failed))
        finally:
            self.disconnect()
            self.disconnected.emit()

    def stop(self):
        self._is_running = False
        self.cleanup()

    def cleanup(self):
        if self.tn:
            try:
                self.tn.close()
            except:
                pass
            self.tn = None

    def disconnect(self):
        if self._is_running:
            self._is_running = False  # 终止循环
        if self._is_connected and self.tn:
            try:
                self.tn.close()
            except Exception as e:
                pass
            finally:
                self.tn = None
            self._is_connected = False

# ==================== 主窗口类 ====================
class NetworkTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("高级网络诊断工具")
        self.setGeometry(200, 200, 900, 700)
        self.setMinimumSize(800, 600)
        self.setup_ui()
        self.setup_styles()

        # 工作线程和线程对象
        self.ping_works = {}
        self.ping_threads = {}

        self.telnet_works = {}
        self.telnet_threads = {}

        self.target_count = 0

        self.ping_finished_count = 0
        self.telnet_finished_count = 0

        self._lock = threading.Lock()

    def setup_ui(self):
        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        # 添加各功能标签页
        self.setup_ping_tab(main_layout)

        # 状态栏
        self.status_bar = self.statusBar()
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)

    def setup_ping_tab(self, parent_layout):
        # 设置区域
        settings_group = QGroupBox()
        settings_layout = QHBoxLayout(settings_group)

        settings_layout.addWidget(QLabel("目标地址:"))
        self.ping_host = QLineEdit()
        self.ping_host.setPlaceholderText("IP地址或域名，空格或逗号分隔")
        settings_layout.addWidget(self.ping_host)

        settings_layout.addWidget(QLabel("测试次数:"))
        self.ping_count = QLineEdit("4")
        self.ping_count.setFixedWidth(50)
        settings_layout.addWidget(self.ping_count)

        settings_layout.addWidget(QLabel("端口号:"))
        self.telnet_port = QLineEdit("23")
        self.telnet_port.setFixedWidth(50)
        settings_layout.addWidget(self.telnet_port)

        self.ping_btn = QPushButton("运行")
        self.ping_btn.setIcon(QIcon.fromTheme("media-playback-start"))
        self.ping_btn.clicked.connect(self.toggle_ping_and_telnet)
        settings_layout.addWidget(self.ping_btn)

        self.import_btn = QPushButton("IP导入")
        self.import_btn.setIcon(QIcon.fromTheme("document-open"))
        self.import_btn.clicked.connect(self.import_files)
        settings_layout.addWidget(self.import_btn)

        self.clean_btn = QPushButton("清空")
        self.clean_btn.setIcon(QIcon.fromTheme("document-open"))
        self.clean_btn.clicked.connect(self.clean_table)
        settings_layout.addWidget(self.clean_btn)
        parent_layout.addWidget(settings_group)

        # 输出区域
        output_group = QGroupBox()
        output_layout = QVBoxLayout(output_group)

        self.ping_table = QTableWidget()
        self.ping_table.setColumnCount(5)
        self.ping_table.setHorizontalHeaderLabels(["IP", "Ping延迟(ms)", "Ping丢包率(%)", "Telnet端口号", "Telnet端口状态"])
        self.ping_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        output_layout.addWidget(self.ping_table)

        parent_layout.addWidget(output_group, stretch=1)

    def setup_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }

            QTabWidget::pane {
                border: none;
            }

            QGroupBox {
                background-color: #ffffff;
                border: 1px solid #dcdcdc;
                border-radius: 6px;
                padding: 10px;
                margin-top: 20px;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 8px;
                font-size: 14px;
                font-weight: bold;
                color: #222;
            }

            QLabel {
                color: #222;
                font-size: 13px;
            }

            QLineEdit, QComboBox {
                border: 1px solid #c4c4c4;
                border-radius: 4px;
                padding: 6px 10px;
                font-size: 13px;
                background-color: #ffffff;
            }

            QTextEdit {
                background: #fcfcfc;
                border: 1px solid #d0d0d0;
                border-radius: 6px;
                padding: 8px;
                font-family: Consolas;
                font-size: 12.5px;
            }

            QPushButton {
                background-color: #2563eb;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: 500;
            }

            QPushButton:hover {
                background-color: #1d4ed8;
            }

            QPushButton:disabled {
                background-color: #bdbdbd;
            }

            QProgressBar {
                border: 1px solid #d0d0d0;
                background-color: #eaeaea;
                border-radius: 6px;
                text-align: center;
                font-size: 12px;
            }

            QProgressBar::chunk {
                background-color: #2563eb;
                border-radius: 6px;
            }

            QTabBar::tab {
                background: transparent;
                padding: 6px 12px;
                border: none;
                font-size: 13px;
                color: #555;
            }

            QTabBar::tab:selected {
                color: #2563eb;
                font-weight: bold;
                border-bottom: 2px solid #2563eb;
            }
            
            QTableWidget {
                border: 1px solid #DDDDDD;
                border-radius: 8px;
                background-color: #FFFFFF;
                gridline-color: #E0E0E0;
            }
            QTableWidget::item {
                border: 1px solid #F1F1F1;
                padding: 10px;
            }
            QTableWidget::item:selected {
                color: white;
            }
            QTableWidget::horizontalHeader {
                background-color: #1E90FF;
                color: white;
                font-weight: bold;
                font-size: 12pt;
                padding: 8px;
            }
            QTableWidget::horizontalHeader::section {
                padding-left: 10px;
                padding-right: 10px;
            }
            QTableWidget::verticalHeader {
                background-color: #F7F7F7;
                border-right: 1px solid #E0E0E0;
            }
            QTableWidget::item:selected {
                background-color: #87CEFA;
                color: white;
            }
        """)

    # ==================== 功能方法 ====================
    def toggle_ping_and_telnet(self):
        try:
            if self.ping_works:
                self.stop_all()
            else:
                self.start_ping()
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"[Ping错误] {e}")
        try:
            self.shutdown_all_telnet()
            self.connect_telnet()
        except Exception as e:
            print(f"[Telnet错误] {e}")

    def import_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择文件", "", "All Files (*)")
        if files:
            with open(files[0], "r") as f:
                ips = f.read().splitlines()
            self.ping_host.setText(",".join(ips))
            self.ping_count.setText(str(len(ips)))
            self.ping_btn.setText("运行")
            self.ping_host.setEnabled(True)
            self.ping_count.setEnabled(True)
            self.ping_table.setRowCount(0)

    def start_ping(self):
        if self.ping_btn.text() == "停止测试":
            self.stop_all()
            return
        targets = self.ping_host.text().strip().split(",")
        targets = set([t.strip() for t in targets if t.strip()])
        self.target_count = len(targets)
        if not targets:
            QMessageBox.warning(self, "错误", "请输入至少一个目标地址")
            return
        try:
            count = int(self.ping_count.text())
            if count <= 0:
                raise ValueError
        except ValueError:
            QMessageBox.warning(self, "错误", "请输入有效的测试次数")
            return

        self._init_table_data(targets)

        # 停止所有已有任务
        self.stop_all_ping()

        self.ping_btn.setText("停止测试")
        self.ping_host.setEnabled(False)
        self.ping_count.setEnabled(False)
        self.telnet_port.setEnabled(False)

        for target in targets:
            self.start_single_ping(target, count)

    def start_single_ping(self, target, count):
        worker = PingWorker(target, count)
        thread = QThread()
        worker.moveToThread(thread)

        # 信号连接
        thread.started.connect(worker.run)
        worker.output.connect(lambda t=target, text="": self.append_ping_output(t, text))
        worker.error.connect(lambda t=target, text="": self.append_ping_output(t, text))
        worker.finished.connect(thread.quit)
        worker.finished.connect(lambda: self.ping_finished(target))
        thread.finished.connect(lambda: self.cleanup_ping_resources(target))

        self.ping_works[target] = worker
        self.ping_threads[target] = thread
        thread.start()

    def _init_table_data(self, targets):
        self.ping_table.setRowCount(0)
        
        for target in targets:
            row = self.ping_table.rowCount()
            self.ping_table.insertRow(row)
            values = [target, "--", "--", "--", "--"]
            for col, val in enumerate(values):
                item = QTableWidgetItem(val)
                self.ping_table.setItem(row, col, item)

    def ping_update_table_row(self, target, time_ms, loss):
        for row in range(self.ping_table.rowCount()):
            item = self.ping_table.item(row, 0)  # 第一列是目标地址
            if item and item.text() == target:
                # 更新延迟(ms)
                time_item = QTableWidgetItem(time_ms)
                time_item.setFlags(time_item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # 禁用编辑
                self.ping_table.setItem(row, 1, time_item)

                # 更新丢包率(%)
                loss_item = QTableWidgetItem(loss)
                loss_item.setFlags(loss_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.ping_table.setItem(row, 2, loss_item)
                break

    def stop_all_ping(self):
        for target in self.ping_works.keys():
            self.stop_ping(target)
        self.ping_works.clear()
        self.ping_threads.clear()
        self.ping_btn.setText("运行")
        self.ping_host.setEnabled(True)
        self.ping_count.setEnabled(True)
        self.telnet_port.setEnabled(True)

    def stop_ping(self, target):
        worker = self.ping_works.pop(target, None)
        thread = self.ping_threads.pop(target, None)

        try:
            if worker and not sip.isdeleted(worker):
                worker.disconnect()
                worker.deleteLater()
        except Exception as e:
            print(f"[释放worker异常] {e}")

        try:
            if thread:
                if thread.isRunning():
                    thread.quit()
                    if not thread.wait(1000):
                        thread.terminate()
                        thread.wait()
                thread.deleteLater()
        except Exception as e:
            print(f"[释放thread异常] {e}")

    def append_ping_output(self, target, text):
        """
        将 ping 输出结果显示到输出框中，并加上目标前缀。
        """
        ping_result_json = json.loads(text)
        time_ms = '--' if ping_result_json.get("rtt_avg") is None else ping_result_json["rtt_avg"]
        loss = '--' if ping_result_json.get("packet_loss_rate") is None else ping_result_json["packet_loss_rate"]
        self.ping_update_table_row(target, str(time_ms), str(round(loss, 0)))

    def cleanup_ping_resources(self, target):
        worker = self.ping_works.pop(target, None)
        thread = self.ping_threads.pop(target, None)
        if worker:
            worker.deleteLater()
        if thread:
            thread.deleteLater()

    def clean_table(self):
        self.ping_table.setRowCount(0)

    def ping_finished(self, target):
        with self._lock:
            self.ping_finished_count+=1
            self.finished_all()

    def finished_all(self):
        print("total count: {}".format(self.target_count))
        print("telnet finished count: {}".format(self.telnet_finished_count))
        print("ping finished count: {}".format(self.ping_finished_count))

        if self.ping_finished_count == self.target_count and self.telnet_finished_count == self.target_count:
            self.ping_finished_count = 0
            self.telnet_finished_count = 0
            self.ping_btn.setText("运行")
            self.ping_host.setEnabled(True)
            self.ping_count.setEnabled(True)
            self.telnet_port.setEnabled(True)

    def shutdown_all_telnet(self):
        targets = list(self.telnet_works.keys()) if self.telnet_works else []
        for target in targets:
            try:
                self.disconnect_telnet(target)
            except Exception as e:
                print(f"[断开失败] {target}: {e}")

        self.telnet_works.clear()
        self.telnet_threads.clear()

    def connect_telnet(self):
        hosts = self.ping_host.text().strip().replace("，", ",").split(",")
        port_text = self.telnet_port.text().strip()

        try:
            port = int(port_text)
            if not (0 < port < 65536):
                raise ValueError
        except ValueError:
            self.telnet_output.append("错误：请输入有效的端口号 (1-65535)")
            return

        self.telnet_total = 0

        for host in hosts:
            if host.strip():
                self.telnet_total += 1
                self.start_single_telnet(host.strip(), port)

    def start_single_telnet(self, target, port):
        worker = TelnetWorker(target, port)
        thread = QThread()
        worker.moveToThread(thread)

        thread.started.connect(worker.run)
        worker.output.connect(self.handle_telnet_result)

        worker.disconnected.connect(lambda: self.on_telnet_finished(target))
        thread.finished.connect(thread.deleteLater)

        # 存入 map
        self.telnet_works[target] = worker
        self.telnet_threads[target] = thread
        thread.start()

    def on_telnet_finished(self, target):
        with self._lock:
            if getattr(self, "_is_stopping", False):
                return
            self.telnet_finished_count += 1
            self.finished_all()

            if self.telnet_finished_count >= self.telnet_total:
                self.on_all_telnet_done()

    def on_all_telnet_done(self):
        self.shutdown_all_telnet()

    def handle_telnet_result(self, result):
        result_json = json.loads(result)
        target = result_json.get("target")
        port = result_json.get("port")
        msg = result_json.get("msg")
        for row in range(self.ping_table.rowCount()):
            item = self.ping_table.item(row, 0)  # 第一列是目标地址
            if item and item.text() == target:
                port_item = QTableWidgetItem(str(port))
                port_item.setFlags(port_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.ping_table.setItem(row, 3, port_item)
                # 更新状态
                status_item = QTableWidgetItem(msg)
                status_item.setFlags(status_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.ping_table.setItem(row, 4, status_item)
                break

    def stop_all(self):
        self._is_stopping = True  # 设置标志位

        telnet_keys = list(self.telnet_works.keys()) if self.telnet_works else []
        for target in telnet_keys:
            self.disconnect_telnet(target)

        ping_keys = list(self.ping_works.keys()) if self.ping_works else []
        for target in ping_keys:
            self.stop_ping(target)

        self.telnet_total = 0
        self.target_count = 0
        self.ping_finished_count = 0
        self.telnet_finished_count = 0
        self.ping_works.clear()
        self.ping_threads.clear()
        self.telnet_works.clear()
        self.telnet_threads.clear()

        self.ping_btn.setText("运行")
        self.ping_host.setEnabled(True)
        self.ping_count.setEnabled(True)
        self.telnet_port.setEnabled(True)

        self._is_stopping = False

    def disconnect_telnet(self, target):
        worker = self.telnet_works.get(target)
        thread = self.telnet_threads.get(target)

        try:
            if worker and not sip.isdeleted(worker):
                worker.disconnect()
                worker.deleteLater()
        except Exception as e:
            print(f"[释放worker异常] {e}")

        try:
            if thread:
                if thread.isRunning():
                    thread.quit()
                    if not thread.wait(1000):
                        thread.terminate()
                        thread.wait()
                thread.deleteLater()
        except Exception as e:
            print(f"[释放thread异常] {e}")

    def closeEvent(self, event):
        # 先停止所有工作线程
        running_tasks = []

        if self.ping_works:
            targets = list(self.ping_works.keys()) if self.ping_works else []
            for target in targets:
                self.stop_ping(target)
            running_tasks.append("正在停止Ping测试...")
        if self.telnet_works:
            targets = list(self.telnet_works.keys()) if self.telnet_works else []
            for target in targets:
                self.disconnect_telnet(target)
            running_tasks.append("正在断开Telnet连接...")
        if running_tasks:
            # 显示等待对话框
            wait_dialog = QMessageBox(self)
            wait_dialog.setWindowTitle("请稍候")
            wait_dialog.setText("正在停止后台任务...\n" + "\n".join(running_tasks))
            wait_dialog.setStandardButtons(QMessageBox.StandardButton.NoButton)
            wait_dialog.show()

            # 非阻塞方式等待任务停止
            QTimer.singleShot(1000, lambda: (
                wait_dialog.close(),
                event.accept() if not self.ping_works else event.ignore()
            ))

            QTimer.singleShot(1000, lambda: (
                wait_dialog.close(),
                event.accept() if not self.telnet_works else event.ignore()
            ))
            event.ignore()
        else:
            event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = NetworkTool()
    window.show()
    sys.exit(app.exec())