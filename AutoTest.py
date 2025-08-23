import sys
import json
import sqlite3
import threading
import time
from datetime import datetime

import pyautogui
from pynput import mouse, keyboard
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTableWidget, QTableWidgetItem, QInputDialog, QMessageBox,
    QLabel, QHeaderView, QAbstractItemView
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont


class AutoActionRecorder(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoAction Recorder - 自动化操作录制与回放")
        self.resize(800, 600)

        # 数据库初始化
        self.db_conn = sqlite3.connect("actions.db", check_same_thread=False)
        self.create_table()

        # 成员变量
        self.is_recording = False
        self.is_playing = False
        self.recorded_events = []
        self.start_time = 0
        self.mouse_listener = None
        self.keyboard_listener = None
        self.play_thread = None
        self.stop_play_flag = threading.Event()
        self.last_move_event = None  # 用于合并 mouse_move

        # 按键映射
        self.key_map = {
            keyboard.Key.esc: 'esc', keyboard.Key.f1: 'f1', keyboard.Key.f2: 'f2',
            keyboard.Key.f3: 'f3', keyboard.Key.f4: 'f4', keyboard.Key.f5: 'f5',
            keyboard.Key.f6: 'f6', keyboard.Key.f7: 'f7', keyboard.Key.f8: 'f8',
            keyboard.Key.f9: 'f9', keyboard.Key.f10: 'f10', keyboard.Key.f11: 'f11',
            keyboard.Key.f12: 'f12', keyboard.Key.tab: 'tab', keyboard.Key.space: 'space',
            keyboard.Key.enter: 'enter', keyboard.Key.backspace: 'backspace',
            keyboard.Key.delete: 'delete', keyboard.Key.up: 'up', keyboard.Key.down: 'down',
            keyboard.Key.left: 'left', keyboard.Key.right: 'right',
            keyboard.Key.shift: 'shift', keyboard.Key.ctrl: 'ctrl',
            keyboard.Key.alt: 'alt', keyboard.Key.cmd: 'cmd',
        }

        self.init_ui()
        self.load_records()

    def create_table(self):
        cursor = self.db_conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS record (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                events_json TEXT NOT NULL,
                event_count INTEGER NOT NULL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                description TEXT
            )
        ''')
        self.db_conn.commit()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 标题
        title_label = QLabel("自动化操作录制与回放工具")
        title_label.setFont(QFont("SimHei", 16, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # 按钮区域（仅保留记录、播放、删除）
        btn_layout = QHBoxLayout()
        self.btn_record = QPushButton("记录")
        self.btn_play = QPushButton("播放")
        self.btn_delete = QPushButton("删除")

        self.btn_record.clicked.connect(self.toggle_record)
        self.btn_play.clicked.connect(self.play_selected_or_current)
        self.btn_delete.clicked.connect(self.delete_record)

        for btn in [self.btn_record, self.btn_play, self.btn_delete]:
            btn.setMinimumHeight(40)
            btn_layout.addWidget(btn)
        layout.addLayout(btn_layout)

        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["ID", "名称", "事件数", "创建时间", "备注"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        layout.addWidget(self.table)

        # 状态栏
        self.statusBar().showMessage("就绪")

    def toggle_record(self):
        if not self.is_recording:
            self.start_recording()
        else:
            self.stop_recording()

    def start_recording(self):
        time.sleep(1)
        self.recorded_events = []
        self.start_time = time.time()
        self.is_recording = True
        self.btn_record.setText("停止")
        self.btn_record.setStyleSheet("background-color: #d32f2f; color: white;")
        self.statusBar().showMessage("正在录制... 按下 '停止' 结束")
        self.mouse_listener = mouse.Listener(on_move=self.on_mouse_move, on_click=self.on_mouse_click, on_scroll=self.on_mouse_scroll)
        self.keyboard_listener = keyboard.Listener(on_press=self.on_key_press, on_release=self.on_key_release)
        self.mouse_listener.start()
        self.keyboard_listener.start()

    def stop_recording(self):
        self.is_recording = False
        self.btn_record.setText("记录")
        self.btn_record.setStyleSheet("")
        self.statusBar().showMessage(f"录制结束，共 {len(self.recorded_events)} 个事件")

        if self.mouse_listener:
            self.mouse_listener.stop()
        if self.keyboard_listener:
            self.keyboard_listener.stop()

        # 新增：删除最后一个事件（如果是点击【停止】按钮）
        last_event = self.recorded_events[-1]
        if (last_event["type"] == "mouse_click" and
            not last_event["pressed"]):
            self.recorded_events.pop()
            self.statusBar().showMessage(f"已移除【停止】按钮点击事件，当前事件数：{len(self.recorded_events)}")
            
        # 保存最后一个 mouse_move 段
        if self.last_move_event:
            self.recorded_events.append(self.last_move_event)
            self.last_move_event = None

        # 录制结束后自动保存
        if self.recorded_events:
            self.save_current_record()
        else:
            self.statusBar().showMessage("录制内容为空，未保存")

    def on_mouse_move(self, x, y):
        if self.is_recording:
            current_time = time.time() - self.start_time

            # 尝试合并到上一个 move
            if (self.last_move_event and
                abs(self.last_move_event["time"] + self.last_move_event["duration"] - current_time) < 0.1):
                self.last_move_event["x_to"] = x
                self.last_move_event["y_to"] = y
                self.last_move_event["duration"] = current_time - self.last_move_event["time"]
                return

            # 否则保存上一个，开启新的
            if self.last_move_event:
                self.recorded_events.append(self.last_move_event)

            self.last_move_event = {
                "type": "mouse_move",
                "x_from": x, "y_from": y,
                "x_to": x, "y_to": y,
                "time": current_time,
                "duration": 0.0
            }

    def on_mouse_click(self, x, y, button, pressed):
        if self.is_recording:
            self.recorded_events.append({
                "type": "mouse_click",
                "x": x, "y": y,
                "button": str(button).split('.')[-1],
                "pressed": pressed,
                "time": time.time() - self.start_time
            })

    def on_mouse_scroll(self, x, y, dx, dy):
        if self.is_recording:
            self.recorded_events.append({
                "type": "mouse_scroll",
                "x": x, "y": y,
                "dx": dx, "dy": dy,
                "time": time.time() - self.start_time
            })

    def on_key_press(self, key):
        if self.is_recording:
            key_str = self.key_map.get(key, None)
            if key_str is None:
                try:
                    key_str = key.char
                except:
                    key_str = f"unknown_{key}"
            self.recorded_events.append({
                "type": "key_down",
                "key": key_str,
                "time": time.time() - self.start_time
            })

    def on_key_release(self, key):
        if self.is_recording:
            key_str = self.key_map.get(key, None)
            if key_str is None:
                try:
                    key_str = key.char
                except:
                    key_str = f"unknown_{key}"
            self.recorded_events.append({
                "type": "key_up",
                "key": key_str,
                "time": time.time() - self.start_time
            })

    def save_current_record(self):
        now = datetime.now()
        default_name = f"Record_{now.strftime('%Y%m%d_%H%M')}"
        name, ok = QInputDialog.getText(self, "保存脚本", "请输入名称：", text=default_name)
        if not ok or not name.strip():
            self.statusBar().showMessage("未保存：用户取消命名")
            return
        name = name.strip()

        cursor = self.db_conn.cursor()
        cursor.execute('''
            INSERT INTO record (name, events_json, event_count, description)
            VALUES (?, ?, ?, ?)
        ''', (name, json.dumps(self.recorded_events), len(self.recorded_events), ""))
        self.db_conn.commit()

        self.load_records()
        self.statusBar().showMessage(f"已保存：{name}")

    def play_selected_or_current(self):
        if self.is_playing:
            self.stop_play_flag.set()
            return

        selected = self.table.selectedItems()
        events_to_play = []

        if selected:
            row = selected[0].row()
            record_id = int(self.table.item(row, 0).text())
            cursor = self.db_conn.cursor()
            cursor.execute("SELECT events_json FROM record WHERE id = ?", (record_id,))
            result = cursor.fetchone()
            if result:
                events_to_play = json.loads(result[0])
            else:
                QMessageBox.critical(self, "错误", "无法加载选中的记录。")
                return
        else:
            if self.recorded_events:
                events_to_play = self.recorded_events
            else:
                QMessageBox.warning(self, "提示", "请先录制或选择一条记录。")
                return

        self.recorded_events = events_to_play
        self.start_playback()

    def start_playback(self):
        self.is_playing = True
        self.btn_play.setText("停止播放")
        self.btn_play.setStyleSheet("background-color: #f57c00; color: white;")
        self.statusBar().showMessage("正在播放... 按下 Esc 键可中断")

        self.stop_play_flag.clear()

        self.play_thread = threading.Thread(target=self.playback, daemon=True)
        self.play_thread.start()

        self.esc_listener = keyboard.Listener(on_release=self.on_esc_release)
        self.esc_listener.start()

    def on_esc_release(self, key):
        if key == keyboard.Key.esc and self.is_playing:
            self.stop_play_flag.set()
            return False

    def playback(self):
        if not self.recorded_events:
            self.is_playing = False
            QTimer.singleShot(0, self.reset_play_button)
            return

        prev_time = self.recorded_events[0]["time"] if self.recorded_events else 0
        current_time_str = datetime.now().strftime("%H%M%S")
        _current_input = ""

        for event in self.recorded_events:
            if self.stop_play_flag.is_set():
                break

            delay = event["time"] - prev_time
            if delay > 0:
                time.sleep(delay)

            prev_time = event["time"]

            try:
                etype = event["type"]
                if etype == "mouse_move":
                    # 平滑移动
                    duration = max(event["duration"], 0.01)
                    pyautogui.moveTo(event["x_to"], event["y_to"], duration=min(duration, 1.0))
                elif etype == "mouse_click":
                    button = {"left": "left", "right": "right", "middle": "middle"}.get(event["button"], "left")
                    if event["pressed"]:
                        pyautogui.mouseDown(button=button, x=event["x"], y=event["y"])
                    else:
                        pyautogui.mouseUp(button=button, x=event["x"], y=event["y"])
                elif etype == "mouse_scroll":
                    pyautogui.scroll(event["dy"], x=event["x"], y=event["y"])
                elif etype == "key_down":
                    key = event["key"]
                    if len(key) == 1 or key in ['space', 'backspace']:
                        if key == 'backspace' and _current_input:
                            _current_input = _current_input[:-1]
                        elif key != 'backspace':
                            _current_input += key
                        pyautogui.keyDown(key)
                    else:
                        pyautogui.keyDown(key)
                elif etype == "key_up":
                    key = event["key"]
                    if key == 'enter' and _current_input.endswith('_'):
                        _current_input = _current_input[:-1] + current_time_str
                        pyautogui.press('backspace')
                        pyautogui.typewrite(current_time_str)
                    elif len(key) == 1 or key in ['space']:
                        pass
                    else:
                        pyautogui.keyUp(key)
            except Exception as e:
                print(f"回放错误: {e}")
                continue

        self.is_playing = False
        QTimer.singleShot(0, self.reset_play_button)

    def reset_play_button(self):
        self.btn_play.setText("播放")
        self.btn_play.setStyleSheet("")
        self.statusBar().showMessage("播放完成")

    def load_records(self):
        cursor = self.db_conn.cursor()
        cursor.execute("SELECT id, name, event_count, created_at, description FROM record ORDER BY created_at DESC")
        records = cursor.fetchall()

        self.table.setRowCount(0)
        self.table.setRowCount(len(records))
        for row_idx, (rid, name, count, created, desc) in enumerate(records):
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(rid)))
            self.table.setItem(row_idx, 1, QTableWidgetItem(name))
            self.table.setItem(row_idx, 2, QTableWidgetItem(str(count)))
            self.table.setItem(row_idx, 3, QTableWidgetItem(created))
            self.table.setItem(row_idx, 4, QTableWidgetItem(desc or ""))

    def delete_record(self):
        selected = self.table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "提示", "请先选中要删除的记录。")
            return
        row = selected[0].row()
        record_id = int(self.table.item(row, 0).text())
        name = self.table.item(row, 1).text()

        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要删除记录 '{name}' 吗？此操作不可恢复！",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            cursor = self.db_conn.cursor()
            cursor.execute("DELETE FROM record WHERE id = ?", (record_id,))
            self.db_conn.commit()
            self.table.removeRow(row)
            self.statusBar().showMessage(f"已删除：{name}")

    def closeEvent(self, event):
        if self.mouse_listener and self.mouse_listener.is_alive():
            self.mouse_listener.stop()
        if self.keyboard_listener and self.keyboard_listener.is_alive():
            self.keyboard_listener.stop()
        if hasattr(self, 'esc_listener') and self.esc_listener.is_alive():
            self.esc_listener.stop()
        self.db_conn.close()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutoActionRecorder()
    window.show()
    sys.exit(app.exec_())