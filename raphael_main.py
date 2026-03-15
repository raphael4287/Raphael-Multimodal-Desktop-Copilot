import os
import datetime
import asyncio
import pvporcupine
import edge_tts
import pygame
import wave
import struct
import vlc
import json
import math
import requests
import webbrowser
import subprocess
import urllib.parse
import yt_dlp  # 用來搜尋與取得 YouTube 音訊串流
import vlc     # python-vlc，用來播放音訊
import psutil
import pyautogui
import pyperclip
import easyocr
import cv2
import numpy as np
import sys
import warnings
import dateparser
import time
import base64
import threading
from PIL import Image
from pvrecorder import PvRecorder
from openai import OpenAI
from dotenv import load_dotenv
from io import BytesIO
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(2)
from PyQt5.QtWidgets import (
    QMainWindow, QProgressBar, QTextEdit, QLineEdit,
    QPushButton, QFrame, QSplitter, QHBoxLayout,
    QVBoxLayout, QWidget, QLabel, QComboBox, QApplication
)
from PyQt5.QtCore import QTimer, pyqtSlot

import queue
speak_queue = queue.Queue()

# ==================== 應用路徑記憶 ====================
APP_PATHS_FILE = "app_paths.json"
app_paths = {}

if os.path.exists(APP_PATHS_FILE):
    try:
        with open(APP_PATHS_FILE, 'r', encoding='utf-8') as f:
            app_paths = json.load(f)
        print(f"[應用路徑] 已載入 {len(app_paths)} 個記住的應用/設定")
    except Exception as e:
        print(f"[應用路徑載入失敗] {e}")
        app_paths = {}

# 強制指定 VLC 安裝路徑（改成你實際的路徑！）
vlc_install_path = r"C:\Program Files\VideoLAN\VLC"   # ← 這裡改成你的路徑

# Python 3.8+ 推薦方式
if hasattr(os, 'add_dll_directory'):
    os.add_dll_directory(vlc_install_path)

# 強制載入 libvlc.dll 確認是否成功
try:
    ctypes.CDLL(os.path.join(vlc_install_path, "libvlc.dll"))
    print(f"[VLC 載入成功] 從 {vlc_install_path} 載入 libvlc.dll")
except Exception as e:
    print(f"[VLC 載入失敗] {e}")
    print("請確認：")
    print("1. VLC 已安裝在該路徑")
    print("2. 路徑是否正確（包含大小寫）")
    print("3. 是否安裝 64-bit VLC（Python 也要 64-bit）")

warnings.filterwarnings("ignore", category=UserWarning, module="torch")

# ==================== 全域初始化 ====================
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
PICO_KEY = os.getenv("PICOVOICE_ACCESS_KEY")
OWM_API_KEY = os.getenv("OPENWEATHER_API_KEY")
PPN_PATH = "Raphael.ppn"
TEMPLATE_DIR = "templates"          # 模板資料夾
DEFAULT_THRESHOLD = 0.73            # 基本門檻，視情況可調 0.68~0.80
SCALES = [0.6, 0.8, 1.0, 1.2, 1.5]  # 多尺度範圍

def find_template(template_name, threshold=DEFAULT_THRESHOLD, scales=SCALES):
    """
    在螢幕上尋找指定模板，返回中心座標 (x, y) 或 None
    template_name: 不含 .png 的檔名，例如 "chrome"、"保存"、"下一頁"
    """
    template_path = os.path.join(TEMPLATE_DIR, f"{template_name}.png")
    
    if not os.path.exists(template_path):
        print(f"[模板不存在] {template_path}")
        return None

    template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
    if template is None:
        print(f"[讀取失敗] {template_path}")
        return None

    # 截取螢幕並轉灰階
    screen = pyautogui.screenshot()
    screen_cv = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2GRAY)

    best_val = -1
    best_loc = None
    best_size = None

    for scale in scales:
        resized = cv2.resize(template, (0, 0), fx=scale, fy=scale)
        if resized.shape[0] > screen_cv.shape[0] or resized.shape[1] > screen_cv.shape[1]:
            continue

        res = cv2.matchTemplate(screen_cv, resized, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)

        if max_val > best_val and max_val >= threshold:
            best_val = max_val
            best_loc = max_loc
            best_size = resized.shape

    if best_loc is not None:
        h, w = best_size
        center_x = best_loc[0] + w // 2
        center_y = best_loc[1] + h // 2
        print(f"[匹配成功] {template_name} | 信心={best_val:.3f} | 位置=({center_x}, {center_y}) | scale={w/template.shape[1]:.2f}")
        return (center_x, center_y)
    
    print(f"[無匹配] {template_name} 信心最高僅 {best_val:.3f} < {threshold}")
    return None

reader = easyocr.Reader(['ch_tra', 'en'], gpu=True)

pygame.mixer.pre_init(44100, -16, 2, 512)
try:
    pygame.mixer.init()
except Exception as e:
    print(f"[警告] Pygame 初始化失敗，嘗試相容模式: {e}")
    os.environ['SDL_AUDIODRIVER'] = 'dsound'
    pygame.mixer.init()

# ==================== 工具描述 ====================
tools = [
    {"type": "function", "function": {"name": "get_current_time", "description": "問現在時間"}},
    {"type": "function", "function": {"name": "get_weather", "description": "查詢天氣", "parameters": {"type": "object", "properties": {"city": {"type": "string"}}}}},
    {"type": "function", "function": {"name": "control_computer", "description": "電腦控制", "parameters": {"type": "object", "properties": {"action": {"type": "string"}}}}},
    {"type": "function", "function": {"name": "open_software", "description": "開啟桌面軟體", "parameters": {"type": "object", "properties": {"app_name": {"type": "string"}}}}},
    {"type": "function", "function": {"name": "vision_click", "description": "點擊無文字圖標", "parameters": {"type": "object", "properties": {"target_description": {"type": "string"}}, "required": ["target_description"]}}},
    {"type": "function", "function": {"name": "text_click", "description": "點擊帶文字按鈕", "parameters": {"type": "object", "properties": {"target_text": {"type": "string"}}, "required": ["target_text"]}}},
    {"type": "function", "function": {"name": "set_timer", "description": "設定計時器", "parameters": {"type": "object", "properties": {"seconds": {"type": "integer"}, "label": {"type": "string"}}, "required": ["seconds", "label"]}}},
    {"type": "function", "function": {"name": "web_search", "description": "網路搜尋", "parameters": {"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]}}},
    {"type": "function", "function": {"name": "music_controller", "description": "音樂控制", "parameters": {"type": "object", "properties": {"action": {"type": "string"}, "song_name": {"type": "string"}, "platform": {"type": "string"}, "language": {"type": "string"}}}}},
    {"type": "function", "function": {"name": "toggle_subtitles", "description": "切換字幕顯示", "parameters": {"type": "object", "properties": {"action": {"type": "string", "enum": ["on", "off"]}}, "required": ["action"]}}},
    {"type": "function", "function": {"name": "shutdown_raphael", "description": "關閉系統"}},
    {"type": "function", "function": {"name": "screen_assistant", "description": "畫面助手", "parameters": {"type": "object", "properties": {"user_intent": {"type": "string"}, "generated_code": {"type": "string"}}, "required": ["user_intent"]}}},
    {"type": "function", "function": {"name": "get_system_status", "description": "系統狀態"}},
]

chat_history = [{"role": "system", "content": "你是拉菲爾。當使用者要求對畫面代碼除錯、優化或生成時，請使用 screen_assistant。"}]

# ==================== 字幕視窗 ====================
class SubtitleSignal(QObject):
    text_updated = pyqtSignal(str)

sub_signal = SubtitleSignal()

class RaphaelSubtitleWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.is_enabled = False
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        layout = QVBoxLayout()
        self.label = QLabel("")
        self.label.setFont(QFont("Microsoft JhengHei", 16, QFont.Bold))
        self.label.setStyleSheet("""
            color: #FFFFFF;
            background-color: rgba(0, 0, 0, 160);
            border-radius: 15px;
            padding: 7px 15px;
        """)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setWordWrap(True)
        layout.addWidget(self.label)
        self.setLayout(layout)
        self.hide()

        self.timer = QTimer()
        self.timer.timeout.connect(self.hide_subtitle)

    def hide_subtitle(self):
        self.label.setText("")
        self.hide()

    def display_text(self, text):
        if not self.is_enabled or not text: return
        self.label.setText(text)
        self.label.adjustSize()
        self.adjustSize()
        screen = QApplication.primaryScreen().geometry()
        new_width = min(screen.width() - 100, self.label.width() + 60)
        new_x = (screen.width() - new_width) // 2
        new_y = screen.height() - self.height() - 50
        self.setGeometry(new_x, new_y, new_width, self.height())
        self.show()
        display_time = 1000 + (len(text) * 300)
        self.timer.start(display_time)

# ==================== 控制中心 UI ====================
class RaphaelControlCenter(QMainWindow):

    chat_updated = pyqtSignal(str)
    need_path_dialog = pyqtSignal(str, str)  # app_name, prompt_text
    rms_updated = pyqtSignal(float)
    

    def __init__(self, subtitle_win):
        super().__init__()
        self.recorder = None
        self.porcupine = None
        self.sub_win = subtitle_win
        self.current_ppn_path = r"C:\Raphael_Bot\Raphael.ppn"  # 預設 PPN 路徑
        self.api_key_input = None  # 稍後在 init_ui 建立
        self.gpu_initialized = False
        try:
            from pynvml import nvmlInit
            nvmlInit()
            self.gpu_initialized = True
            print("[系統] GPU 監控初始化成功")
        except Exception as e:
            print(f"[系統] GPU 監控跳過: {e}")
        self.init_ui()
        self.mic_selector.currentIndexChanged.connect(self.on_mic_changed)

        self.stats_timer = QTimer()
        self.stats_timer.timeout.connect(self.refresh_stats)
        self.stats_timer.start(2000)
        self.chat_history_list = [{"role": "system", "content": "你是拉菲爾，一位專業助理。"}]
        self.rms_updated.connect(self.update_rms_display)
        self.chat_updated.connect(self.append_chat_message)
        self.need_path_dialog.connect(self.show_path_dialog)
        QTimer.singleShot(500, self.initialize_audio)  # 延遲 0.5 秒確保 UI 就緒


    # 然後在主執行緒（例如 RaphaelControlCenter 的 __init__ 後）啟動一個 worker
    def speak_worker():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        while True:
            try:
                text = speak_queue.get(timeout=1.0)
                loop.run_until_complete(speak(text))
            except queue.Empty:
                continue
            except Exception as e:
                print(f"語音 worker 錯誤: {e}")

    # 在 if __name__ == "__main__": 裡面
    threading.Thread(target=speak_worker, daemon=True).start()
    # 在 RaphaelControlCenter 類別裡的 show_path_dialog 修改版
    def show_path_dialog(self, app_name, prompt_text):
        from PyQt5.QtWidgets import QMessageBox, QFileDialog

        clean_name = app_name.lower().replace("幫我開啟", "").replace("打開", "").replace("開啟", "").strip()

        # 先顯示提示，並優先建議選「應用程式檔案」
        reply = QMessageBox.question(
            self,
            "找不到應用程式",
            f"{prompt_text}\n\n建議：直接選擇應用程式的執行檔（.exe 或 .lnk）\n\n"
            "要現在選擇執行檔嗎？\n（選「否」則會詢問資料夾）",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes   # 預設 Yes → 先選檔案
        )

        msg = ""

        if reply == QMessageBox.Yes:
            # 優先選擇單一執行檔
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                f"請選擇「{app_name}」的執行檔",
                os.path.expanduser("~"),
                "執行檔 (*.exe *.lnk);;所有檔案 (*.*)"
            )

            if file_path:
                try:
                    subprocess.Popen(f'start "" "{file_path}"', shell=True)
                    app_paths[clean_name] = file_path
                    with open(APP_PATHS_FILE, 'w', encoding='utf-8') as f:
                        json.dump(app_paths, f, ensure_ascii=False, indent=4)
                    msg = f"已開啟並記住：{os.path.basename(file_path)}"
                except Exception as e:
                    msg = f"開啟失敗：{str(e)}"
            else:
                msg = "已取消選擇執行檔。"

        elif reply == QMessageBox.No:
            # 次要選項：選擇資料夾
            folder = QFileDialog.getExistingDirectory(
                self,
                "選擇桌面或應用資料夾",
                os.path.expanduser("~")
            )
            if folder:
                app_paths["custom_desktop_path"] = folder
                with open(APP_PATHS_FILE, 'w', encoding='utf-8') as f:
                    json.dump(app_paths, f, ensure_ascii=False, indent=4)
                msg = "已記住新桌面路徑，下次搜尋會使用此資料夾。"
                # 可以選擇這裡再呼叫一次 open_software(app_name) 自動重試
            else:
                msg = "已取消選擇資料夾。"

        else:
            msg = "已取消操作。"
        self.chat_updated.emit(f"<b style='color:#82AAFF;'>拉菲爾：</b> {msg}")

    def append_chat_message(self, message):
        self.chat_history.append(message)
        # 強制捲動到底
        QTimer.singleShot(50, lambda: self.chat_history.verticalScrollBar().setValue(
            self.chat_history.verticalScrollBar().maximum()
        ))

    def initialize_audio(self):
        if not hasattr(self, 'current_ppn_path'):
            self.current_ppn_path = r"C:\Raphael_Bot\Raphael.ppn"
            print("[警告] current_ppn_path 未定義，已使用預設值")
        try:
            device_idx = self.mic_selector.currentIndex()
            device_name = self.mic_selector.currentText()
            print(f"[音訊初始化] 索引：{device_idx}，裝置名稱：{device_name}")

            # 停止舊的（安全處理）
            if self.recorder:
                try:
                    self.recorder.stop()
                    self.recorder.delete()
                except Exception as stop_e:
                    print(f"[停止舊 recorder 失敗，但繼續] {stop_e}")
                self.recorder = None

            # 重建 porcupine
            if self.porcupine:
                self.porcupine.delete()
            self.porcupine = pvporcupine.create(
                access_key=os.getenv("PICOVOICE_ACCESS_KEY"),
                keyword_paths=[self.current_ppn_path],
                sensitivities=[1]
            )
            frame_len = self.porcupine.frame_length
            print(f"[Porcupine] frame_length = {frame_len}")

            # 建立新 recorder
            self.recorder = PvRecorder(
                device_index=device_idx,
                frame_length=frame_len
            )
            self.recorder.start()
            print(f"[初始化成功] 使用裝置：{self.recorder.selected_device}")

            # 強制測試讀取 3 次
            for i in range(3):
                test_frame = self.recorder.read()
                if test_frame:
                    max_val = max(map(abs, test_frame))
                    print(f"[測試 {i+1}/3] 長度={len(test_frame)}, 最大振幅={max_val}")
                    if max_val > 0:
                        print("→ 測試成功！麥克風有輸入")
                        break
                else:
                    print(f"[測試 {i+1}/3] 讀取失敗（空 frame）")
                time.sleep(0.2)

        except Exception as e:
            print(f"[音訊初始化失敗] {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()

    def select_ppn_file(self):
        from PyQt5.QtWidgets import QFileDialog

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "選擇 PPN 喚醒詞模型檔案",
            r"C:\Raphael_Bot",  # 預設開啟目錄
            "PPN Files (*.ppn);;All Files (*)"
        )

        if file_path:
            self.current_ppn_path = file_path
            self.ppn_label.setText(f"目前：{os.path.basename(file_path)}")
            print(f"[PPN 變更] 新路徑：{file_path}")
            
            # 立即重新初始化音訊（使用新 PPN）
            self.initialize_audio()

    def apply_new_api_key(self):
        new_key = self.api_key_input.text().strip()
        if not new_key:
            print("[警告] API Key 為空")
            return

        # 更新環境變數（供 pvporcupine 使用）
        os.environ["PICOVOICE_ACCESS_KEY"] = new_key
        print(f"[API Key 套用] 新 Key 已設定（長度：{len(new_key)}）")

        # 重新初始化音訊（使用新 Key）
        self.initialize_audio()

    def on_mic_changed(self, index):
        print(f"[麥克風變更] 新索引：{index} ({self.mic_selector.currentText()})")
        
        # 先安全停止舊的
        if self.recorder:
            try:
                self.recorder.stop()
                self.recorder.delete()
            except Exception as stop_err:
                print(f"[停止舊 recorder 失敗] {stop_err}")
            self.recorder = None

        # 強制延遲，讓 Windows 音訊堆疊釋放（關鍵！）
        time.sleep(1.5)  # 1.5 秒通常足夠

        # 重試 2 次初始化
        for attempt in range(2):
            try:
                self.initialize_audio()
                print(f"[變更成功] 已使用新裝置")
                return
            except Exception as e:
                print(f"[變更重試 {attempt+1}/2] 失敗：{e}")
                time.sleep(0.5)
        print("[變更最終失敗] 請重啟程式或檢查麥克風權限")
    
    def update_rms_display(self, rms_value):
        color = "#82AAFF"
        status = "安靜"
        if rms_value > 800:
            color = "red"
            status = "大聲"
        elif rms_value > 300:
            color = "orange"
            status = "正常"
        
        self.rms_label.setText(f"RMS: {rms_value:.1f} | 音量: {status}")
        self.rms_label.setStyleSheet(f"color: {color}; font-size: 14px; font-weight: bold;")
        
        # 進度條顯示（可加 log 縮放讓它更好看）
        display_value = min(int(rms_value * 1.5), 1200)  # 放大一點讓低音量也看得見
        self.rms_bar.setValue(display_value)

    def init_ui(self):
        self.setWindowTitle("Raphael AI - 指令控制中心")
        self.resize(1000, 700)
        self.setStyleSheet("""
            QMainWindow { background-color: #0F111A; }
            QLabel { color: #82AAFF; font-weight: bold; }
            QProgressBar { border: 1px solid #333; border-radius: 5px; text-align: center; color: white; background-color: #1A1C25; }
            QProgressBar::chunk { background-color: #82AAFF; }
            QTextEdit { background-color: #1A1C25; color: #D6DEEB; border: none; border-radius: 10px; font-size: 14px; }
            QLineEdit { background-color: #242936; color: white; border: 1px solid #333; border-radius: 20px; padding: 10px 20px; }
            QPushButton#SendBtn { background-color: #82AAFF; color: #0F111A; border-radius: 20px; font-weight: bold; }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # 左側面板
        left_panel = QVBoxLayout()
        left_panel.addWidget(QLabel("📊 硬體監控系統"))
        self.cpu_bar = QProgressBar(); self.cpu_bar.setRange(0, 100)
        self.ram_bar = QProgressBar(); self.ram_bar.setRange(0, 100)
        self.gpu_bar = QProgressBar(); self.gpu_bar.setRange(0, 100)
        left_panel.addWidget(QLabel("CPU Usage")); left_panel.addWidget(self.cpu_bar)
        left_panel.addWidget(QLabel("RAM Usage")); left_panel.addWidget(self.ram_bar)
        left_panel.addWidget(QLabel("GPU Usage")); left_panel.addWidget(self.gpu_bar)

        left_panel.addSpacing(30)
        left_panel.addWidget(QLabel("🎤 語音輸入裝置"))
        self.mic_selector = QComboBox()
        self.mic_selector.addItems(PvRecorder.get_available_devices())
        self.mic_selector.setStyleSheet("background: #242936; color: white; padding: 5px;")
        left_panel.addWidget(self.mic_selector)
        # ── 新增：即時 RMS 顯示 ──
        left_panel.addSpacing(20)
        left_panel.addWidget(QLabel("即時音量監測 (RMS)"))
        
        self.rms_label = QLabel("RMS: 0.0 | 音量: 安靜")
        self.rms_label.setStyleSheet("color: #82AAFF; font-size: 14px; font-weight: bold;")
        left_panel.addWidget(self.rms_label)

        self.rms_bar = QProgressBar()
        self.rms_bar.setRange(0, 1200)  # 調高一點範圍，避免太快滿
        self.rms_bar.setValue(0)
        self.rms_bar.setFormat("RMS: %v")
        self.rms_bar.setStyleSheet("""
            QProgressBar { background-color: #1A1C25; border: 1px solid #333; }
            QProgressBar::chunk { background-color: #82AAFF; }
        """)
        left_panel.addWidget(self.rms_bar)

        left_panel.addSpacing(20)
        self.api_status = QLabel("● OpenAI API: Connected")
        self.api_status.setStyleSheet("color: #C3E88D;")
        left_panel.addWidget(self.api_status)
        left_panel.addStretch()
        main_layout.addLayout(left_panel, 1)
        left_panel.addSpacing(10)
        self.reinit_btn = QPushButton("重新初始化麥克風")
        self.reinit_btn.clicked.connect(self.initialize_audio)
        left_panel.addWidget(self.reinit_btn)
        left_panel.addSpacing(20)
        left_panel.addWidget(QLabel("喚醒詞模型 (.ppn)"))

        # 顯示目前 PPN 路徑
        self.ppn_label = QLabel(f"目前：{os.path.basename(self.current_ppn_path)}")
        self.ppn_label.setStyleSheet("color: #82AAFF; font-size: 12px;")
        left_panel.addWidget(self.ppn_label)

        # 選擇 PPN 按鈕
        self.select_ppn_btn = QPushButton("選擇 PPN 檔案")
        self.select_ppn_btn.clicked.connect(self.select_ppn_file)
        left_panel.addWidget(self.select_ppn_btn)

        left_panel.addSpacing(20)
        left_panel.addWidget(QLabel("自訂桌面路徑（用來開啟應用）"))

        # 顯示目前路徑
        self.desktop_path_label = QLabel("目前未設定")
        self.desktop_path_label.setStyleSheet("color: #82AAFF; font-size: 12px;")
        left_panel.addWidget(self.desktop_path_label)

        # 選擇按鈕
        self.select_desktop_btn = QPushButton("選擇桌面資料夾")
        self.select_desktop_btn.clicked.connect(self.select_desktop_path)
        left_panel.addWidget(self.select_desktop_btn)

        # 如果有記住的路徑，顯示出來
        if "custom_desktop_path" in app_paths:
            self.desktop_path_label.setText(f"目前：{app_paths['custom_desktop_path']}")


        left_panel.addSpacing(10)
        left_panel.addWidget(QLabel("Picovoice API Key"))

        # API Key 輸入框
        self.api_key_input = QLineEdit(os.getenv("PICOVOICE_ACCESS_KEY", ""))
        self.api_key_input.setEchoMode(QLineEdit.Password)  # 隱藏輸入（像密碼）
        self.api_key_input.setPlaceholderText("輸入你的 Picovoice Access Key")
        left_panel.addWidget(self.api_key_input)

        # 套用按鈕
        self.apply_api_btn = QPushButton("套用 API Key")
        self.apply_api_btn.clicked.connect(self.apply_new_api_key)
        left_panel.addWidget(self.apply_api_btn)

        # left_panel.addSpacing(10)

        # 右側面板
        right_panel = QVBoxLayout()
        self.chat_history = QTextEdit()
        self.chat_history.setReadOnly(True)
        self.chat_history.setPlaceholderText("拉菲爾正在待命...")
        right_panel.addWidget(self.chat_history)

        input_row = QHBoxLayout()
        self.text_input = QLineEdit()
        self.text_input.setPlaceholderText("輸入指令或問題...")
        self.text_input.returnPressed.connect(self.handle_text_submit)
        self.send_button = QPushButton("發送")
        self.send_button.setObjectName("SendBtn")
        self.send_button.setFixedSize(80, 40)
        self.send_button.clicked.connect(self.handle_text_submit)
        input_row.addWidget(self.text_input)
        input_row.addWidget(self.send_button)
        right_panel.addLayout(input_row)
        main_layout.addLayout(right_panel, 3)

    def select_desktop_path(self):
        from PyQt5.QtWidgets import QFileDialog

        folder = QFileDialog.getExistingDirectory(
            self,
            "選擇您的桌面或常用應用資料夾",
            app_paths.get("custom_desktop_path", os.path.expanduser("~"))
        )

        if folder:
            app_paths["custom_desktop_path"] = folder
            with open(APP_PATHS_FILE, 'w', encoding='utf-8') as f:
                json.dump(app_paths, f, ensure_ascii=False, indent=4)
            self.desktop_path_label.setText(f"目前：{folder}")
            print(f"[桌面路徑已設定] {folder}")
        else:
            print("[使用者取消選擇桌面路徑]")

    def reset_desktop_path(self):
        if "custom_desktop_path" in app_paths:
            del app_paths["custom_desktop_path"]
            with open(APP_PATHS_FILE, 'w', encoding='utf-8') as f:
                json.dump(app_paths, f, ensure_ascii=False, indent=4)
            print("[桌面路徑已重設為預設]")

    def refresh_stats(self):
        self.cpu_bar.setValue(int(psutil.cpu_percent()))
        self.ram_bar.setValue(int(psutil.virtual_memory().percent))
        if self.gpu_initialized:
            try:
                from pynvml import nvmlDeviceGetHandleByIndex, nvmlDeviceGetUtilizationRates
                handle = nvmlDeviceGetHandleByIndex(0)
                util = nvmlDeviceGetUtilizationRates(handle)
                self.gpu_bar.setValue(util.gpu)
            except:
                pass
        else:
            self.gpu_bar.setValue(15)

    def closeEvent(self, event):
        if self.gpu_initialized:
            try:
                from pynvml import nvmlShutdown
                nvmlShutdown()
            except: pass
        event.accept()

    def handle_text_submit(self):
        query = self.text_input.text().strip()
        if query:
            self.chat_history.append(f"<b style='color:#C792EA;'>你：</b> {query}")
            self.text_input.clear()
            threading.Thread(target=self.sync_process_wrapper, args=(query,), daemon=True).start()

    def sync_process_wrapper(self, query):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(self.process_ai_logic(query))

    async def process_ai_logic(self, user_text):
        if not user_text.strip():
            await speak("沒聽清楚，再說一次好嗎？")
            return

        print(f"→ 使用者說：{user_text}")

        # 把使用者輸入加進歷史
        self.chat_history_list.append({"role": "user", "content": user_text})

        # 判斷模型
        coding_keywords = ["程式", "寫", "代碼", "code", "編寫", "生成", "腳本", "python", "html", "cpp", "java"]
        is_coding = any(k in user_text.lower() for k in coding_keywords)
        selected_model = "gpt-4o" if is_coding else "gpt-4o-mini"
        print(f"系統判定任務類型，使用模型: {selected_model}")

        try:
            # 第一輪：帶工具
            response = client.chat.completions.create(
                model=selected_model,
                messages=self.chat_history_list,
                tools=tools,
                tool_choice="auto"
            )

            msg = response.choices[0].message
            self.chat_history_list.append(msg)  # 先存 assistant 訊息（包含 tool_calls）

            final_ans = ""

            if msg.tool_calls:
                tool_messages = []  # 確保這裡一定有定義

                for tc in msg.tool_calls:
                    fn = tc.function.name
                    args = json.loads(tc.function.arguments)

                    print(f"[Tool Call] {fn} → {args}")

                    res = ""

                    if fn == "get_current_time":
                        res = get_current_time()
                    elif fn == "get_weather":
                        res = get_weather(args.get("city"))
                    elif fn == "control_computer":
                        res = control_computer(args.get("action"))
                    elif fn == "open_software":
                        res = open_software(args.get("app_name"))

                        if "[NEED_FILEDIALOG]" in res:
                            prompt_text = res.replace("[NEED_FILEDIALOG]", "").strip()
                            # 發送信號彈窗
                            self.need_path_dialog.emit(args.get("app_name"), prompt_text)
                            res = "正在詢問您選擇應用路徑，請稍等..."
                            # self.chat_updated.emit(f"<b style='color:#82AAFF;'>拉菲爾：</b> {res}")
                            await speak(res)
                            
                            # 👇 【修復 Error 400】必須把工具執行結果存入歷史，再 return
                            self.chat_history_list.append({
                                "role": "tool",
                                "tool_call_id": tc.id,
                                "name": fn,
                                "content": "已跳出視窗等待使用者選擇路徑，本次對話流程暫停。"
                            })
                            
                            return  # 結束本次 loop，等待彈窗結果
                    elif fn == "vision_click":
                        res = vision_click(args.get("target_description"))
                    elif fn == "text_click":
                        res = text_click(args.get("target_text"))
                    elif fn == "set_timer":
                        await set_timer(args.get("seconds"), args.get("label"))
                        res = f"已設定 {args.get('label')} 在 {args.get('seconds')} 秒後。"
                    elif fn == "web_search":
                        res = web_search(args.get("query"))
                    elif fn == "music_controller":
                        res = music_controller(args)
                    elif fn == "toggle_subtitles":
                        res = toggle_subtitles(args.get("action"))
                    elif fn == "shutdown_raphael":
                        res = shutdown_raphael()
                        await speak(res)
                        QApplication.quit()
                        return
                    elif fn == "screen_assistant":
                        res = screen_assistant(args.get("user_intent"))
                    elif fn == "get_system_status":
                        res = get_system_status()

                    # 只在有真實結果時才加 tool 訊息
                    tool_messages.append({
                        "role": "tool",
                        "tool_call_id": tc.id,
                        "name": fn,
                        "content": res or "已執行"
                    })

                # 把所有 tool 結果一次加進歷史（保持順序）
                self.chat_history_list.extend(tool_messages)

                # 第二輪總結（不帶 tools）
                final_response = client.chat.completions.create(
                    model=selected_model,
                    messages=self.chat_history_list
                )
                final_ans = final_response.choices[0].message.content
            else:
                final_ans = msg.content

            # 最終存歷史並回覆
            self.chat_history_list.append({"role": "assistant", "content": final_ans})
            await speak(final_ans)

            # UI 更新
            # self.chat_updated.emit(f"<b style='color:#82AAFF;'>拉菲爾：</b> {final_ans}")

        except Exception as e:
            print(f"AI 處理錯誤: {e}")
            await speak("抱歉，剛剛出了點問題，能再說一次嗎？")
    

# ==================== 工具函數實作 ====================
def play_sound(filename):
    try:
        if os.path.exists(filename):
            pygame.mixer.Sound(filename).play()
        elif filename == "notify.mp3": print("\a")
    except: pass

def get_current_time():
    return f"現在時間是 {datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}"

def get_weather(city="Chiayi"):
    if not city: city = "嘉義"
    try:
        clean_city = city.replace("市", "").replace("縣", "")
        city_map = {"台北": "Taipei", "新北": "New Taipei", "桃園": "Taoyuan", "台中": "Taichung", "台南": "Tainan", "高雄": "Kaohsiung", "嘉義": "Chiayi", "雲林": "Yunlin"}
        search_city = city_map.get(clean_city, clean_city)
        url = f"http://api.openweathermap.org/data/2.5/weather?q={search_city}&appid={OWM_API_KEY}&units=metric&lang=zh_tw"
        res = requests.get(url, timeout=5).json()
        if res.get("cod") == 200:
            return f"{city}目前{res['weather'][0]['description']}，氣溫 {res['main']['temp']} 度。"
        return f"找不到 {city} 的天氣資訊。"
    except Exception as e:
        print(f"天氣錯誤: {e}")
        return "天氣查詢失敗。"

def control_computer(action):
    if action == "shutdown": os.system("shutdown /s /t 15"); return "十秒後關機。"
    if action == "restart": os.system("shutdown /r /t 15"); return "十秒後重啟。"
    if action == "cancel": os.system("shutdown /a"); return "已取消關機。"
    if action == "sleep": os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0"); return "進入睡眠。"
    return "無效操作。"

# ==================== 工具函數：open_software ====================
def open_software(app_name):
    global app_paths

    clean_name = app_name.replace("幫我開啟", "").replace("打開", "").replace("開啟", "").strip().lower()

    # 1. 先檢查是否已經記住單一應用路徑
    if clean_name in app_paths and clean_name != "custom_desktop_path":
        saved_path = app_paths[clean_name]
        if os.path.exists(saved_path):
            try:
                subprocess.Popen(f'start "" "{saved_path}"', shell=True)
                return f"已開啟記住的應用：{clean_name}"
            except Exception as e:
                del app_paths[clean_name]
                return f"[NEED_FILEDIALOG]記住的路徑失效，請重新選擇「{clean_name}」的執行檔。"
        else:
            del app_paths[clean_name]
            return f"[NEED_FILEDIALOG]記住的路徑不存在，請重新選擇「{clean_name}」的執行檔。"

    # 2. 檢查是否有自訂桌面路徑
    desktop_path = app_paths.get("custom_desktop_path")
    if desktop_path and os.path.exists(desktop_path):
        try:
            for item in os.listdir(desktop_path):
                name_without_ext = os.path.splitext(item)[0].lower()
                if clean_name in name_without_ext and item.lower().endswith((".lnk", ".exe", ".url")):
                    full_path = os.path.join(desktop_path, item)
                    if os.path.exists(full_path):
                        try:
                            subprocess.Popen(f'start "" "{full_path}"', shell=True)
                            app_paths[clean_name] = full_path  # 記住這次找到的路徑
                            with open(APP_PATHS_FILE, 'w', encoding='utf-8') as f:
                                json.dump(app_paths, f, ensure_ascii=False, indent=4)
                            return f"已從自訂桌面開啟：{item} (已記住)"
                        except Exception as e:
                            return f"[NEED_FILEDIALOG]開啟失敗：{e}。請手動選擇應用程式。"
        except Exception as e:
            return f"[NEED_FILEDIALOG]自訂桌面搜尋失敗：{e}。請手動選擇。"

    # 3. 都找不到 → 要求彈出選擇視窗（優先選檔案）
    return f"[NEED_FILEDIALOG]找不到「{clean_name}」，請選擇應用程式檔案（.exe 或 .lnk）或資料夾。"

def web_search(query):
    webbrowser.open(f"https://www.google.com/search?q={query}")
    return f"已為您搜尋：{query}"

def get_system_status():
    return f"CPU: {psutil.cpu_percent()}%，記憶體: {psutil.virtual_memory().percent}%"

# 替換原本的 set_timer
async def set_timer(time_str, label="計時器"):
    """支援自然語言與絕對時間（中文超強）"""
    try:
        # 先讓 dateparser 解析
        dt = dateparser.parse(time_str, languages=['zh', 'zh-TW', 'en'])
        if not dt:
            # 如果沒解析成功，當成相對秒數
            seconds = int(''.join(filter(str.isdigit, time_str))) or 10
            await asyncio.sleep(seconds)
            await speak(f"「{label}」時間到！")
            return

        # 計算距離現在的秒數
        now = datetime.datetime.now()
        delta = (dt - now).total_seconds()
        if delta < 0:
            delta += 86400  # 隔天同時間

        print(f"[計時器] 預計在 {dt.strftime('%H:%M:%S')} 觸發（還有 {int(delta)} 秒）")
        await asyncio.sleep(delta)
        await speak(f"「{label}」時間到！")
    except Exception as e:
        await speak("計時器格式我沒聽懂，請再說一次～")

def move_with_dynamic_speed(start_x, start_y, target_x, target_y):
    distance = math.sqrt((target_x - start_x)**2 + (target_y - start_y)**2)
    screen_w = pyautogui.size()[0]
    duration = 0.4 + (distance / screen_w) * 0.8
    duration = min(1.2, max(0.4, duration))
    pyautogui.moveTo(target_x, target_y, duration=duration, tween=pyautogui.easeInOutQuart)
    time.sleep(0.1)

def vision_click(target_description):
    """
    用法：說「點擊 chrome」「點擊設定齒輪」「點擊紅色 X」
    對應 templates/chrome.png、templates/設定齒輪.png 等
    """
    print(f"[vision_click] 目標：{target_description}")
    
    # 簡單正規化檔名
    name = target_description.lower().replace(" ", "").replace("圖示", "").replace("按鈕", "")
    
    center = find_template(name)
    if center:
        move_with_dynamic_speed(*pyautogui.position(), *center)
        pyautogui.click()
        return f"已點擊圖示/按鈕「{target_description}」"
    else:
        return f"找不到「{target_description}」的模板，請確認 templates/{name}.png 存在"

def text_click(target_text):
    """
    用法：說「點擊 儲存」「點擊 確認」「點擊 取消」
    對應 templates/儲存.png、templates/確認.png 等
    """
    print(f"[text_click] 目標文字：{target_text}")
    
    # 檔名直接使用輸入文字（可再客製規則）
    name = target_text.strip()
    
    center = find_template(name, threshold=0.70)  # 文字按鈕可稍降低門檻
    if center:
        move_with_dynamic_speed(*pyautogui.position(), *center)
        pyautogui.click()
        return f"已點擊文字按鈕「{target_text}」"
    else:
        return f"找不到「{target_text}」的按鈕模板"

def screen_assistant(user_intent):
    """
    原版完整功能：畫面文字/程式碼全自動助手
    支援：更改、刪除、優化、生成、擴充、翻譯、除錯
    使用 EasyOCR + GPT 多輪精準編輯 + 物理退格
    """
    print(f"[螢幕助手] 啟動執行流程，原始指令：{user_intent}")
    
    # --- 第一階段：意圖與語言解析 ---
    intent_prompt = (
        f"妳是任務解析器。請將需求轉換為 JSON 工單：『{user_intent}』\n\n"
        "回傳 JSON 格式：{\"branch\": \"...\", \"action\": \"...\", \"language\": \"...\", \"anchor\": \"...\", \"target_goal\": \"...\"}"
    )
    
    try:
        intent_res = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": intent_prompt}],
            response_format={"type": "json_object"}
        )
        it = json.loads(intent_res.choices[0].message.content)
        t_branch = it.get('branch')
        t_act = it.get('action')
        t_lang = it.get('language')
        t_anchor = it.get('anchor')
        t_goal = it.get('target_goal')
        print(f"[工單確認] 分支:{t_branch} | 語系:{t_lang} | 動作:{t_act} | 定位:{t_anchor}")
    except Exception as e:
        return f"意圖解析失敗：{e}"

    # --- 第二階段：任務點計數 (全量偵測機制) ---
    screenshot = pyautogui.screenshot()
    img_array = np.array(screenshot.convert('L'))
    initial_results = reader.readtext(img_array)
    screen_h = pyautogui.size()[1]
    filtered_init = [res for res in initial_results if 60 < res[0][0][1] < (screen_h - 60)]
    
    # 判斷是否為全量模式
    is_all_task = any(word in user_intent for word in ["所有", "全部", "所有內容", "all", "every", "每個"])
    
    if is_all_task:
        count_prompt = (
            f"任務：在『{t_anchor}』處執行『{t_act}』。目標：{t_goal}\n"
            f"目前螢幕內容：\n" + "\n".join([f"ID {i}: [{res[1]}]" for i, res in enumerate(filtered_init)]) + "\n"
            "請計算共有幾個 ID 需要處理？請只回傳數字。"
        )
        try:
            count_res = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": count_prompt}]
            )
            task_count = int(''.join(filter(str.isdigit, count_res.choices[0].message.content)))
            print(f"[計數器] 偵測到全量模式，共 {task_count} 個目標點。")
        except:
            task_count = 1
    else:
        task_count = 1
        print("[系統] 單次執行模式，預設處理 1 個匹配目標。")

    # --- 第三階段：精確循環執行（原版核心）---
    for i in range(task_count):
        print(f"[進度] 執行中：{i+1}/{task_count}...")
        
        # 每一輪重新截圖確保座標正確
        curr_screenshot = pyautogui.screenshot()
        curr_img = np.array(curr_screenshot.convert('L'))
        curr_results = reader.readtext(curr_img)
        curr_filtered = [res for res in curr_results if 60 < res[0][0][1] < (screen_h - 60)]
        curr_content = "\n".join([f"ID {j}: [{res[1]}]" for j, res in enumerate(curr_filtered)])

        decision_prompt = (
            f"任務：{t_act} ({t_branch})。語系：{t_lang}。目標：{t_goal}\n"
            f"螢幕內容：\n{curr_content}\n"
            "請指派下一個『尚未處理過』的 ID 並提供寫入內容。\n"
            "回傳 JSON：{\"line_idx\": ID, \"content\": \"...\", \"reason\": \"...\"}"
        )

        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "妳是精密自動化助理，負責計算內容並回傳 JSON。"},
                    {"role": "user", "content": decision_prompt}
                ],
                response_format={"type": "json_object"}
            )
            data = json.loads(response.choices[0].message.content)
            
            idx = data.get("line_idx")
            final_content = data.get("content", "").strip().lstrip('[').rstrip(']')

            if idx is not None and idx < len(curr_filtered):
                bbox = curr_filtered[idx][0]
                original_text = curr_filtered[idx][1]
                real_len = len(original_text)

                # 定位並點擊
                target_x = int(bbox[1][0])
                target_y = int((bbox[1][1] + bbox[2][1]) / 2)
                pyautogui.click(target_x, target_y)
                time.sleep(0.3)

                # 決定是否要刪除原文字
                should_delete = t_act in ['delete', 'optimize', 'translate', 'debug', 'replace'] or '翻譯' in user_intent
                if should_delete:
                    print(f"[執行] 精確物理退格：{real_len} 個字元")
                    for _ in range(real_len):
                        pyautogui.press('backspace')

                # 貼上新內容
                if final_content:
                    pyperclip.copy(final_content)
                    pyautogui.hotkey('ctrl', 'v')
                
                print(f"[完成] ID {idx}：{data.get('reason')}")
                time.sleep(1.2)
            else:
                print("[跳過] 無法定位匹配的 ID。")
                break
        except Exception as e:
            print(f"[錯誤] 循環執行中斷：{e}")
            break

    return f"拉菲爾：已完成「{t_act}」任務，處理點共計 {task_count} 處。"

player = None

def music_controller(args):
    global player
    action = args.get('action')
    song_name = args.get('song_name', '')

    if action == 'play' and song_name:
        try:
            with yt_dlp.YoutubeDL({'quiet': True, 'format': 'bestaudio'}) as ydl:
                info = ydl.extract_info(f"ytsearch:{song_name}", download=False)
                url = info['entries'][0]['url'] if 'entries' in info else info['url']

            if player is None:
                player = vlc.MediaPlayer()
            else:
                player.stop()

            player.set_media(vlc.Media(url))
            player.play()
            return f"正在為您播放：{info['entries'][0]['title'] if 'entries' in info else info['title']}"
        except Exception as e:
            return f"播放失敗：{e}"

    elif action in ['stop', 'next', 'previous']:
        if player:
            if action == 'stop': player.stop()
            elif action == 'next': player.stop()  # 可再擴充
        return f"已執行 {action}"
    
    return "音樂控制已執行。"

def toggle_subtitles(action):
    if action == "on":
        subtitle_win.is_enabled = True
        return "字幕已開啟。"
    elif action == "off":
        subtitle_win.is_enabled = False
        subtitle_win.hide()
        return "字幕已關閉。"
    return "無效指令。"

def shutdown_raphael():
    return "系統即將關閉，再見！"

# ==================== 語音相關 ====================
async def speak(text):
    print(f"拉菲爾：{text}")
    sub_signal.text_updated.emit(text)
    try:
        pygame.mixer.music.unload()
        await edge_tts.Communicate(text, "zh-TW-YunJheNeural").save("res.mp3")
        pygame.mixer.music.load("res.mp3")
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy(): await asyncio.sleep(0.1)
    except Exception as e:
        print(f"語音播放失敗: {e}")

async def get_voice_input(recorder, timeout_seconds=8):
    print(f"[聆聽中...] 等待 {timeout_seconds} 秒")
    start_time = time.time()
    audio_data = []
    while (time.time() - start_time) < timeout_seconds:
        frame = recorder.read()
        audio_data.extend(frame)
        if len(audio_data) >= 1600:
            rms = math.sqrt(sum(s**2 for s in audio_data[-1600:]) / 1600)
            if rms > 300:
                print("[偵測到聲音]")
                play_sound("reply_notify.mp3")
                while len(audio_data) < 110 * recorder.frame_length:
                    audio_data.extend(recorder.read())
                with wave.open("req.wav", 'wb') as wf:
                    wf.setnchannels(1); wf.setsampwidth(2); wf.setframerate(16000)
                    wf.writeframes(struct.pack('<' + ('h' * len(audio_data)), *audio_data))
                try:
                    with open("req.wav", "rb") as f:
                        return client.audio.transcriptions.create(model="whisper-1", file=f).text.strip()
                except: return ""
    print("[逾時]")
    await speak("未聽到指令。")
    play_sound("Stop_listen_notify.mp3")
    return ""

async def main():
    global control_panel

    print("\n=== Raphael 待命中... ===\n")

    while True:
        try:
            if not control_panel.recorder:
                print("[錯誤] recorder 未初始化，等待 1 秒...")
                await asyncio.sleep(1.0)
                continue

            frame = control_panel.recorder.read()
            if not frame or len(frame) == 0:
                await asyncio.sleep(0.01)
                continue

            triggered = False

            if control_panel.porcupine is not None:
                keyword_index = control_panel.porcupine.process(frame)
                if keyword_index >= 0:
                    print(f"[喚醒成功] index={keyword_index} @ {datetime.datetime.now().strftime('%H:%M:%S.%f')[:-3]}")
                    triggered = True
            else:
                audio_np = np.array(frame, dtype=np.float32)
                rms = np.sqrt(np.mean(audio_np ** 2))
                if rms > 300:
                    print(f"[音量觸發] RMS = {rms:.1f} @ {datetime.datetime.now().strftime('%H:%M:%S.%f')[:-3]}")
                    triggered = True

            # RMS 顯示
            audio_np = np.array(frame, dtype=np.float32)
            rms = np.sqrt(np.mean(audio_np ** 2)) if len(audio_np) > 0 else 0.0
            control_panel.rms_updated.emit(rms)

            if triggered:
                print("=== 觸發成功，開始錄音 ===")
                play_sound("notify.mp3")
                await speak("我在，請說～")
                user_text = await get_voice_input(control_panel.recorder, timeout_seconds=8)
                print("[觸發時間]", datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3])

                if user_text and user_text.strip():
                    print(f"→ 使用者說：{user_text}")
                    await control_panel.process_ai_logic(user_text)
                else:
                    print("[無有效語音輸入]")

        except Exception as e:
            print(f"[主迴圈異常] {e}")
            await asyncio.sleep(0.1)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    subtitle_win = RaphaelSubtitleWindow()
    global control_panel
    control_panel = RaphaelControlCenter(subtitle_win)
    control_panel.show()

    sub_signal.text_updated.connect(subtitle_win.display_text)
    sub_signal.text_updated.connect(
        lambda t: (
            control_panel.chat_history.append(f"<b style='color:#82AAFF;'>拉菲爾：</b> {t}"),
            QTimer.singleShot(50, lambda: control_panel.chat_history.verticalScrollBar().setValue(
                control_panel.chat_history.verticalScrollBar().maximum()
            ))
        )
    )

    threading.Thread(target=lambda: asyncio.run(main()), daemon=True).start()
    sys.exit(app.exec_())
