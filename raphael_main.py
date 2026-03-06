import os
import datetime
import asyncio
import pvporcupine
import edge_tts
import pygame
import wave
import struct
import json
import math
import requests
import webbrowser
import subprocess
import urllib.parse
import psutil
import pyautogui
import pyperclip
import easyocr
import numpy as np
import sys
import warnings
import time
import base64
import math
import threading
from PIL import Image 
from pvrecorder import PvRecorder
from openai import OpenAI
from dotenv import load_dotenv
from io import BytesIO
from PyQt5.QtWidgets import QApplication, QLabel, QWidget, QVBoxLayout
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(2)
class SubtitleSignal(QObject):
    # 定義一個信號，用來接收要顯示的文字
    text_updated = pyqtSignal(str)

class RaphaelSubtitleWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.is_enabled = False
        # 置頂、無邊框、點擊穿透、不在工具列顯示
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool | Qt.WindowTransparentForInput)
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        layout = QVBoxLayout()
        self.label = QLabel("")
        # 科技感字體與顏色
        self.label.setFont(QFont("Microsoft JhengHei",16, QFont.Bold))
        self.label.setStyleSheet("""
            color: #FFFFFF;
            background-color: rgba(0, 0, 0, 160); /* 稍微加深一點點更有質感 */
            border-radius: 15px;
            padding: 7px 15px; /* 增加左右間距 */
        """)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setWordWrap(True)
        
        layout.addWidget(self.label)
        self.setLayout(layout)

        # 初始化隱藏
        self.hide()
        
        self.timer = QTimer()
        # 修正點：時間到時呼叫 self.hide，讓黑幕徹底消失
        self.timer.timeout.connect(self.hide_subtitle)

    def hide_subtitle(self):
        self.label.setText("")
        self.hide()

    def display_text(self, text):
        if not text: return
        if not self.is_enabled or not text: return
        
        self.label.setText(text)
        
        # --- 動態調整視窗大小與位置 ---
        # 根據文字內容重新計算標籤所需大小
        self.label.adjustSize()
        self.adjustSize() 
        
        screen = QApplication.primaryScreen().geometry()
        # 將視窗定位在螢幕底部中央 (距離底部 100 像素)
        new_width = min(screen.width() - 100, self.label.width() + 60)
        new_x = (screen.width() - new_width) // 2
        new_y = screen.height() - self.height() - 50
        
        self.setGeometry(new_x, new_y, new_width, self.height())
        
        self.show()
        
        # --- 延長顯示時間邏輯 ---
        # 基礎 3 秒 + 每個字 0.3 秒 (例如 10 個字會顯示 6 秒)
        display_time = 1000 + (len(text) * 300)
        self.timer.start(display_time)

def move_with_dynamic_speed(start_x, start_y, target_x, target_y):
    """根據距離動態調整速度並移動滑鼠"""
    import math # 確保有匯入
    distance = math.sqrt((target_x - start_x)**2 + (target_y - start_y)**2)
    screen_w = pyautogui.size()[0]
    
    # 計算時間：0.4 ~ 1.2 秒
    duration = 0.4 + (distance / screen_w) * 0.8
    duration = min(1.2, max(0.4, duration))
    
    print(f"[移動中] 距離: {int(distance)}px, 耗時: {duration:.2f}s")
    
    # 執行移動
    pyautogui.moveTo(target_x, target_y, duration=duration, tween=pyautogui.easeInOutQuart)
    time.sleep(0.1) # 抵達後緩衝

def vision_click(target_description):
    """專門處理圖標、Logo、無文字按鈕"""
    print(f"[視覺搜尋] 正在鎖定圖標：{target_description}...")
    try:
        # 獲取初始位置
        start_x, start_y = pyautogui.position()
        
        # 獲取螢幕截圖與實際解析度
        screenshot = pyautogui.screenshot()
        screen_w, screen_h = pyautogui.size()
        
        # 將圖片轉為 Base64
        buffered = BytesIO()
        screenshot.save(buffered, format="PNG")
        img_str = base64.b64encode(buffered.getvalue()).decode('utf-8')
        
        prompt_text = (
            f"這是一個 {screen_w}x{screen_h} 的電腦螢幕截圖。"
            f"請精確定位「{target_description}」圖標的中心位置。"
            f"請以圖片左上角為 (0,0)，右下角為 (1000,1000) 比例計算。"
            f"請回傳 JSON 格式: {{\"x\": 0-1000, \"y\": 0-1000}}"
        )

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt_text},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{img_str}"}}
                ]
            }],
            response_format={ "type": "json_object" }
        )

        data = json.loads(response.choices[0].message.content)
        
        # 核心計算邏輯
        # 使用 float 確保精確度，再轉回 int
        target_x = int(float(data['x']) * screen_w / 1000)
        target_y = int(float(data['y']) * screen_h / 1000)
        
        print(f"DEBUG: AI 回傳原始值: x={data['x']}, y={data['y']}")
        print(f"DEBUG: 映射到螢幕 ({screen_w}x{screen_h}) 的座標是: ({target_x}, {target_y})")
        
        # 執行移動與點擊
        move_with_dynamic_speed(start_x, start_y, target_x, target_y)
        pyautogui.click()
        
        return f"Success: 已定位並點擊圖標「{target_description}」。"
    except Exception as e:
        return f"圖形辨識失敗：{e}"
    
def text_click(target_text):
    """專門處理帶有文字的按鈕或連結"""
    print(f"[文字辨識] 正在尋找關鍵字：{target_text}...")
    try:
        start_x, start_y = pyautogui.position()
        screenshot = pyautogui.screenshot()
        
        # 使用 EasyOCR 直接精確定位文字中心
        results = reader.readtext(np.array(screenshot))
        
        for (bbox, text, prob) in results:
            if target_text.lower() in text.lower():
                # 計算中心點
                (tl, tr, br, bl) = bbox
                target_x = int((tl[0] + br[0]) / 2)
                target_y = int((tl[1] + br[1]) / 2)
                
                print(f"DEBUG: 計算出的目標座標是 ({target_x}, {target_y})")
                print(f"[執行] 找到文字「{text}」，信心度：{prob:.2f}")
                move_with_dynamic_speed(start_x, start_y, target_x, target_y)
                pyautogui.click()
                return f"Success: 已找到並點擊文字「{text}」。"
        
        return f"文字辨識失敗：找不到包含「{target_text}」的按鈕。"
    except Exception as e:
        return f"文字處理出錯：{e}"
# 隱藏警告
warnings.filterwarnings("ignore", category=UserWarning, module="torch")

# --- 1. 初始化與全域設定 ---
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
PICO_KEY = os.getenv("PICOVOICE_ACCESS_KEY")
OWM_API_KEY = os.getenv("OPENWEATHER_API_KEY") 
PPN_PATH = "Raphael.ppn"

reader = easyocr.Reader(['ch_tra', 'en'], gpu=True)

def select_microphone():
    devices = PvRecorder.get_available_devices()
    print("\n--- 可用的麥克風裝置列表 ---")
    for i, device in enumerate(devices):
        print(f"[{i}] {device}")
    try:
        choice = input(f"\n請輸入麥克風 ID (預設使用 ID 0): ").strip()
        return int(choice) if choice else 0
    except: return 0
# DEVICE_INDEX = 2
DEVICE_INDEX = select_microphone()
try:
    pygame.mixer.pre_init(44100, -16, 2, 512) # 預設採樣率，縮小緩衝區
    pygame.mixer.init()
except Exception as e:
    print(f"[警告] Pygame Mixer 初始化失敗，嘗試相容模式: {e}")
    os.environ['SDL_AUDIODRIVER'] = 'dsound' # 強制切換至 DirectSound 驅動
    pygame.mixer.init()

# --- 2. 工具函數定義 (保留所有原有功能) ---

def play_sound(filename):
    try:
        if os.path.exists(filename):
            pygame.mixer.Sound(filename).play()
        elif filename == "notify.mp3": print("\a")
    except: pass

def get_current_time():
    return f"現在時間是 {datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}"

def get_weather(city="Chiayi"):
    # 關鍵修正：如果 city 是 None 或空字串，強制設為預設值
    if not city:
        city = "嘉義"
    try:
        # 現在這裡絕對安全，因為 city 一定是字串
        clean_city = city.replace("市", "").replace("縣", "")
        city_map = {
            "台北": "Taipei", "新北": "New Taipei", "桃園": "Taoyuan", 
            "台中": "Taichung", "台南": "Tainan", "高雄": "Kaohsiung", 
            "嘉義": "Chiayi", "雲林": "Yunlin"
        }
        search_city = city_map.get(clean_city, clean_city)
        url = f"http://api.openweathermap.org/data/2.5/weather?q={search_city}&appid={OWM_API_KEY}&units=metric&lang=zh_tw"
        res = requests.get(url, timeout=5).json()
        if res.get("cod") == 200:
            return f"{city}目前{res['weather'][0]['description']}，氣溫 {res['main']['temp']} 度。"
        return f"找不到 {city} 的天氣資訊（錯誤碼：{res.get('cod')}）。"
    except Exception as e:
        print(f"DEBUG 天氣錯誤: {e}")
        return "天氣連線失敗，請檢查網路或 API KEY。"

def control_computer(action):
    if action == "shutdown": os.system("shutdown /s /t 15"); return "十秒後關機。"
    if action == "restart": os.system("shutdown /r /t 15"); return "十秒後重啟。"
    if action == "cancel": os.system("shutdown /a"); return "已為您取消關機或重啟排程。"
    if action == "sleep": os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0"); return "進入睡眠。"
    return "無效操作。"

def open_software(app_name):
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    target_name = app_name.replace("幫我開啟", "").replace("打開", "").strip().lower()
    
    try:
        items = os.listdir(desktop_path)
        for item in items:
            name_without_ext = os.path.splitext(item)[0].lower()
            
            if target_name in name_without_ext and item.lower().endswith((".lnk", ".url")):
                full_path = os.path.normpath(os.path.join(desktop_path, item)) # 確保路徑格式正確
                
                # 改用這行：強制透過 Windows Shell 開啟
                subprocess.Popen(f'start "" "{full_path}"', shell=True)
                
                return f"已成功下達開啟指令：{item}"
        
        return "桌面找不到捷徑。"
    except Exception as e:
        return f"失敗：{e}"

def web_search(query):
    webbrowser.open(f"https://www.google.com/search?q={query}")
    return f"已搜尋 {query}。"

def get_system_status():
    return f"CPU: {psutil.cpu_percent()}%，記憶體: {psutil.virtual_memory().percent}%。"

async def set_timer(seconds, label="計時器"):
    await asyncio.sleep(seconds)
    return "" 

def click_text_on_screen(target_text):
    """執行單次搜尋與點擊，回傳是否成功找到目標"""
    try:
        screenshot = pyautogui.screenshot()
        img_gray = np.array(screenshot.convert('L'))
        results = reader.readtext(img_gray)
        
        # 轉小寫進行模糊比對
        search_target = target_text.lower()
        for (bbox, text, prob) in results:
            if search_target in text.lower():
                # 點擊該文字塊的右側中心
                rx = int(bbox[1][0])
                ry = int((bbox[1][1] + bbox[2][1]) / 2)
                pyautogui.click(rx, ry)
                return True # 找到並點擊了
        return False # 畫面上已無目標
    except Exception as e:
        print(f"視覺辨識異常: {e}")
        return False

def screen_assistant(user_intent):
    """
    整合型螢幕助手：具備 9 大功能分支、語言感知、精確物理退格。
    新增邏輯：
    1. 關鍵字偵測：只有說到「所有/全部」才會執行多點任務。
    2. 任務點計數：預設為 1，防止 generate 動作導致的無限循環。
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
            response_format={ "type": "json_object" }
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
    
    # 判斷使用者是否要求「全部/所有」
    is_all_task = any(word in user_intent for word in ["所有", "全部", "所有內容", "all", "every", "每個"])
    
    if is_all_task:
        # 讓 AI 掃描畫面並計算目標點數量
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
        # 預設為 1，即使畫面上有 testing=testing，也只處理一個
        task_count = 1
        print("[系統] 單次執行模式，預設處理 1 個匹配目標。")

    # --- 第三階段：精確循環執行 ---
    for i in range(task_count):
        print(f"[進度] 執行中：{i+1}/{task_count}...")
        
        # 每一輪重新截圖確保位移後的座標依然正確
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
                messages=[{"role": "system", "content": "妳是精密自動化助理，負責計算內容並回傳 JSON。"},
                          {"role": "user", "content": decision_prompt}],
                response_format={ "type": "json_object" }
            )
            data = json.loads(response.choices[0].message.content)
            
            idx = data.get("line_idx")
            # 移除可能誤加的括號
            final_content = data.get("content", "").strip().lstrip('[').rstrip(']')

            if idx is not None and idx < len(curr_filtered):
                bbox = curr_filtered[idx][0]
                original_text = curr_filtered[idx][1]
                real_len = len(original_text) # 計算物理退格次數

                # 定位到末端
                target_x = int(bbox[1][0])
                target_y = int((bbox[1][1] + bbox[2][1]) / 2)
                
                # 移動與執行
                pyautogui.click(target_x, target_y)
                time.sleep(0.3)

                should_delete = t_act in ['delete', 'optimize', 'translate', 'debug', 'replace'] or '翻譯' in user_intent

                if should_delete:
                    print(f"[執行] 精確物理退格：偵測到「{original_text}」，共 {real_len} 個字元")
                    for _ in range(real_len):
                        pyautogui.press('backspace')
                        # 稍微給一點延遲，讓系統反應，防止刪太快漏字
                        # time.sleep(0.01) 
                
                # 只有非純「生成」任務才刪除原文字 (生成不刪除，翻譯/更改要刪除)
                if t_act in ['delete', 'optimize', 'translate', 'debug', 'replace']:
                    print(f"[執行] 退格刪除 {real_len} 個字元...")
                    for _ in range(real_len):
                        pyautogui.press('backspace')
                
                # 貼上新內容
                if final_content:
                    pyperclip.copy(final_content)
                    pyautogui.hotkey('ctrl', 'v')
                
                print(f"[完成] ID {idx}：{data.get('reason')}")
                time.sleep(1.2) # 介面穩定緩衝
            else:
                print("[跳過] 無法定位匹配的 ID。")
                break
        except Exception as e:
            print(f"[錯誤] 循環執行中斷：{e}")
            break

    return f"拉菲爾：已完成「{t_act}」任務，處理點共計 {task_count} 處。"

def music_controller(args):
    """
    處理點歌與多媒體控制邏輯
    args: 直接接收來自 AI 的工具參數字典
    """
    # 從傳入的字典直接抓取欄位
    action = args.get('action')
    target = args.get('song_name')
    platform = args.get('platform', 'youtube')
    lang = args.get('language', 'auto')
    
    if action == 'play' and target:
        import webbrowser
        import urllib.parse
        
        query = urllib.parse.quote(f"{target} {lang if lang != 'auto' else ''}")
        
        if platform == 'youtube':
            url = f"https://www.youtube.com/results?search_query={query}"
        else:
            # 也可以擴充其他平台的 URL
            url = f"https://www.google.com/search?q={target}+music"
            
        print(f"[音樂助手] 語系:{lang} | 平台:{platform} | 搜尋：{target}")
        webbrowser.open(url)
        return f"好的，正在為您在 {platform} 搜尋並播放 {target}。"
    
    elif action == 'next':
        pyautogui.press('medianexttrack')
        return "切換至下一首。"
    
    # ... 其他動作 (stop, volume_up 等) ...
    return "已執行音樂控制指令。"

def toggle_subtitles(action):
    """控制字幕顯示開關"""
    if action == "on":
        subtitle_win.is_enabled = True
        return "好的，已為您開啟字幕顯示。"
    elif action == "off":
        subtitle_win.is_enabled = False
        subtitle_win.hide() # 立即隱藏目前顯示中的字幕
        return "沒問題，我會把字幕隱藏起來。"
    return "無效的字幕指令。"

def shutdown_raphael():
    print("[系統] 準備執行自我關閉...")
    return "好的，系統即將關閉。期待下次與您見面，祝您生活愉快，再見。"
    
# --- 4. 工具描述 ---
tools = [
    {"type": "function", "function": {"name": "get_system_status", "description": "問系統狀況"}},
    {"type": "function", "function": {"name": "get_current_time", "description": "問時間"}},
    {"type": "function", "function": {"name": "get_weather", "description": "查天氣", "parameters": {"type": "object", "properties": {"city": {"type": "string"}}}}},
    {"type": "function", "function": {"name": "control_computer", "description": "電腦控制", "parameters": {"type": "object", "properties": {"action": {"type": "string"}}}}},
    {"type": "function", "function": {"name": "open_software", "description": "開軟體", "parameters": {"type": "object", "properties": {"app_name": {"type": "string"}}}}},
    {"type": "function", "function": {"name": "vision_click","description": "點擊『沒有文字』的圖標、圖案或顏色按鈕（如：紅色的圓按鈕、Chrome 圖示）。","parameters": {"type": "object",  "properties": { "target_description": {"type": "string","description": "圖標的描述"}},"required": ["target_description"]}}},
    {"type": "function", "function": {"name": "text_click","description": "點擊『帶有文字』的按鈕、選單或連結（如：『我的電腦』、網頁上的『新聞』）。","parameters": {"type": "object","properties": {"target_text": {"type": "string","description": "按鈕上的文字內容"}},"required": ["target_text"]}}},
    {"type": "function", "function": {"name": "set_timer", "description": "當使用者要求在一段時間後提醒時必須呼叫此工具。禁止直接用對話回覆時間到了。須抓取seconds(秒數)與label(標籤)。", "parameters": {"type": "object", "properties": {"seconds": {"type": "integer", "description": "延遲總秒數"}, "label": {"type": "string", "description": "提醒內容"}}, "required": ["seconds", "label"]}}},
    {"type": "function", "function": {"name": "web_search", "description": "當使用者要求搜尋資訊、查詢網路、或是詢問拉菲爾不知道的最新消息時使用。", "parameters": {"type": "object", "properties": {"query": {"type": "string", "description": "搜尋關鍵字"}}, "required": ["query"]}}},
    {"type": "function","function": {"name": "screen_assistant","description": "全能畫面助手。處理畫面上所有文字或程式碼的生成、刪除、優化、翻譯、除錯。AI 會自動判斷對象類型與執行多次檢查。","parameters": {"type": "object","properties": {"user_intent": {"type": "string", "description": "使用者的原始需求內容"}},"required": ["user_intent"]}}},
    {"type": "function","function": {"name": "music_controller","description": "點歌與多媒體控制助手。負責搜尋歌曲、播放音樂、切換曲目及調整音量。支援 YouTube Music、Spotify 或系統媒體控制。","parameters": {"type": "object","properties": {"action": {"type": "string","enum": ["play", "stop", "next", "previous", "volume_up", "volume_down"],"description": "要執行的音樂操作動作，例如播放、停止、下一首。"},"song_name": {"type": "string","description": "想要播放的歌名或歌手關鍵字，例如：'周杰倫 告白氣球'。"},"platform": {"type": "string","enum": ["youtube", "spotify", "system"],"description": "指定的播放平台，預設為 youtube。"},"language": {"type": "string","description": "歌曲的語系，用於精準搜尋。例如：'日文', '英文'。"}},"required": ["action"]}}},
    {"type": "function","function": {"name": "toggle_subtitles","description": "開啟或關閉桌面字幕顯示。","parameters": {"type": "object","properties": {"action": {"type": "string", "enum": ["on", "off"], "description": "on 代表開啟，off 代表關閉"}},"required": ["action"]}}},
    {"type": "function","function": {"name": "shutdown_raphael","description": "當使用者要求關閉程式、退出、關閉拉菲爾、或說再見並要求關閉時使用。"}},
]

chat_history = [{"role": "system", "content": "你是拉菲爾。當使用者要求對畫面代碼除錯、優化或生成時，請使用 code_assistant。"}]

sub_signal = SubtitleSignal()

async def speak(text):
    print(f"拉菲爾：{text}")
    sub_signal.text_updated.emit(text)
    try:
        pygame.mixer.music.unload()
        await edge_tts.Communicate(text, "zh-TW-YunJheNeural").save("res.mp3")
        pygame.mixer.music.load("res.mp3")
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy(): await asyncio.sleep(0.1)
    except: pass

async def get_voice_input(recorder, timeout_seconds=5):
    """
    修正版：使用 await 呼叫 speak，避免 event loop 衝突
    """
    print(f"\n[拉菲爾聆聽中...] (緩衝時間: {timeout_seconds}秒)")
    
    start_time = time.time()
    audio_data = []
    
    while (time.time() - start_time) < timeout_seconds:
        frame = recorder.read()
        audio_data.extend(frame)
        
        if len(audio_data) >= 1600:
            rms = math.sqrt(sum(s**2 for s in audio_data[-1600:]) / 1600)
            
            if rms > 450:
                print("[偵測到指令，處理中...]")
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
    
    # 超時未聽到聲音
    print("[逾時] 未偵測到指令。")
    play_sound("Stop_listen_notify.mp3")
    
    # 修正處：直接 await，不要用 asyncio.run
    await speak("未聽到指令。") 
    return ""

def run_async_main():
    """在獨立線程中運行原本的監聽邏輯"""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(main())

async def main():
    porcupine = pvporcupine.create(access_key=PICO_KEY, keyword_paths=[PPN_PATH], sensitivities=[0.95])
    recorder = PvRecorder(device_index=DEVICE_INDEX, frame_length=porcupine.frame_length)
    print(f"\n--- 拉菲爾系統啟動 (包含聆聽提示詞) ---")
    
    recorder.start()
    try:
        while True:
            if porcupine.process(recorder.read()) >= 0:
                play_sound("notify.mp3")
                user_text = await get_voice_input(recorder)
                if not user_text: continue 
                print(f"你說：{user_text}")
                play_sound("reply_notify.mp3") 
                print("[處理中]")
                chat_history.append({"role": "user", "content": user_text})
                comp = client.chat.completions.create(model="gpt-4o-mini", messages=chat_history, tools=tools)
                msg = comp.choices[0].message
                if msg.tool_calls:
                    chat_history.append(msg)

                    for tc in msg.tool_calls:
                        fn, args = tc.function.name, json.loads(tc.function.arguments)
                        res = ""

                        if fn == "set_timer":
                            u_sec, u_lab = args.get("seconds"), args.get("label")
                            asyncio.create_task(set_timer(u_sec, u_lab))
                            chat_history.append({"role": "tool", "tool_call_id": tc.id, "name": fn, "content": "timer_set"})
                            ans = f"沒問題，我已經設定好在 {u_sec} 秒後提醒您「{u_lab}」。"
                            await speak(ans)
                            break
                        
                        if fn == "get_current_time": res = get_current_time()
                        elif fn == "music_controller": res = music_controller(args)
                        elif fn == "web_search": res = web_search(args.get("query"))
                        elif fn == "get_weather": res = get_weather(args.get("city"))
                        elif fn == "control_computer": res = control_computer(args.get("action"))
                        elif fn == "screen_assistant": res = screen_assistant(user_intent=args.get("user_intent"))
                        elif fn == "open_software": res = open_software(args.get("app_name"))
                        elif fn == "vision_click": res = vision_click(args.get("target_description"))
                        elif fn == "text_click": res = text_click(args.get("target_text"))
                        elif fn == "get_system_status": res = get_system_status()
                        elif fn == "toggle_subtitles": res = toggle_subtitles(args.get("action"))
                        elif fn == "shutdown_raphael":
                            res = shutdown_raphael()
                            await speak(res) 
                            if pygame.mixer.get_init():
                                while pygame.mixer.music.get_busy():
                                    await asyncio.sleep(0.1)
                            print("[系統] 語音播放結束，清理資源中...")
                            pygame.mixer.music.stop()
                            pygame.mixer.quit()
                            await asyncio.sleep(0.5)
                            os._exit(0)

                        chat_history.append({"role": "tool", "tool_call_id": tc.id, "name": fn, "content": res})
                    
                    final_comp = client.chat.completions.create(model="gpt-4o-mini", messages=chat_history)
                    ans = final_comp.choices[0].message.content
                else: ans = msg.content 
                chat_history.append({"role": "assistant", "content": ans})
                await speak(ans)
    finally:
        recorder.stop(); recorder.delete(); porcupine.delete()

if __name__ == "__main__":
    # asyncio.run(main())
    try:
        pygame.mixer.init()
    except:
        pass

    app = QApplication(sys.argv)
    
    subtitle_win = RaphaelSubtitleWindow()
    sub_signal.text_updated.connect(subtitle_win.display_text)
    logic_thread = threading.Thread(target=run_async_main, daemon=True)
    logic_thread.start()
    sys.exit(app.exec_())