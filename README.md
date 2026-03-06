# Raphael: Multimodal AI Desktop Copilot
# 專案簡介 (Project Overview)
Raphael (拉菲爾) 是一款次世代的智慧桌面助理，旨在重塑人機協作體驗。透過整合語音感官、視覺辨識與大型語言模型 (LLM)，Raphael 能像真人助手一樣「看見」你的螢幕、「聽懂」你的需求，並直接在你的電腦上執行複雜任務。

Raphael is a next-generation AI desktop assistant designed to reshape the human-computer interaction experience. By integrating voice sensing, computer vision, and LLMs, Raphael can "see" your screen and "hear" your needs just like a human assistant.

# 核心功能 (Key Features)
語音交互 (Voice Interaction)：支援自定義喚醒詞（Raphael）與自然語言指令理解。

視覺感知 (Vision Sensing)：利用 EasyOCR 與 GPT-4o 實現螢幕內容精確定位與點擊。

自動化工作流 (Automation)：自動進行代碼除錯、翻譯、文字優化及系統控制（如音樂、天氣、關機）。

邊緣運算優化 (Edge AI Optimization)：採用輕量化邊緣運算技術進行喚醒偵測，降低系統延遲。

科技感字幕 (Dynamic Subtitles)：即時顯示 AI 回饋文字，提供極佳的互動視覺體驗。

# 技術架構 (Technical Architecture)
本專案採用多模態協作架構：

Wake-word Detection: Picovoice Porcupine (Local Edge AI)

Speech-to-Text: OpenAI Whisper API

Reasoning Engine: OpenAI GPT-4o / GPT-4o-mini

Computer Vision: EasyOCR + OpenCV

Action Execution: PyAutoGUI & Subprocess

GUI Framework: PyQt5

# 快速上手 (Quick Start)
1. 環境需求 (Prerequisites)
Python 3.9+

OpenAI API Key

Picovoice Access Key

OpenWeatherMap API Key

2. 安裝步驟 (Installation)
Bash

# 克隆專案
git clone https://github.com/raphael4287/Raphael-Multimodal-Desktop-Copilot.git

# 安裝依賴
pip install -r requirements.txt
3. 設定環境變數 (Environment Variables)
請建立 .env 檔案並填入以下資訊：

Plaintext

OPENAI_API_KEY=你的金鑰
PICOVOICE_ACCESS_KEY=你的金鑰
OPENWEATHER_API_KEY=你的金鑰
AI 輕量化說明 (AI Lightweight Implementation)
本專案在設計上特別考量了效能平衡：

本地端喚醒：使用 Picovoice 在本地端進行 24/7 監聽，無需上傳音訊至雲端，極大節省頻寬與運算資源。

混合視覺處裡：優先使用本地 EasyOCR 進行文字定位，僅在複雜圖形辨識時調用 GPT-4o，優化 API 調用效率。
