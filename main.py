import os
import pandas as pd
import requests
import base64
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageSendMessage
from PIL import Image, ImageDraw, ImageFont
import uuid
import textwrap
import re
from io import StringIO

app = Flask(__name__)

# --- 設定區 ---
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')
SHEET_URL = os.environ.get('SHEET_URL')
IMGBB_API_KEY = "f65fa2212137d99c892644b1be26afac" 

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(BASE_DIR, "myfont.ttf")

@app.route("/", methods=['GET'])
def index():
    return "機器人服務中 - 亂碼與寬度修正版"

def upload_to_imgbb(image_path):
    try:
        with open(image_path, "rb") as file:
            url = "https://api.imgbb.com/1/upload"
            payload = {"key": IMGBB_API_KEY, "image": base64.b64encode(file.read())}
            res = requests.post(url, data=payload, timeout=20)
            return res.json()['data']['url'] if res.status_code == 200 else "Err"
    except: return "UploadFail"

def create_table_image_pil(df):
    # --- 修正：加大總寬度計算，確保右側不超出 ---
    col_widths = [60, 150, 220, 130, 200, 400, 500] # 最後一欄加寬到 500
    line_height, padding = 45, 25
    rows_data = []
    
    # 標題
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1))
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        # 對應欄位的每行字數限制 (根據新寬度調整)
        char_counts = [4, 12, 12, 10, 12, 18, 22] 
        for i in range(min(7, len(row))):
            val = row.iloc[i]
            
            # 處理排序欄位去除 .0
            if i == 0:
                try: text = str(int(float(val)))
                except: text = str(val).replace("nan", "").strip()
            else:
                text = str(val).replace("nan", "").strip()
            
            # --- 核心修正：強制移除所有非列印字元與常見亂碼字元 ---
            # 1. 移除 ASCII 控制字元 (0-31)
            text = "".join(ch for ch in text if ord(ch) >= 32 or ch == '\n')
            # 2. 針對截圖中出現的方框符號 (常見於 Google Sheet 換行符號)
            replacements = {
                '\u3000': ' ', '\xa0': ' ', '\r': '', '\t': ' ',
                '\u200b': '', '\u200c': '', '\u200d': '', '\ufeff': '',
                '\\n': '\n' # 確保文字內的換行符號能正確運作
            }
            for old, new in replacements.items():
                text = text.replace(old, new)
            
            # 使用更穩定的斷行演算法
            lines = []
            for part in text.split('\n'):
                lines.extend(textwrap.wrap(part, width=char_counts[i]))
            
            wrapped_row.append("\n".join(lines) if lines else "")
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    # 寬度重新計算：總和 + 左右 padding
    canvas_width = sum(col_widths) + (padding * 2) + 20 
    total_h = sum([m * line_height + 25 for _, m in rows_data]) + (padding * 2)
    
    image = Image.new('RGB', (int(canvas_width), int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    try:
        font = ImageFont.truetype(FONT_PATH, 24) if os.path.exists(FONT_PATH) else ImageFont.load_default()
        h_font = ImageFont.truetype(FONT_PATH, 26) if os.path.exists(FONT_PATH) else ImageFont.load_default()
    except:
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 25
        
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        
        # 畫列背景
        draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg)
        
        for c_idx, text in enumerate(text_list):
            # 畫格子框
            draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            # 填入文字
            draw.text((x + 12, y + 10), text, fill=tc, font=font if r_idx > 0 else h_font, spacing=8)
            x += col_widths[c_idx]
        y += row_h

    temp_file = f"{uuid.uuid4()}.png"
    image.save(temp_file, "PNG", optimize=True)
    return temp_file

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try: handler.handle(body, signature)
    except InvalidSignatureError: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    msg = event.message.text.strip()
    if msg == "總表":
        try:
            res = requests.get(SHEET_URL, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
            res.encoding = 'utf-8-sig'
            df = pd.read_csv(StringIO(res.text), on_bad_lines='skip', header=0)
            
            df = df[df.iloc[:, 0].notna()]
            df = df[df.iloc[:, 0].astype(str).str.strip() != ""]
            df = df[~df.iloc[:, 0].astype(str).str.lower().isin(["nan", "none", "0", "0.0"])]

            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="目前試算表內無有效報修資料。"))
                return

            if len(df) > 20: df = df.head(20)

            img_path = create_table_image_pil(df)
            img_url = upload_to_imgbb(img_path)
            
            if img_url and img_url.startswith("http"):
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="圖片服務暫時繁忙，請稍後。"))
            
            if os.path.exists(img_path): os.remove(img_path)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"讀取失敗，請確認網路連線。"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="連線正常！請輸入「總表」。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
