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
    return "機器人服務中 - 字元相容與排版優化版"

def upload_to_imgbb(image_path):
    try:
        with open(image_path, "rb") as file:
            url = "https://api.imgbb.com/1/upload"
            payload = {"key": IMGBB_API_KEY, "image": base64.b64encode(file.read())}
            res = requests.post(url, data=payload, timeout=20)
            return res.json()['data']['url'] if res.status_code == 200 else "Err"
    except: return "UploadFail"

def clean_text(text):
    """強力清洗函數：移除所有導致亂碼的特殊字元"""
    if not text: return ""
    # 1. 替換掉常見的特殊換行與空格
    text = text.replace('\u3000', ' ').replace('\xa0', ' ').replace('\r', '').replace('\t', ' ')
    # 2. 核心修正：使用正規表達式只保留 中文、英文、數字與基本標點，其餘特殊符號轉為空格
    # 這能有效防止截圖中出現的方框亂碼
    text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9\s,.，。？！、：()（）*#\-+]', ' ', text)
    return text.strip()

def create_table_image_pil(df):
    # 寬度配置：最後一欄加大到 550 避免地址與描述被切掉
    col_widths = [60, 150, 220, 130, 200, 400, 550] 
    line_height, padding = 42, 25 # 稍微調低行高
    rows_data = []
    
    # 標題
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1))
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        # 字數限制調優：根據新寬度微調
        char_counts = [4, 12, 12, 10, 12, 18, 25] 
        for i in range(min(7, len(row))):
            val = row.iloc[i]
            
            # A 欄排序處理
            if i == 0:
                try: text = str(int(float(val)))
                except: text = str(val).replace("nan", "").strip()
            else:
                text = clean_text(str(val).replace("nan", "").strip())
            
            # 使用更彈性的斷行處理
            lines = []
            for part in text.split('\n'):
                if part.strip():
                    lines.extend(textwrap.wrap(part, width=char_counts[i]))
            
            final_cell_text = "\n".join(lines) if lines else ""
            wrapped_row.append(final_cell_text)
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    # 畫布總寬度計算
    canvas_width = sum(col_widths) + (padding * 2) + 30 
    total_h = sum([m * line_height + 25 for _, m in rows_data]) + (padding * 2)
    
    image = Image.new('RGB', (int(canvas_width), int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    try:
        # 字體微調：從 24 降到 22，增加在格內的舒適度
        font = ImageFont.truetype(FONT_PATH, 22) if os.path.exists(FONT_PATH) else ImageFont.load_default()
        h_font = ImageFont.truetype(FONT_PATH, 24) if os.path.exists(FONT_PATH) else ImageFont.load_default()
    except:
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 25
        
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        
        draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg)
        for c_idx, text in enumerate(text_list):
            draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            # 文字垂直居中微調
            draw.text((x + 12, y + 12), text, fill=tc, font=font if r_idx > 0 else h_font, spacing=6)
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
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="報修清單目前是空的。"))
                return

            if len(df) > 20: df = df.head(20)

            img_path = create_table_image_pil(df)
            img_url = upload_to_imgbb(img_path)
            
            if img_url and img_url.startswith("http"):
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="圖床連線不穩定，請稍後再試。"))
            
            if os.path.exists(img_path): os.remove(img_path)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"讀取失敗，請確認網路連線。"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="請輸入「總表」來產出報表。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
