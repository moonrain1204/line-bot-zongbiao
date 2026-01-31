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
    return "機器人服務中 - 欄位優化版"

def upload_to_imgbb(image_path):
    try:
        with open(image_path, "rb") as file:
            url = "https://api.imgbb.com/1/upload"
            payload = {"key": IMGBB_API_KEY, "image": base64.b64encode(file.read())}
            res = requests.post(url, data=payload, timeout=20)
            return res.json()['data']['url'] if res.status_code == 200 else "Err"
    except: return "UploadFail"

def create_table_image_pil(df):
    # --- 修正 1：加大 col_widths 寬度，避免日期與店別重疊 ---
    # 原本 [60, 130, 180, ...] -> 改為 [60, 150, 220, ...]
    col_widths = [60, 150, 220, 130, 200, 400, 450] 
    line_height, padding = 45, 25
    rows_data = []
    
    # 標題
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1))
    
    # 強制清洗資料與處理排序顯示
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        # 對應欄位的每行字數限制
        char_counts = [4, 12, 12, 10, 12, 18, 20]
        for i in range(min(7, len(row))):
            val = row.iloc[i]
            
            # 處理 A 欄 (排序) 去除小數點
            if i == 0:
                try:
                    text = str(int(float(val)))
                except:
                    text = str(val).replace("nan", "").strip()
            else:
                text = str(val).replace("nan", "").strip()
            
            # --- 修正 2：更全面的亂碼與隱形字元排除 ---
            replacements = {
                '\u3000': ' ', '\xa0': ' ', '\r': '', '\t': ' ',
                '\u2611': '[v]', '\u2610': '[ ]', 
                '\u2715': 'x', '\u2716': 'x', 
                '\uf06c': '*', '\ufb01': 'fi',
                '\\n': ' ', # 排除字串化的換行符
            }
            for old, new in replacements.items():
                text = text.replace(old, new)
            
            # 使用 textwrap 處理斷行
            lines = textwrap.wrap(text, width=char_counts[i])
            wrapped_row.append("\n".join(lines) if lines else "")
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    total_h = sum([m * line_height + 25 for _, m in rows_data]) + (padding * 2)
    image = Image.new('RGB', (sum(col_widths) + padding * 2, int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    try:
        if os.path.exists(FONT_PATH):
            font = ImageFont.truetype(FONT_PATH, 24)
            h_font = ImageFont.truetype(FONT_PATH, 26)
        else:
            font = h_font = ImageFont.load_default()
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
            # 增加文字水平偏移 (x+12)，讓內容不緊貼邊線
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
            
            # 嚴格 A 欄過濾
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
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="圖片產生成功但圖床服務繁忙，請稍後。"))
            
            if os.path.exists(img_path): os.remove(img_path)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"系統連線繁忙，請重新輸入「總表」。"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="機器人在線中！請輸入「總表」。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
