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
    return "機器人服務中 - 編碼優化版"

def upload_to_imgbb(image_path):
    try:
        with open(image_path, "rb") as file:
            url = "https://api.imgbb.com/1/upload"
            payload = {"key": IMGBB_API_KEY, "image": base64.b64encode(file.read())}
            res = requests.post(url, data=payload, timeout=20)
            return res.json()['data']['url'] if res.status_code == 200 else "Err"
    except: return "UploadFail"

def create_table_image_pil(df):
    col_widths = [60, 130, 180, 130, 200, 400, 400] 
    line_height, padding = 45, 25
    rows_data = []
    
    # 標題
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1))
    
    # 強制清洗資料
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        char_counts = [4, 10, 8, 10, 12, 18, 18]
        for i in range(min(7, len(row))):
            # 取得原始內容並過濾 nan
            text = str(row.iloc[i]).replace("nan", "").strip()
            
            # 針對剩餘的小方塊符號進行簡單替換 (備援邏輯)
            text = text.replace('\u3000', ' ').replace('\xa0', ' ')
            
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
        
        # 僅標題列使用深綠色，其餘純白，徹底解決底部黃塊問題
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        
        draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg)
        for c_idx, text in enumerate(text_list):
            draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            draw.text((x + 10, y + 10), text, fill=tc, font=font if r_idx > 0 else h_font, spacing=8)
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
            # 加入 headers 模擬瀏覽器，增加連線穩定度
            res = requests.get(SHEET_URL, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
            res.encoding = 'utf-8-sig'
            
            df = pd.read_csv(StringIO(res.text), on_bad_lines='skip', header=0)
            
            # --- 嚴格 A 欄過濾邏輯 ---
            df = df[df.iloc[:, 0].notna()]
            df = df[df.iloc[:, 0].astype(str).str.strip() != ""]
            df = df[~df.iloc[:, 0].astype(str).str.lower().isin(["nan", "none", "0", "0.0"])]

            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="目前試算表內無有效報修資料。"))
                return

            # 最多顯示 20 筆
            if len(df) > 20: df = df.head(20)

            img_path = create_table_image_pil(df)
            img_url = upload_to_imgbb(img_path)
            
            if img_url and img_url.startswith("http"):
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="圖片上傳圖床失敗，請檢查網路。"))
            
            if os.path.exists(img_path): os.remove(img_path)
        except Exception as e:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"讀取資料異常，請稍後再試。"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="連線正常！輸入「總表」產生報表。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
