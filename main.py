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

def create_table_image_pil(df):
    # --- 關鍵修正：大幅度增加欄位寬度，確保描述不被切掉 ---
    # 排序, 日期, 店別, 型號, 電話, 地址, 問題描述
    col_widths = [70, 160, 240, 150, 220, 500, 650] 
    line_height, padding = 45, 50 
    rows_data = []
    
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1))
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        # 設定各欄位每行字數限制，配合微軟正黑體的字體大小
        char_counts = [5, 12, 12, 10, 15, 23, 32] 
        for i in range(min(7, len(row))):
            val = row.iloc[i]
            text = str(val).replace("nan", "").strip() if i != 0 else str(int(float(val))) if pd.notna(val) else ""
            
            # 清理特殊換行符號
            text = text.replace('\r', '').replace('\u3000', ' ').replace('\xa0', ' ')
            
            lines = []
            for part in text.split('\n'):
                # 使用文字折疊功能
                lines.extend(textwrap.wrap(part, width=char_counts[i]))
            
            wrapped_row.append("\n".join(lines) if lines else "")
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    # 計算總畫布寬度 (包含緩衝)
    total_table_width = sum(col_widths)
    canvas_width = total_table_width + (padding * 2) + 60
    total_h = sum([m * line_height + 30 for _, m in rows_data]) + (padding * 2)
    
    image = Image.new('RGB', (int(canvas_width), int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    # 載入字型 (微軟正黑體建議設在 22-24 號字)
    try:
        font = ImageFont.truetype(FONT_PATH, 22) if os.path.exists(FONT_PATH) else ImageFont.load_default()
        h_font = ImageFont.truetype(FONT_PATH, 24) if os.path.exists(FONT_PATH) else ImageFont.load_default()
    except:
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 30
        
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        
        draw.rectangle([x, y, x + total_table_width, y + row_h], fill=bg)
        
        curr_x = x
        for c_idx, text in enumerate(text_list):
            draw.rectangle([curr_x, y, curr_x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            # 文字渲染，增加微幅內距避免壓線
            draw.text((curr_x + 12, y + 12), text, fill=tc, font=font if r_idx > 0 else h_font, spacing=8)
            curr_x += col_widths[c_idx]
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
            df = df[df.iloc[:, 0].notna()].head(20)

            img_path = create_table_image_pil(df)
            img_url = upload_to_imgbb(img_path)
            
            if img_url and img_url.startswith("http"):
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            
            if os.path.exists(img_path): os.remove(img_path)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="讀取失敗，請稍後。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
