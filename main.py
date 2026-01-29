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
import time

app = Flask(__name__)

# --- 設定區 (從環境變數讀取，確保安全) ---
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')
# 使用穩定發佈連結 (這串網址結尾必須是 output=csv)
CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQ2RjCpIscmC1QHOkO2MsWtkpzhQ4ppLUy5xsOSWRsiaFV1zXjQiRwrF7_QUuyMaO2Dt9bwBQJOJgUt/pub?output=csv"
IMGBB_API_KEY = "f65fa2212137d99c892644b1be26afac" 
FONT_PATH = "fonts/LINESeedJP-Regular.ttf"

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

@app.route("/", methods=['GET'])
def index():
    return "<h1>總表機器人運行中</h1><p>請將 LINE Webhook 設為此網址後方加上 /callback</p>"

def upload_to_imgbb(image_path):
    try:
        with open(image_path, "rb") as file:
            url = "https://api.imgbb.com/1/upload"
            payload = {"key": IMGBB_API_KEY, "image": base64.b64encode(file.read())}
            res = requests.post(url, payload)
            if res.status_code == 200:
                return res.json()['data']['url']
    except Exception as e:
        print(f"上傳失敗: {e}")
    return None

def create_table_image_pil(df):
    # 【優化：加寬防止重疊】地址給 600, 電話給 300
    col_widths = [80, 180, 240, 140, 300, 600, 620] 
    line_height, padding = 45, 30
    rows_data = []
    # 標題行
    rows_data.append(["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"])
    
    for _, row in df.iterrows():
        wrapped_row, max_lines = [], 1
        # 設定各欄位的換行字數限制
        char_counts = [4, 10, 10, 8, 18, 24, 26] 
        for i in range(min(7, len(row))):
            val = row.iloc[i] if i < len(row) else ""
            text = str(val) if pd.notna(val) else ""
            # 移除不可見亂碼
            text = "".join(c for c in text if c.isprintable()) 
            lines = textwrap.wrap(text, width=char_counts[i])
            wrapped_row.append("\n".join(lines) if lines else "")
            max_lines = max(max_lines, len(lines))
        
        # 判定 A 欄 (排序) 是否為空，用於底色判定
        sort_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        rows_data.append((wrapped_row, max_lines, sort_val))

    # 計算總高度與寬度
    total_h = 80 + sum([m * line_height + 40 for _, m, _ in rows_data[1:]]) + (2 * padding)
    total_w = sum(col_widths) + (2 * padding)
    image = Image.new('RGB', (total_w, int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    try:
        font = ImageFont.truetype(FONT_PATH, 28)
        h_font = ImageFont.truetype(FONT_PATH, 30)
    except:
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, row_item in enumerate(rows_data):
        x = padding
        if r_idx == 0:
            # 繪製標題行
            row_h = 80
            draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=(45, 90, 45))
            for c_idx, text in enumerate(row_item):
                draw.text((x + 15, y + 25), text, fill=(255, 255, 255), font=h_font)
                x += col_widths[c_idx]
        else:
            # 繪製內容行
            text_list, m_lines, sort_val = row_item
            row_h = m_lines * line_height + 40
            
            # 【優化：A欄空白上色】若 A 欄為空則顯示淺綠底
            is_empty_a = not sort_val or sort_val.lower() == "nan" or sort_val == ""
            bg_color = (235, 245, 235) if is_empty_a else (255, 255, 255)
            
            draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg_color)
            for c_idx, text in enumerate(text_list):
                draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(200, 200, 200), width=1)
                draw.text((x + 15, y + 20), text, fill=(0, 0, 0), font=font, spacing=8)
                x += col_widths[c_idx]
        y += row_h

    temp_file = f"temp_{uuid.uuid4()}.png"
    image.save(temp_file)
    return temp_file

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    if event.message.text == "總表":
        try:
            # 抓取 CSV 資料，加上時間戳防止快取
            df_raw = pd.read_csv(f"{CSV_URL}&t={int(time.time())}", encoding='utf-8-sig', header=None) 
            
            # 自動搜尋含有「排序」的那一行作為標題
            header_idx = 0
            for i, row in df_raw.iterrows():
                if "排序" in str(row.values):
                    header_idx = i
                    break
            
            df = df_raw.iloc[header_idx+1:].copy()
            df.columns = df_raw.iloc[header_idx]
            # 過濾空白行
            df = df.dropna(subset=[df.columns[1]], how='all').reset_index(drop=True)
            
            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="試算表內目前無資料。"))
                return

            local_path = create_table_image_pil(df)
            public_url = upload_to_imgbb(local_path)
            
            if public_url:
                line_bot_api.reply_message(
                    event.reply_token,
                    ImageSendMessage(original_content_url=public_url, preview_image_url=public_url)
                )
            if os.path.exists(local_path): os.remove(local_path)
        except Exception as e:
            print(f"詳細錯誤: {e}")
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"生成失敗，請確認試算表發佈連結正確。"))

if __name__ == "__main__":
    # Koyeb / Render 等平台會自動給 Port，這裡設定預設為 5000
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
