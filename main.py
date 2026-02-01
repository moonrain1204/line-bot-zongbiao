import os
import pandas as pd
import requests
import base64
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageSendMessage
from PIL import Image, ImageDraw, ImageFont
import uuid
import textwrap
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import json

app = Flask(__name__)

# --- 設定區 ---
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')
IMGBB_API_KEY = "f65fa2212137d99c892644b1be26afac" 
SHEET_KEY = os.environ.get('SHEET_KEY')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

def get_sheet_client():
    """檢查環境變數並連線"""
    creds_raw = os.environ.get('GOOGLE_CREDS')
    if not creds_raw:
        return "ERR_NO_CREDS (找不到 GOOGLE_CREDS 變數)"
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_json = json.loads(creds_raw)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
        return gspread.authorize(creds)
    except Exception as e:
        return f"JSON_ERROR: {str(e)}"

def create_table_image_pil(df):
    """繪製 A 欄非空的總表圖片"""
    col_widths = [80, 160, 240, 160, 220, 580, 750] 
    line_height, padding = 55, 60 
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data = [(headers, 1)]
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        char_counts = [5, 12, 12, 10, 15, 22, 28] 
        for i in range(len(headers)):
            text = str(row.iloc[i]).replace("nan", "").strip()
            lines = textwrap.wrap(text, width=char_counts[i]) if text else [" "]
            wrapped_row.append("\n".join(lines))
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    total_table_width = sum(col_widths)
    total_h = sum([m * line_height + 45 for _, m in rows_data]) + (padding * 2)
    image = Image.new('RGB', (int(total_table_width + padding * 2), int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    # 字型設定
    font_path = os.path.join(os.getcwd(), "boldfonts.ttf")
    font = ImageFont.truetype(font_path, 28) if os.path.exists(font_path) else ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines) in enumerate(rows_data):
        row_h = m_lines * line_height + 45
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        draw.rectangle([padding, y, padding + total_table_width, y + row_h], fill=bg)
        curr_x = padding
        for c_idx, text in enumerate(text_list):
            draw.rectangle([curr_x, y, curr_x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            draw.text((curr_x + 15, y + 15), text, fill=tc, font=font, spacing=8)
            curr_x += col_widths[c_idx]
        y += row_h

    temp_file = f"{uuid.uuid4()}.png"
    image.save(temp_file, "PNG")
    return temp_file

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try: handler.handle(body, signature)
    except: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    msg = event.message.text.strip()
    client = get_sheet_client()
    
    if isinstance(client, str):
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"連線失敗：{client}"))
        return

    try:
        sh = client.open_by_key(SHEET_KEY)
        wks = sh.worksheet("永慶待修")

        if msg == "總表":
            data = wks.get_all_values()
            df = pd.DataFrame(data[1:], columns=data[0])
            
            # 【關鍵】過濾 A 欄非空的範圍
            display_df = df[df['排序'].str.strip() != ""].copy()
            
            if not display_df.empty:
                img_path = create_table_image_pil(display_df)
                with open(img_path, "rb") as f:
                    img_res = requests.post("https://api.imgbb.com/1/upload", 
                                            data={"key": IMGBB_API_KEY, "image": base64.b64encode(f.read())})
                    img_url = img_res.json()['data']['url']
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
                if os.path.exists(img_path): os.remove(img_path)
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="目前無待修資料(A欄皆為空)。"))

    except Exception as e:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"執行錯誤: {e}"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
