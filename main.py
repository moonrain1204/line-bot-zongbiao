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
FONT_PATH = "myfont.ttf" 

@app.route("/", methods=['GET'])
def index():
    return "機器人運行中"

def upload_to_imgbb(image_path):
    try:
        with open(image_path, "rb") as file:
            url = "https://api.imgbb.com/1/upload"
            payload = {"key": IMGBB_API_KEY, "image": base64.b64encode(file.read())}
            res = requests.post(url, data=payload, timeout=15)
            return res.json()['data']['url'] if res.status_code == 200 else f"Err:{res.status_code}"
    except: return "Upload Fail"

def create_table_image_pil(df):
    # 稍微縮減欄寬以節省記憶體
    col_widths = [60, 130, 180, 130, 200, 400, 400] 
    line_height, padding = 40, 20
    rows_data = []
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "描述"]
    rows_data.append((headers, 1, False))
    
    df = df.astype(str)
    # 限制處理行數，避免 OOM 崩潰（先取前 15 筆最嚴重的）
    if len(df) > 15: df = df.head(15)

    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        val_a = str(row.iloc[0]).strip()
        is_empty = (val_a == "" or val_a.lower() == "nan")
        char_counts = [4, 10, 8, 10, 12, 18, 18]
        for i in range(min(7, len(row))):
            text = str(row.iloc[i])
            lines = textwrap.wrap(text, width=char_counts[i])
            wrapped_row.append("\n".join(lines) if lines else "")
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines, is_empty))

    total_h = sum([m * line_height + 25 for _, m, _ in rows_data]) + (padding * 2)
    # 降低圖片品質 (RGB 改為更省的模式或限制尺寸)
    image = Image.new('RGB', (sum(col_widths) + padding * 2, int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    try:
        font = ImageFont.truetype(FONT_PATH, 24)
        h_font = ImageFont.truetype(FONT_PATH, 26)
    except: font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines, is_empty) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 25
        bg = (45, 90, 45) if r_idx == 0 else ((245, 245, 220) if is_empty else (255, 255, 255))
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg)
        for c_idx, text in enumerate(text_list):
            draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(220, 220, 220))
            draw.text((x + 10, y + 10), text, fill=tc, font=font if r_idx > 0 else h_font)
            x += col_widths[c_idx]
        y += row_h

    temp_file = f"{uuid.uuid4()}.png"
    image.save(temp_file, "PNG", optimize=True) # 優化存檔大小
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
            res = requests.get(SHEET_URL, timeout=15)
            if res.status_code != 200:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"連線失敗:{res.status_code}"))
                return
            
            df = pd.read_csv(StringIO(res.text), encoding='utf-8-sig', on_bad_lines='skip')
            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="試算表無資料"))
                return
            
            # 標題處理
            df.columns = df.iloc[0]
            df = df.drop(df.index[0])

            img_path = create_table_image_pil(df)
            img_url = upload_to_imgbb(img_path)
            
            if img_url.startswith("http"):
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"圖片傳送失敗，請稍後再試"))
            
            if os.path.exists(img_path): os.remove(img_path)
            
        except Exception as e:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="系統繁忙中，請縮減試算表行數後再試"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="機器人連線正常！輸入「總表」看資料。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"[上傳錯誤] {img_result}"))
            
            if os.path.exists(img_path): os.remove(img_path)
            
        except Exception as e:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"[系統報錯] 執行時發生問題。\n原因：{str(e)}"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"機器人連線正常！\n輸入「總表」可產生報表。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
