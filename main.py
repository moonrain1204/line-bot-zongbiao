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

app = Flask(__name__)

# --- 設定區 ---
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')
GOOGLE_SHEET_ID = os.environ.get('GOOGLE_SHEET_ID')
SHEET_GID = os.environ.get('SHEET_GID')
IMGBB_API_KEY = "f65fa2212137d99c892644b1be26afac" 

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
FONT_PATH = "myfont.ttf" 

@app.route("/", methods=['GET'])
def index():
    return "<h1>機器人連線中！</h1>"

def upload_to_imgbb(image_path):
    try:
        with open(image_path, "rb") as file:
            url = "https://api.imgbb.com/1/upload"
            payload = {
                "key": IMGBB_API_KEY,
                "image": base64.b64encode(file.read()),
            }
            res = requests.post(url, data=payload, timeout=20) # 增加超時等待
            if res.status_code == 200:
                return res.json()['data']['url']
            return f"ImgBB Error: {res.status_code} - {res.text[:50]}"
    except Exception as e:
        return f"Upload Exception: {str(e)}"

def create_table_image_pil(df):
    col_widths = [80, 160, 220, 150, 250, 550, 550] 
    line_height, padding = 45, 30
    rows_data = []
    # 標題
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1, False)) 
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        val_a = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        is_empty = (val_a == "" or val_a.lower() == "nan")
        char_counts = [4, 10, 8, 10, 14, 22, 22] 
        for i in range(min(7, len(row))):
            text = str(row.iloc[i]) if pd.notna(row.iloc[i]) else ""
            lines = textwrap.wrap(text, width=char_counts[i])
            wrapped_row.append("\n".join(lines) if lines else "")
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines, is_empty))

    total_h = sum([m * line_height + 35 for _, m, _ in rows_data]) + (padding * 2)
    image = Image.new('RGB', (sum(col_widths) + padding * 2, int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    try:
        font = ImageFont.truetype(FONT_PATH, 28)
        h_font = ImageFont.truetype(FONT_PATH, 30)
    except:
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines, is_empty) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 35
        if r_idx == 0:
            bg_color, text_color = (45, 90, 45), (255, 255, 255)
        elif is_empty:
            bg_color, text_color = (245, 245, 220), (0, 0, 0)
        else:
            bg_color, text_color = (255, 255, 255), (0, 0, 0)

        draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg_color)
        for c_idx, text in enumerate(text_list):
            draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            draw.text((x + 15, y + 15), text, fill=text_color, font=font if r_idx > 0 else h_font, spacing=8)
            x += col_widths[c_idx]
        y += row_h

    temp_file = f"table_{uuid.uuid4()}.png"
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
    msg = event.message.text
    if msg == "總表":
        try:
            # 1. 讀取資料
            sheet_url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv&gid={SHEET_GID}"
            df = pd.read_csv(sheet_url, encoding='utf-8-sig') 
            
            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="讀取成功但表內沒有資料。"))
                return

            # 2. 產圖
            img_path = create_table_image_pil(df)
            
            # 3. 上傳
            img_result = upload_to_imgbb(img_path)
            
            if img_result.startswith("http"):
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_result, img_result))
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"圖片上傳錯誤：\n{img_result}"))
            
            if os.path.exists(img_path): os.remove(img_path)
        except Exception as e:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"處理失敗，錯誤原因：\n{str(e)}"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"連線正常！輸入的是：{msg}\n輸入「總表」可產生報表。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
