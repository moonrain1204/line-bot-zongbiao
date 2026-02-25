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
# 注意：這裡的 SHEET_URL 必須是 CSV 導出連結
# 格式為：https://docs.google.com/spreadsheets/d/ID/export?format=csv&gid=分頁GID
SHEET_URL = os.environ.get('SHEET_URL') 
IMGBB_API_KEY = os.environ.get('IMGBB_API_KEY') or "f65fa2212137d99c892644b1be26afac"

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

def get_font():
    """尋找可用字型"""
    possible_names = ["boldfonts.ttf", "Linefonts.ttf", "myfont.ttf", "simhei.ttf"]
    for name in possible_names:
        path = os.path.join(os.getcwd(), name)
        if os.path.exists(path): return path
    return None

def create_table_image_pil(df):
    """繪製總表圖片"""
    # 欄位寬度設定
    col_widths = [80, 160, 240, 160, 220, 580, 750] 
    line_height, padding = 55, 60 
    rows_data = []
    
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1))
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        # 設定換行寬度
        char_counts = [5, 12, 10, 10, 15, 22, 28] 
        for i in range(min(7, len(row))):
            val = row.iloc[i]
            # 處理排序欄位(A欄)
            if i == 0:
                try:
                    text = str(int(float(val))) if pd.notna(val) and str(val).strip() != "" else ""
                except:
                    text = str(val).strip()
            else:
                text = str(val).replace("nan", "").strip()
            
            text = text.replace('\r', '').replace('\u3000', ' ').replace('\xa0', ' ')
            lines = textwrap.wrap(text, width=char_counts[i]) if text else [" "]
            wrapped_row.append("\n".join(lines))
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    total_table_width = sum(col_widths)
    canvas_width = total_table_width + (padding * 2) + 50 
    total_h = sum([m * line_height + 40 for _, m in rows_data]) + (padding * 2)
    
    # 防止圖片過短
    total_h = max(total_h, 300)
    
    image = Image.new('RGB', (int(canvas_width), int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    font_path = get_font()
    try:
        font = ImageFont.truetype(font_path, 28) if font_path else ImageFont.load_default()
    except:
        font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 40
        # 標題深綠色，資料白色
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        
        draw.rectangle([x, y, x + total_table_width, y + row_h], fill=bg)
        curr_x = x
        for c_idx, text in enumerate(text_list):
            draw.rectangle([curr_x, y, curr_x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            draw.text((curr_x + 15, y + 15), text, fill=tc, font=font, spacing=8)
            curr_x += col_widths[c_idx]
        y += row_h

    temp_file = f"{uuid.uuid4()}.png"
    image.save(temp_file, "PNG")
    return temp_file

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_msg = event.message.text.strip()
    
    if user_msg == "總表":
        try:
            # 1. 從 Google Sheets 抓取資料
            res = requests.get(SHEET_URL, timeout=15)
            res.encoding = 'utf-8-sig'
            df = pd.read_csv(StringIO(res.text))
            
            # 2. 核心修正：正確篩選 A 欄(第一欄) 非空的資料
            # 將第一欄轉換成字串並去除空白，過濾掉長度為 0 的資料
            df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.replace('nan', '').str.strip()
            df = df[df.iloc[:, 0] != ""]
            
            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="目前無待修項目。"))
                return

            # 3. 繪圖
            img_path = create_table_image_pil(df.head(30))
            
            # 4. 上傳到 ImgBB
            with open(img_path, "rb") as f:
                img_data = base64.b64encode(f.read())
                img_res = requests.post("https://api.imgbb.com/1/upload", 
                                        data={"key": IMGBB_API_KEY, "image": img_data})
                
                if img_res.status_code == 200:
                    img_url = img_res.json()['data']['url']
                    # 5. 回傳圖片訊息
                    line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
                else:
                    line_bot_api.reply_message(event.reply_token, TextSendMessage(text="圖片上傳失敗，請檢查 ImgBB API Key。"))
            
            if os.path.exists(img_path): os.remove(img_path)
            
        except Exception as e:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"總表呼叫失敗: {str(e)}"))

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=os.environ.get('PORT', 5000))
