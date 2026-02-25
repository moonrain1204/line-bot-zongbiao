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

def get_font():
    """自動尋找當前目錄下可用的中文延伸字型檔"""
    # 優先搜尋您截圖中顯示的 boldfonts.ttf
    possible_names = ["boldfonts.ttf", "boldfonts.TTF", "Linefonts.ttf", "myfont.ttf"]
    for name in possible_names:
        path = os.path.join(os.getcwd(), name)
        if os.path.exists(path):
            print(f"--- 成功找到指定字型檔: {name} ---")
            return path
    
    # 如果指定名稱都沒找到，掃描目錄下所有字型檔案作為後備
    for file in os.listdir('.'):
        if file.lower().endswith(".ttf") or file.lower().endswith(".otf"):
            print(f"--- 自動偵測到可用字型: {file} ---")
            return os.path.join(os.getcwd(), file)
    return None

def create_table_image_pil(df):
    # 欄位寬度設定（根據地址與描述內容加寬）
    col_widths = [80, 160, 240, 160, 220, 550, 700] 
    line_height, padding = 50, 60 
    rows_data = []
    
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1))
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        char_counts = [5, 12, 12, 10, 15, 25, 35] 
        for i in range(min(7, len(row))):
            val = row.iloc[i]
            # 處理排序數字與空值
            text = str(val).replace("nan", "").strip() if i != 0 else str(int(float(val))) if pd.notna(val) else ""
            # 清理可能導致渲染錯誤的特殊字元
            text = text.replace('\r', '').replace('\u3000', ' ').replace('\xa0', ' ')
            lines = textwrap.wrap(text, width=char_counts[i]) if text else [" "]
            wrapped_row.append("\n".join(lines))
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    total_table_width = sum(col_widths)
    canvas_width = total_table_width + (padding * 2) + 100 
    total_h = sum([m * line_height + 40 for _, m in rows_data]) + (padding * 2)
    
    image = Image.new('RGB', (int(canvas_width), int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    # 執行字型載入
    font_path = get_font()
    try:
        if font_path:
            # 使用較大字級 (30/32) 確保手機瀏覽清晰
            font = ImageFont.truetype(font_path, 30)
            h_font = ImageFont.truetype(font_path, 32)
        else:
            print("--- 警告: 找不到字型檔，使用 Default Font 將導致亂碼 ---")
            font = h_font = ImageFont.load_default()
    except Exception as e:
        print(f"字型加載錯誤: {e}")
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 40
        # 設定標題列顏色 (深綠色) 與文字顏色
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        
        draw.rectangle([x, y, x + total_table_width, y + row_h], fill=bg)
        curr_x = x
        for c_idx, text in enumerate(text_list):
            draw.rectangle([curr_x, y, curr_x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            draw.text((curr_x + 15, y + 15), text, fill=tc, font=font, spacing=10)
            curr_x += col_widths[c_idx]
        y += row_h

    temp_file = f"{uuid.uuid4()}.png"
    image.save(temp_file, "PNG")
    return temp_file

@app.route("/", methods=['GET'])
def index(): 
    return "總表機器人運行中 - 目前字型設定: boldfonts.ttf"

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
    if event.message.text.strip() == "總表":
        try:
            res = requests.get(SHEET_URL, timeout=15)
            res.encoding = 'utf-8-sig'
            df = pd.read_csv(StringIO(res.text)).head(20)
            
            img_path = create_table_image_pil(df)
            
            # 上傳至 ImgBB
            with open(img_path, "rb") as f:
                img_res = requests.post("https://api.imgbb.com/1/upload", 
                                        data={"key": IMGBB_API_KEY, "image": base64.b64encode(f.read())})
                img_url = img_res.json()['data']['url']
            
            line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            
            # 刪除臨時檔案
            if os.path.exists(img_path): 
                os.remove(img_path)
        except Exception as e:
            print(f"處理失敗: {e}")
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="讀取失敗，請確認日誌訊息。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=os.environ.get('PORT', 5000))
