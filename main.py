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

# 偵測絕對路徑以確保伺服器讀得到檔案
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_PATH = os.path.join(BASE_DIR, "myfont.ttf")

@app.route("/", methods=['GET'])
def index():
    return "機器人服務中"

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
    
    df = df.astype(str)
    
    # 核心繪圖邏輯：將資料轉為多行文字
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        char_counts = [4, 10, 8, 10, 12, 18, 18]
        for i in range(min(7, len(row))):
            text = str(row.iloc[i]).replace("nan", "").strip()
            # textwrap 有助於中文斷行
            lines = textwrap.wrap(text, width=char_counts[i])
            wrapped_row.append("\n".join(lines) if lines else "")
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines))

    # 計算總高度
    total_h = sum([m * line_height + 25 for _, m in rows_data]) + (padding * 2)
    image = Image.new('RGB', (sum(col_widths) + padding * 2, int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    # --- 關鍵修正：字體加載強化 ---
    try:
        if os.path.exists(FONT_PATH):
            # 加入布林參數以強化加載 (針對 Linux 環境)
            font = ImageFont.truetype(FONT_PATH, 24)
            h_font = ImageFont.truetype(FONT_PATH, 26)
        else:
            font = h_font = ImageFont.load_default()
    except Exception:
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 25
        
        # 底色設定：標題深綠，內容純白 (解決底部黃塊問題)
        bg = (45, 90, 45) if r_idx == 0 else (255, 255, 255)
        tc = (255, 255, 255) if r_idx == 0 else (0, 0, 0)
        
        draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg)
        for c_idx, text in enumerate(text_list):
            draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            # 渲染文字：加入 spacing 參數優化顯示
            draw.text((x + 10, y + 10), text, fill=tc, font=font if r_idx > 0 else h_font, spacing=8)
            x += col_widths[c_idx]
        y += row_h

    temp_file = f"{uuid.uuid4()}.png"
    image.save(temp_file, "PNG")
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
            # 抓取 Google Sheet 資料
            res = requests.get(SHEET_URL, timeout=15)
            # 使用 utf-8-sig 處理 BOM 編碼
            df = pd.read_csv(StringIO(res.text), encoding='utf-8-sig', on_bad_lines='skip', header=0)
            
            # --- 過濾 A 欄(第一欄)非空值 ---
            # 確保第一欄位名稱正確或使用 index
            df = df[df.iloc[:, 0].notna()]
            df = df[df.iloc[:, 0].astype(str).str.strip() != ""]
            # 移除常見的空字串內容
            df = df[~df.iloc[:, 0].astype(str).str.lower().isin(["nan", "none", "0", "0.0"])]

            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="目前無有效的報修資料。"))
                return

            # 安全限制：前 20 筆 (避免 Koyeb 記憶體超載)
            if len(df) > 20: df = df.head(20)

            # 產生圖片並上傳
            img_path = create_table_image_pil(df)
            img_url = upload_to_imgbb(img_path)
            
            if img_url and img_url.startswith("http"):
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="圖片上傳圖床失敗，請檢查 API Key。"))
            
            if os.path.exists(img_path): os.remove(img_path)
        except Exception:
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="讀取資料失敗，請檢查 SHEET_URL 或網路發佈狀態。"))
    else:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="連線正常！輸入「總表」產生報表。"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
