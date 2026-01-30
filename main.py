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
GOOGLE_SHEET_ID = "1O2Uy1Diw4Y01rvSFSigRHPxwslw40gvbdt93BqM4ywQ"
SHEET_GID = "596601469" 
IMGBB_API_KEY = "f65fa2212137d99c892644b1be26afac" 

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# 確保字體路徑正確，建議檢查 GitHub 上的實際檔名
FONT_PATH = "LINESeedJP-Regular.ttf" 

@app.route("/", methods=['GET'])
def index():
    return "<h1>表格優化版運行中！ (空值自動上色模式)</h1>"

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
    # --- 優化點 2: 欄寬大幅調整，避免地址與電話重疊 ---
    # 順序：排序, 日期, 店別, 型號, 電話, 地址, 問題描述
    col_widths = [80, 160, 220, 150, 250, 550, 550] 
    line_height = 45
    padding = 30
    
    rows_data = []
    # 標題列
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data.append((headers, 1, False)) 
    
    # 內容列處理
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        # --- 優化點 1: 判斷 A 欄（排序）是否為空 ---
        val_a = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        is_a_empty = (val_a == "" or val_a.lower() == "nan")
        
        # 換行字數設定：電話設窄一點強迫換行，地址設寬一點
        char_counts = [4, 10, 8, 10, 14, 22, 22] 
        for i in range(min(7, len(row))):
            text = str(row.iloc[i]) if pd.notna(row.iloc[i]) else ""
            lines = textwrap.wrap(text, width=char_counts[i])
            if not lines: lines = [""]
            wrapped_row.append("\n".join(lines))
            max_lines = max(max_lines, len(lines))
            
        # 補齊不足的欄位
        while len(wrapped_row) < 7:
            wrapped_row.append("")
            
        rows_data.append((wrapped_row, max_lines, is_a_empty))

    total_h = sum([m * line_height + 35 for _, m, _ in rows_data]) + (2 * padding)
    total_w = sum(col_widths) + (2 * padding)
    
    image = Image.new('RGB', (total_w, int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    try:
        font = ImageFont.truetype(FONT_PATH, 28)
        h_font = ImageFont.truetype(FONT_PATH, 30)
    except:
        # 字體若讀取失敗，使用內建字體（雖無中文但不會報錯）
        font = h_font = ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines, is_empty) in enumerate(rows_data):
        x = padding
        row_h = m_lines * line_height + 35
        
        # --- 優化點 1: 背景顏色邏輯：A欄為空則變色 ---
        if r_idx == 0:
            bg_color = (45, 90, 45) # 標題深綠
            text_color = (255, 255, 255)
        elif is_empty:
            bg_color = (245, 245, 220) # A欄為空：米黃色底（明顯區隔）
            text_color = (0, 0, 0)
        else:
            bg_color = (255, 255, 255) # A欄有值：純白底
            text_color = (0, 0, 0)

        draw.rectangle([x, y, x + sum(col_widths), y + row_h], fill=bg_color)
        
        for c_idx, text in enumerate(text_list):
            draw.rectangle([x, y, x + col_widths[c_idx], y + row_h], outline=(200, 200, 200), width=1)
            # 優化點 3: 增加左邊距 15px 避免文字貼邊，減少亂碼感
            draw.text((x + 15, y + 15), text, fill=text_color, font=font if r_idx > 0 else h_font, spacing=8)
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
    user_msg = event.message.text
    
    # 測試用邏輯：只要傳送「測試」或隨便打，確認機器人會動
    if user_msg == "總表":
        try:
            # --- 優化點 3: 強制 utf-8-sig 解決中文亂碼 ---
            sheet_url = f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export?format=csv&gid={SHEET_GID}"
            df = pd.read_csv(sheet_url, encoding='utf-8-sig') 
            
            # 清理第一列標題
            if df.shape[1] >= 7:
                df.columns = df.iloc[0]
                df = df.drop(df.index[0])
            
            if df.empty:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="目前試算表內沒有資料。"))
                return

            local_path = create_table_image_pil(df)
            public_url = upload_to_imgbb(local_path)
            
            if public_url:
                line_bot_api.reply_message(
                    event.reply_token,
                    ImageSendMessage(original_content_url=public_url, preview_image_url=public_url)
                )
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="圖片上傳 ImgBB 失敗，請檢查 API Key。"))
                
            if os.path.exists(local_path): os.remove(local_path)
        except Exception as e:
            print(f"錯誤日誌: {e}")
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"系統錯誤，請聯絡管理員檢查 Log。"))
    else:
        # 如果收到非「總表」的訊息，回覆這個確認連線是通的
        line_bot_api.reply_message(
            event.reply_token, 
            TextSendMessage(text=f"機器人連線正常！您輸入的是：{user_msg}\n輸入「總表」可產生報表。")
        )

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)
