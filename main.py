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
    """連線 Google Sheets，若環境變數出錯會回報錯誤訊息"""
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_raw = os.environ.get('GOOGLE_CREDS')
        if not creds_raw: return "ERR_NO_CREDS"
        creds_json = json.loads(creds_raw)
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
        return gspread.authorize(creds)
    except Exception as e:
        print(f"Creds Error: {e}")
        return str(e)

def get_font():
    """尋找字型檔"""
    possible_names = ["boldfonts.ttf", "Linefonts.ttf", "myfont.ttf"]
    for name in possible_names:
        path = os.path.join(os.getcwd(), name)
        if os.path.exists(path): return path
    return None

def create_table_image_pil(df, highlight_store=None):
    """繪製表格圖片"""
    col_widths = [80, 160, 240, 160, 220, 580, 750] 
    line_height, padding = 55, 60 
    headers = ["排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述"]
    rows_data = [(headers, 1, None)]
    
    for _, row in df.iterrows():
        wrapped_row = []
        max_lines = 1
        char_counts = [5, 12, 12, 10, 15, 22, 28] 
        for i in range(min(7, len(row))):
            val = row.iloc[i]
            text = str(val).replace("nan", "").strip()
            if i == 0 and text: 
                try: text = str(int(float(text)))
                except: pass
            lines = textwrap.wrap(text, width=char_counts[i]) if text else [" "]
            wrapped_row.append("\n".join(lines))
            max_lines = max(max_lines, len(lines))
        rows_data.append((wrapped_row, max_lines, row['店別']))

    total_table_width = sum(col_widths)
    total_h = sum([m * line_height + 45 for _, m, _ in rows_data]) + (padding * 2)
    image = Image.new('RGB', (int(total_table_width + padding * 2 + 50), int(total_h)), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    
    font_path = get_font()
    font = ImageFont.truetype(font_path, 28) if font_path else ImageFont.load_default()

    y = padding
    for r_idx, (text_list, m_lines, store_name) in enumerate(rows_data):
        row_h = m_lines * line_height + 45
        bg = (45, 90, 45) if r_idx == 0 else ((255, 220, 180) if store_name == highlight_store else (255, 255, 255))
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

def reorder_sheet(wks):
    """重新排序 A 欄"""
    vals = wks.get_all_values()
    if len(vals) < 2: return
    df = pd.DataFrame(vals[1:], columns=vals[0])
    count = 1
    updates = []
    for i, row in df.iterrows():
        if str(row.get('排序', '')).strip() != "":
            updates.append({'range': f'A{i+2}', 'values': [[count]]})
            count += 1
    if updates: wks.batch_update(updates)

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
    client = get_sheet_client()
    
    # 檢查 API 連線是否正常
    if isinstance(client, str):
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"系統連線失敗: {client}"))
        return

    try:
        sh = client.open_by_key(SHEET_KEY)
        wks_repair = sh.worksheet("永慶待修")
        wks_info = sh.worksheet("永慶安裝資訊")
        highlight_store = None

        # 1. 減少內容 (完修)
        if "完修" in msg:
            store_name = msg.split()[0]
            try:
                cell = wks_repair.find(store_name)
                if wks_repair.cell(cell.row, 1).value:
                    wks_repair.update_cell(cell.row, 1, "")
                    wks_repair.update_cell(cell.row, 8, datetime.now().strftime("%Y/%m/%d"))
                    wks_repair.update_cell(cell.row, 9, msg)
                    highlight_store = store_name
                    reorder_sheet(wks_repair)
            except: pass

        # 2. 增加內容 (報修)
        elif "報修" in msg:
            parts = msg.split()
            store_name = parts[0]
            issue = parts[1] if len(parts) > 1 else "報修"
            all_rows = wks_repair.get_all_records()
            if not any(r.get('店別') == store_name and str(r.get('排序')).strip() != "" for r in all_rows):
                info_list = wks_info.get_all_records()
                info = next((i for i in info_list if i.get('店別') == store_name), None)
                if info:
                    new_row = [99, datetime.now().strftime("%Y/%m/%d"), store_name, 
                               info.get('型號',''), info.get('電話',''), info.get('地址',''), issue]
                    wks_repair.append_row(new_row)
                    reorder_sheet(wks_repair)

        # 3. 輸出總表 (A欄非空範圍)
        if msg == "總表" or highlight_store or "報修" in msg:
            data = wks_repair.get_all_values()
            df = pd.DataFrame(data[1:], columns=data[0])
            display_df = df[df['排序'].str.strip() != ""].copy()
            
            if not display_df.empty:
                img_path = create_table_image_pil(display_df, highlight_store)
                with open(img_path, "rb") as f:
                    img_res = requests.post("https://api.imgbb.com/1/upload", 
                                            data={"key": IMGBB_API_KEY, "image": base64.b64encode(f.read())})
                    img_url = img_res.json()['data']['url']
                line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
                if os.path.exists(img_path): os.remove(img_path)
            else:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text="目前無待修項目。"))

    except Exception as e:
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"處理失敗: {e}"))

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=os.environ.get('PORT', 5000))
