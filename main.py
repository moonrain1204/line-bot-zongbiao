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
SHEET_KEY = "您的試算表ID" # 從試算表網址找出：https://docs.google.com/spreadsheets/d/這串就是ID/edit

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# --- Google Sheets API 連線設定 ---
def get_sheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    # 建議將 JSON 金鑰內容放入環境變數 GOOGLE_CREDS
    creds_json = json.loads(os.environ.get('GOOGLE_CREDS'))
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
    client = gspread.authorize(creds)
    return client.open_by_key(SHEET_KEY)

def get_font():
    possible_names = ["boldfonts.ttf", "boldfonts.TTF", "Linefonts.ttf", "myfont.ttf"]
    for name in possible_names:
        path = os.path.join(os.getcwd(), name)
        if os.path.exists(path): return path
    return None

def create_table_image_pil(df, highlight_store=None):
    # 欄位與排版設定
    col_widths = [80, 160, 240, 160, 220, 580, 750] 
    line_height, padding = 55, 60 
    rows_data = [("排序", "日期", "店別", "型號", "電話", "地址", "問題與故障描述")]
    
    # 整理資料列
    for _, row in df.iterrows():
        r = []
        char_counts = [5, 12, 12, 10, 15, 22, 28]
        for i, val in enumerate(row[:7]):
            text = str(val).replace("nan", "").strip()
            if i == 0 and text: text = str(int(float(text)))
            lines = textwrap.wrap(text, width=char_counts[i]) if text else [" "]
            r.append("\n".join(lines))
        rows_data.append(("\n".join(r[0]), "\n".join(r[1]), r[2], r[3], r[4], r[5], r[6]))

    total_table_width = sum(col_widths)
    canvas_width = total_table_width + (padding * 2) + 50
    total_h = (len(rows_data) * 120) + (padding * 2) # 估算高度
    
    image = Image.new('RGB', (int(canvas_width), 3000), (255, 255, 255))
    draw = ImageDraw.Draw(image)
    font_path = get_font()
    font = ImageFont.truetype(font_path, 28) if font_path else ImageFont.load_default()

    y = padding
    for r_idx, row_text in enumerate(rows_data):
        # 計算此行最高高度
        max_lines = max([len(t.split('\n')) for t in row_text])
        row_h = max_lines * 45 + 40
        
        # 邏輯 2：完修上底色
        store_name = row_text[2].split('\n')[0]
        if r_idx == 0: bg, tc = (45, 90, 45), (255, 255, 255)
        elif store_name == highlight_store: bg, tc = (255, 220, 180), (0, 0, 0) # 橘黃色
        else: bg, tc = (255, 255, 255), (0, 0, 0)
        
        draw.rectangle([padding, y, padding + total_table_width, y + row_h], fill=bg)
        curr_x = padding
        for c_idx, text in enumerate(row_text):
            draw.rectangle([curr_x, y, curr_x + col_widths[c_idx], y + row_h], outline=(200, 200, 200))
            draw.text((curr_x + 15, y + 15), text, fill=tc, font=font, spacing=8)
            curr_x += col_widths[c_idx]
        y += row_h

    final_image = image.crop((0, 0, canvas_width, y + padding))
    temp_file = f"{uuid.uuid4()}.png"
    final_image.save(temp_file, "PNG")
    return temp_file

def reorder_and_update(wks):
    """將 A 欄重新依序編號並寫回 Google Sheets"""
    data = wks.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    # 過濾出 A 欄原本就有內容的行進行重新編號
    count = 1
    for i, row in df.iterrows():
        if row['排序'].strip() != "":
            wks.update_cell(i + 2, 1, count)
            count += 1

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    msg = event.message.text.strip()
    try:
        sh = get_sheet()
        wks_repair = sh.worksheet("永慶待修")
        wks_info = sh.worksheet("永慶安裝資訊")
        
        highlight_store = None

        # --- 邏輯 2：減少內容 (完修回報) ---
        if "完修" in msg:
            store_name = msg.split()[0]
            cell = wks_repair.find(store_name)
            if cell:
                # 只處理 A 欄有數字的（待修中的）
                if wks_repair.cell(cell.row, 1).value:
                    wks_repair.update_cell(cell.row, 1, "") # 清空排序
                    wks_repair.update_cell(cell.row, 8, datetime.now().strftime("%Y/%m/%d")) # 完修日期 (H欄)
                    wks_repair.update_cell(cell.row, 9, msg) # 完修回報紀錄 (I欄)
                    highlight_store = store_name
                    reorder_and_update(wks_repair)

        # --- 邏輯 3：增加內容 (報修回報) ---
        elif "報修" in msg:
            parts = msg.split()
            store_name = parts[0]
            issue = parts[1] if len(parts) > 1 else "報修"
            
            # 邏輯 1：檢查重複報修
            all_repair = wks_repair.get_all_records()
            is_duplicate = any(r['店別'] == store_name and str(r['排序']) != "" for r in all_repair)
            
            if not is_duplicate:
                # 從安裝資訊抓資料 (VLOOKUP)
                info_data = wks_info.get_all_records()
                store_info = next((item for item in info_data if item["店別"] == store_name), None)
                
                if store_info:
                    new_row = [
                        99, # 暫時排序
                        datetime.now().strftime("%Y/%m/%d"),
                        store_name,
                        store_info.get('型號', ''),
                        store_info.get('電話', ''),
                        store_info.get('地址', ''),
                        issue
                    ]
                    wks_repair.append_row(new_row)
                    reorder_and_update(wks_repair)

        # --- 共通輸出：總表圖片 ---
        data = wks_repair.get_all_values()
        df = pd.DataFrame(data[1:], columns=data[0])
        # 過濾 A 欄非空值
        display_df = df[df['排序'].str.strip() != ""].copy()
        
        if not display_df.empty or msg == "總表":
            img_path = create_table_image_pil(display_df, highlight_store)
            # 上傳 ImgBB (略過，同前次代碼)
            with open(img_path, "rb") as f:
                img_res = requests.post("https://api.imgbb.com/1/upload", data={"key": IMGBB_API_KEY, "image": base64.b64encode(f.read())})
                img_url = img_res.json()['data']['url']
            line_bot_api.reply_message(event.reply_token, ImageSendMessage(img_url, img_url))
            if os.path.exists(img_path): os.remove(img_path)

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=os.environ.get('PORT', 5000))
