import threading
import webview
from flask import Flask, render_template, request
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

import os, sys

def resource_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)



app = Flask(__name__)

# ========================================
# 🔹 Googleスプレッドシートの読み込み
# ========================================
def get_sheet():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(
        resource_path("credentials.json"),
        scopes=scope
    )
    client = gspread.authorize(creds)

    # ▼スプレッドシートID（固定）
    sheet_id =""

    # ▼タブ名を「受付_2025_11_7」の形式で自動生成
    now = datetime.now()
    sheet_name = f"受付_{now.year}_{now.month}_{now.day}"

    spreadsheet = client.open_by_key(sheet_id)
    sheet = spreadsheet.worksheet(sheet_name)

    return sheet


# ========================================
# 🔹 Flask ページ
# ========================================
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/search', methods=['POST'])
def search():
    name_input = request.form['name']
    address_input = request.form['address'].strip()
    phone = request.form['phone']

    sheet = get_sheet()
    data = sheet.get_all_values()

    result = []
    for row in data:
        if len(row) >= 2:
            name = row[0].strip()
            address = row[1].strip()
            if address == address_input:   # 完全一致
                result.append({"名前": name, "住所": address})

    return render_template(
        'result.html',
        name=name_input,
        address=address_input,
        phone=phone,
        result=result
    )


# ========================================
# 🔹 Flask を別スレッドで起動
# ========================================
def start_flask():
    app.run(host="127.0.0.1", port=5000, debug=False)


# ========================================
# 🔹 PyWebView 起動
# ========================================
if __name__ == '__main__':
    # Flask をバックグラウンドで動かす
    threading.Thread(target=start_flask, daemon=True).start()

    # WebView ウィンドウを開く
    webview.create_window(
        title="にげてきまっし",
        url="http://127.0.0.1:5000",
        width=450,
        height=700
    )
    webview.start()
