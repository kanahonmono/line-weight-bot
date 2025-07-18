from flask import Flask, request, abort, send_file
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, ImageSendMessage
from slugify import slugify
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta
import pandas as pd
import matplotlib.pyplot as plt

app = Flask(__name__)

# === 環境変数 ===
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
SPREADSHEET_ID = "1mmdxzloT6rOmx7SiVT4X2PtmtcsBxivcHSoMUvjDCqc"
YOUR_PUBLIC_BASE_URL = os.getenv("YOUR_PUBLIC_BASE_URL")

if not (LINE_CHANNEL_SECRET and LINE_CHANNEL_ACCESS_TOKEN and GOOGLE_CREDENTIALS):
    raise Exception("環境変数が足りません")

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

credentials_dict = json.loads(GOOGLE_CREDENTIALS)
credentials = service_account.Credentials.from_service_account_info(
    credentials_dict,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
sheet_service = build('sheets', 'v4', credentials=credentials)
sheet = sheet_service.spreadsheets()

# === ユーザー取得（user_id） ===
def get_user_info_by_id(user_id):
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:E").execute()
        for row in result.get("values", []):
            if len(row) >= 5 and row[4] == user_id:
                return {
                    "username": row[0],
                    "mode": row[1],
                    "weight_col": row[2],
                    "mode_col": row[3],
                    "user_id": row[4],
                }
        return None
    except Exception as e:
        print(f"ユーザー情報取得エラー: {e}")
        return None

# === ユーザー取得（username） ===
def get_user_info_by_username(username):
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:E").execute()
        for row in result.get("values", []):
            if len(row) >= 1 and row[0] == username:
                return {
                    "username": row[0],
                    "mode": row[1] if len(row) > 1 else "",
                    "weight_col": row[2] if len(row) > 2 else "",
                    "mode_col": row[3] if len(row) > 3 else "",
                    "user_id": row[4] if len(row) > 4 else "",
                }
        return None
    except Exception as e:
        print(f"ユーザー名による情報取得エラー: {e}")
        return None

# === 登録・削除・記録 ===
def register_user(username, mode, user_id):
    if get_user_info_by_id(user_id):
        return "すでに登録済みです。"
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!B1:Z1").execute()
    header = result.get('values', [[]])[0]
    for i in range(1, 25, 2):
        if (len(header) <= i or header[i] == '') and (len(header) <= i+1 or header[i+1] == ''):
            weight_col = chr(ord('A') + i)
            mode_col = chr(ord('A') + i + 1)
            break
    else:
        raise Exception("空き列がありません")

    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="Users!A:E",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [[username, mode, weight_col, mode_col, user_id]]}
    ).execute()
    return f"{username} さんを登録しました！"

def reset_user(user_id):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:E").execute()
    for i, row in enumerate(result.get("values", [])):
        if len(row) >= 5 and row[4] == user_id:
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"Users!A{i+2}:E{i+2}",
                valueInputOption="USER_ENTERED",
                body={"values": [["" for _ in range(5)]]}
            ).execute()
            return "登録をリセットしました。"
    return "ユーザーが見つかりませんでした。"

def append_vertical_weight(user_info, date, weight):
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="Weights!A:D",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [[user_info["username"], date, weight, user_info["mode"]]]}
    ).execute()

# === グラフ処理 ===
def get_last_month_weight_data(username):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Weights!A2:D").execute()
    df = pd.DataFrame(result.get("values", []), columns=["ユーザー名", "日付", "体重", "モード"])
    df = df[df["ユーザー名"] == username]
    if df.empty:
        return None
    df['日付'] = pd.to_datetime(df['日付'])
    return df[df['日付'] >= datetime.now() - timedelta(days=30)].sort_values('日付')

def create_monthly_weight_graph(df, username):
    df['日付'] = pd.to_datetime(df['日付'])
    plt.figure(figsize=(8, 4))
    plt.plot(df['日付'], df['体重'].astype(float), marker='o', linestyle='-', color='blue')
    plt.title(f"{username} さんの直近1か月の体重推移")
    plt.xlabel("日付")
    plt.ylabel("体重 (kg)")
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    static_dir = os.path.join(app.root_path, "static", "graphs")
    os.makedirs(static_dir, exist_ok=True)
    safe_username = slugify(username)
    path = os.path.join(static_dir, f"{safe_username}_weight_1month.png")
    plt.savefig(path)
    plt.close()
    print(f"グラフ画像を保存しました: {path}")
    return path

def send_monthly_weight_graph_to_line(user_info):
    df = get_last_month_weight_data(user_info['username'])
    if df is None or df.empty:
        raise Exception("直近1か月の体重データが見つかりません。")
    local_path = create_monthly_weight_graph(df, user_info['username'])
    filename = os.path.basename(local_path)
    if not YOUR_PUBLIC_BASE_URL:
        raise Exception("YOUR_PUBLIC_BASE_URL が設定されていません")
    img_url = f"{YOUR_PUBLIC_BASE_URL}/static/graphs/{filename}"
    line_bot_api.push_message(user_info['user_id'], ImageSendMessage(
        original_content_url=img_url,
        preview_image_url=img_url
    ))

# === LINE webhook ===
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

# === メッセージ処理 ===
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    text = event.message.text.strip()
    user_id = event.source.user_id
    parts = text.split()
    try:
        if text.lower() == "ヘルプ":
            reply = (
                "こんにちは！\n"
                "■体重記録コマンド\n"
                "体重 65.5\n"
                "体重 2025-07-13 65.5\n"
                "登録 ユーザー名 モード\n"
                "リセット\n"
                "グラフ送信\n"
                "グラフ ユーザー名"
            )
        elif parts[0] == "登録" and len(parts) == 3:
            reply = register_user(parts[1], parts[2], user_id)
        elif parts[0] == "リセット":
            reply = reset_user(user_id)
        elif parts[0] == "体重":
            user_info = get_user_info_by_id(user_id)
            if not user_info:
                reply = "登録されていません。まず登録してください。"
            elif len(parts) == 2:
                append_vertical_weight(user_info, datetime.now().strftime('%Y-%m-%d'), float(parts[1]))
                reply = f"{user_info['username']} さんの体重 {parts[1]}kg を記録しました！"
            elif len(parts) == 3:
                append_vertical_weight(user_info, parts[1], float(parts[2]))
                reply = f"{user_info['username']} さんの体重 {parts[2]}kg（{parts[1]}）を記録しました！"
            else:
                reply = "体重コマンドの形式が正しくありません。"
        elif text.lower() == "グラフ送信":
            user_info = get_user_info_by_id(user_id)
            if not user_info:
                reply = "登録されていません。先に登録してください。"
            else:
                send_monthly_weight_graph_to_line(user_info)
                return
        elif parts[0] == "グラフ" and len(parts) == 2:
            user_info = get_user_info_by_username(parts[1])
            if not user_info:
                reply = f"{parts[1]} さんは未登録です。"
            else:
                df = get_last_month_weight_data(parts[1])
                if df is None or df.empty:
                    reply = "データがありません。"
                else:
                    local_path = create_monthly_weight_graph(df, parts[1])
                    safe_username = slugify(parts[1])
                    filename = f"{safe_username}_weight_1month.png"
                    img_url = f"{YOUR_PUBLIC_BASE_URL}/static/graphs/{filename}"
                    line_bot_api.reply_message(event.reply_token, ImageSendMessage(
                        original_content_url=img_url,
                        preview_image_url=img_url
                    ))
                    return
        else:
            reply = "コマンドが正しくありません。ヘルプと送ってください。"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))
    except Exception as e:
        print(f"エラー: {e}")
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"エラーが発生しました: {e}"))
@app.route("/list_graphs")
def list_graphs():
    static_dir = os.path.join(app.root_path, "static", "graphs")
    try:
        files = os.listdir(static_dir)
    except Exception as e:
        return f"エラー: {e}"
    return "<br>".join(files)
# === 画像配信用 ===
@app.route("/static/graphs/<filename>")
def serve_image(filename):
    return send_file(os.path.join(app.root_path, "static", "graphs", filename), mimetype="image/png")
