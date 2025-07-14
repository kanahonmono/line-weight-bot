from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime

app = Flask(__name__)

# === 環境変数から設定読み込み ===
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_APPLICATION_CREDENTIALS_JSON")
SPREADSHEET_ID = "1mmdxzloT6rOmx7SiVT4X2PtmtcsBxivcHSoMUvjDCqc"

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

# === ユーザー情報取得 ===
def get_user_info_by_id(user_id):
    try:
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:E").execute()
        for row in result.get("values", []):
            if row[4] == user_id:
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

# === 空き列を自動で見つける ===
def find_next_available_columns():
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!B1:Z1").execute()
    header = result.get('values', [[]])[0]
    for i in range(1, 25, 2):
        if (len(header) <= i or header[i] == '') and (len(header) <= i+1 or header[i+1] == ''):
            return chr(ord('A') + i), chr(ord('A') + i + 1)
    raise Exception("空き列がありません")

# === ユーザー登録 ===
def register_user(username, mode, user_id):
    user_info = get_user_info_by_id(user_id)
    if user_info:
        return "すでに登録済みです。"
    weight_col, mode_col = find_next_available_columns()
    sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="Users!A:E",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": [[username, mode, weight_col, mode_col, user_id]]}
    ).execute()
    sheet.values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": [
            {"range": f"Weights!{weight_col}1", "values": [[f"{username}体重"]]},
            {"range": f"Weights!{mode_col}1", "values": [[f"{username}モード"]]}
        ]}
    ).execute()
    return f"{username} さんを登録しました！"

# === ユーザー削除 ===
def reset_user(user_id):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Users!A2:E").execute()
    values = result.get("values", [])
    for i, row in enumerate(values):
        if len(row) >= 5 and row[4] == user_id:
            sheet.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"Users!A{i+2}:E{i+2}",
                valueInputOption="USER_ENTERED",
                body={"values": [["" for _ in range(5)]]}
            ).execute()
            return "登録をリセットしました。"
    return "ユーザーが見つかりませんでした。"

# === 体重記録 ===
def append_weight_data(user_id, weight, date=None):
    user_info = get_user_info_by_id(user_id)
    if user_info is None:
        raise Exception(f"ユーザー情報が見つかりません: {user_id}")

    weight_col = user_info["weight_col"]
    mode_col = user_info["mode_col"]
    weights_sheet = "Weights"

    if date is None:
        date = datetime.now().strftime('%Y-%m-%d')

    # 1行目のA列に日付一覧を取得
    header_range = f"{weights_sheet}!A1:1"
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=header_range).execute()
    header = result.get("values", [[]])[0]

    # 日付が既に存在しているか確認
    if date in header:
        col_index = header.index(date) + 1  # 1-based index
    else:
        col_index = len(header) + 1
        col_letter = chr(ord('A') + col_index - 1)
        # 日付を1行目に追加
        sheet.values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{weights_sheet}!{col_letter}1",
            valueInputOption="USER_ENTERED",
            body={"values": [[date]]}
        ).execute()

    # 対応する列のアルファベットを求める
    col_letter = chr(ord('A') + col_index - 1)

    # ✅ 「1行目」に体重・モードを書き込む
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{weights_sheet}!{col_letter}1",
        valueInputOption="USER_ENTERED",
        body={"values": [[date]]}
    ).execute()

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{weights_sheet}!{col_letter}2",
        valueInputOption="USER_ENTERED",
        body={"values": [[weight]]}
    ).execute()

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{weights_sheet}!{col_letter}3",
        valueInputOption="USER_ENTERED",
        body={"values": [[user_info['mode']]]}
    ).execute()

    return f"{user['username']} さんの体重 {weight}kg を記録しました！"

# === LINEコールバック ===
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
            reply = "こんにちは！\n■体重記録コマンド\n体重 65.5\n体重 2025-07-13 65.5\n登録 ユーザー名 モード\nリセット"

        elif parts[0] == "登録" and len(parts) == 3:
            reply = register_user(parts[1], parts[2], user_id)

        elif parts[0] == "リセット":
            reply = reset_user(user_id)

        elif parts[0] == "体重":
            if len(parts) == 2:
                weight = float(parts[1])
                reply = append_weight(user_id, weight)
            elif len(parts) == 3:
                date = parts[1]
                weight = float(parts[2])
                reply = append_weight(user_id, weight, date)
            else:
                reply = "体重コマンドの形式が正しくありません。"
        else:
            reply = "コマンドが正しくありません。ヘルプと送ってください。"
    except Exception as e:
        reply = f"エラーが発生しました: {e}"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

@app.route("/", methods=["GET"])
def home():
    return "LINEダイエットBot起動中"

if __name__ == "__main__":
    app.run(debug=True)
